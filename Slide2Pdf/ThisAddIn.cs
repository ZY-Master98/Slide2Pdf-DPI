using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;

namespace Slide2Pdf
{
    public struct Rect
    {
        public double Top { get; set; }
        public double Left { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Bottom { get { return Top + Height; } }
        public double Right { get { return Left + Width; } }
        public Rect(double top, double left, double width, double height)
        {
            Top = top;
            Left = left;
            Width = width;
            Height = height;
        }
    }

    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void ExportCurrentSlideAsPdf(string outPath)
        {
            var presentation = Application.ActivePresentation;
            presentation.ExportAsFixedFormat(
                outPath,
                PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                Intent: PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                RangeType: PowerPoint.PpPrintRangeType.ppPrintCurrent
            );
        }

        public void ExportCurrentSlideAsImage(string outPath, string powerPointFilter, int dpi, Rect? cropRect = null)
        {
            using (Bitmap image = RenderCurrentSlideImage(dpi, cropRect))
            {
                SaveImage(image, outPath, powerPointFilter);
            }
        }

        public void CopyCurrentSlideAsImage(int dpi, Rect? cropRect = null)
        {
            using (Bitmap image = RenderCurrentSlideImage(dpi, cropRect))
            {
                // Explicitly publish only the standard Bitmap clipboard format.
                // Passing an Image directly makes Windows Forms auto-serialize its
                // MemoryBmp raw format, which can fail with a generic GDI+ error.
                var clipboardData = new System.Windows.Forms.DataObject();
                clipboardData.SetData(System.Windows.Forms.DataFormats.Bitmap, true, image);
                System.Windows.Forms.Clipboard.SetDataObject(clipboardData, true, 5, 100);
            }
        }

        private Bitmap RenderCurrentSlideImage(int dpi, Rect? cropRect)
        {
            var slide = Application.ActiveWindow?.View?.Slide as PowerPoint.Slide;
            if (slide == null)
            {
                throw new InvalidOperationException("No active slide is selected.");
            }

            float slideWidthPoints = Application.ActivePresentation.PageSetup.SlideWidth;
            float slideHeightPoints = Application.ActivePresentation.PageSetup.SlideHeight;
            int width = Math.Max(1, (int)Math.Round(slideWidthPoints / 72.0 * dpi));
            int height = Math.Max(1, (int)Math.Round(slideHeightPoints / 72.0 * dpi));

            if (width > 16384 || height > 16384)
            {
                throw new InvalidOperationException(
                    $"The selected DPI produces a {width} x {height} pixel image. " +
                    "Reduce the DPI so neither dimension exceeds 16384 pixels.");
            }

            string temporaryPngPath = Path.Combine(Path.GetTempPath(), $"Slide2Pdf_{Guid.NewGuid():N}.png");
            try
            {
                slide.Export(temporaryPngPath, "PNG", width, height);
                using (var source = LoadTemporaryBitmap(temporaryPngPath))
                {
                    Rectangle pixelBounds = cropRect.HasValue
                        ? GetPixelCropBounds(cropRect.Value, source.Width, source.Height)
                        : new Rectangle(0, 0, source.Width, source.Height);
                    Bitmap rendered = source.Clone(pixelBounds, PixelFormat.Format32bppArgb);
                    try
                    {
                        rendered.SetResolution(dpi, dpi);
                        return rendered;
                    }
                    catch
                    {
                        rendered.Dispose();
                        throw;
                    }
                }
            }
            finally
            {
                DeleteTemporaryFileBestEffort(temporaryPngPath);
            }
        }

        private static Bitmap LoadTemporaryBitmap(string path)
        {
            Exception lastException = null;
            for (int attempt = 1; attempt <= 10; attempt++)
            {
                try
                {
                    // PowerPoint can keep the exported PNG open briefly after Slide.Export
                    // returns. Shared access plus a detached clone avoids retaining that lock.
                    using (var stream = new FileStream(
                        path,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite | FileShare.Delete))
                    using (var temporary = new Bitmap(stream))
                    {
                        return temporary.Clone(
                            new Rectangle(0, 0, temporary.Width, temporary.Height),
                            PixelFormat.Format32bppArgb);
                    }
                }
                catch (IOException ex)
                {
                    lastException = ex;
                }
                catch (UnauthorizedAccessException ex)
                {
                    lastException = ex;
                }
                catch (ArgumentException ex)
                {
                    // The image may still be incomplete while PowerPoint is flushing it.
                    lastException = ex;
                }

                if (attempt < 10)
                {
                    Thread.Sleep(50 * attempt);
                }
            }

            throw new IOException("PowerPoint did not finish writing the temporary slide image.", lastException);
        }

        private static void DeleteTemporaryFileBestEffort(string path)
        {
            for (int attempt = 1; attempt <= 10; attempt++)
            {
                try
                {
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                    return;
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }

                if (attempt < 10)
                {
                    Thread.Sleep(50 * attempt);
                }
            }

            // A stale temp file is preferable to failing an otherwise successful copy.
        }

        private static Rectangle GetPixelCropBounds(Rect cropRect, int imageWidth, int imageHeight)
        {
            int left = Math.Max(0, (int)Math.Floor(cropRect.Left * imageWidth));
            int top = Math.Max(0, (int)Math.Floor(cropRect.Top * imageHeight));
            int right = Math.Min(imageWidth, (int)Math.Ceiling(cropRect.Right * imageWidth));
            int bottom = Math.Min(imageHeight, (int)Math.Ceiling(cropRect.Bottom * imageHeight));
            if (right <= left || bottom <= top)
            {
                throw new InvalidOperationException("The visible content bounds are empty after cropping.");
            }
            return Rectangle.FromLTRB(left, top, right, bottom);
        }

        private static void SaveImage(Bitmap image, string outPath, string powerPointFilter)
        {
            switch ((powerPointFilter ?? string.Empty).ToUpperInvariant())
            {
                case "JPG":
                case "JPEG":
                    ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders().First(codec => codec.FormatID == ImageFormat.Jpeg.Guid);
                    using (var encoderParameters = new EncoderParameters(1))
                    {
                        encoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 95L);
                        image.Save(outPath, jpegCodec, encoderParameters);
                    }
                    break;
                case "TIF":
                case "TIFF":
                    image.Save(outPath, ImageFormat.Tiff);
                    break;
                case "BMP":
                    image.Save(outPath, ImageFormat.Bmp);
                    break;
                case "GIF":
                    image.Save(outPath, ImageFormat.Gif);
                    break;
                default:
                    image.Save(outPath, ImageFormat.Png);
                    break;
            }
        }

        public bool GetCurrentSlideContentBoundingRect(out Rect rect)
        {
            var slide = Application.ActiveWindow.View.Slide as PowerPoint.Slide;
            double minX = double.MaxValue;
            double minY = double.MaxValue;
            double maxX = double.MinValue;
            double maxY = double.MinValue;
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (shape.Visible == Office.MsoTriState.msoTrue)
                {
                    minX = Math.Min(minX, shape.Left);
                    minY = Math.Min(minY, shape.Top);
                    maxX = Math.Max(maxX, shape.Left + shape.Width);
                    maxY = Math.Max(maxY, shape.Top + shape.Height);
                }
            }
            if (minX == double.MaxValue || minY == double.MaxValue || maxX == double.MinValue || maxY == double.MinValue)
            {
                rect = new Rect(0, 0, 0, 0);
                return false;
            }
            // Return the bounding rectangle in relative coordinates
            double slideWidth = slide.Master.Width;
            double slideHeight = slide.Master.Height;
            minX = Math.Max(minX, 0);
            minY = Math.Max(minY, 0);
            maxX = Math.Min(maxX, slideWidth);
            maxY = Math.Min(maxY, slideHeight);
            rect = new Rect(
                minY / slideHeight,
                minX / slideWidth,
                (maxX - minX) / slideWidth,
                (maxY - minY) / slideHeight
            );
            return true;
        }

        public void CropPdf(string pdfPath, Rect rect)
        {
            using (var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify))
            {
                foreach (var page in document.Pages)
                {
                    double pdfWidth = page.Width.Point;
                    double pdfHeight = page.Height.Point;
                    var pdfRect = new PdfRectangle(
                        new XPoint(rect.Left * pdfWidth, (1 - rect.Bottom) * pdfHeight),
                        new XPoint(rect.Right * pdfWidth, (1 - rect.Top) * pdfHeight)
                    );
                    page.TrimBox = pdfRect;
                    page.CropBox = page.TrimBox;
                }
                string tempPath = System.IO.Path.GetTempFileName();
                document.Save(tempPath);
                document.Close();
                System.IO.File.Delete(pdfPath);
                System.IO.File.Move(tempPath, pdfPath);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
