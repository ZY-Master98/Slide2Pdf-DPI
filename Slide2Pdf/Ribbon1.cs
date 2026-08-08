using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms; // Required for Control.ModifierKeys, Keys, SaveFileDialog, DialogResult, MessageBox

namespace Slide2Pdf
{
    public partial class Ribbon1
    {
        private string currentPresentationFullName = string.Empty;
        private readonly Dictionary<int, string> slideSavePaths = new Dictionary<int, string>();
        private readonly Dictionary<int, string> slideImageSavePaths = new Dictionary<int, string>();

        private sealed class ImageFormatOption
        {
            public string Label { get; }
            public string PowerPointFilter { get; }
            public string Extension { get; }
            public string SaveDialogFilter { get; }

            public ImageFormatOption(string label, string powerPointFilter, string extension)
            {
                Label = label;
                PowerPointFilter = powerPointFilter;
                Extension = extension;
                SaveDialogFilter = $"{label} image (*.{extension})|*.{extension}";
            }
        }

        private static readonly ImageFormatOption[] ImageFormats =
        {
            new ImageFormatOption("PNG", "PNG", "png"),
            new ImageFormatOption("JPEG", "JPG", "jpg"),
            new ImageFormatOption("TIFF", "TIF", "tif"),
            new ImageFormatOption("BMP", "BMP", "bmp"),
            new ImageFormatOption("GIF", "GIF", "gif"),
        };

        // This method is assumed to be in ThisAddIn.cs or a similar helper class
        // public void ExportCurrentSlideAsPdf(string filePath) { /* ... */ }
        // public bool GetCurrentSlideContentBoundingRect(out Rect rect) { /* ... */ }
        // public void CropPdf(string pdfPath, Rect cropRect) { /* ... */ }
        // public struct Rect { public float X1, Y1, X2, Y2; } // Define if not already defined

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // Initialization code, if any
        }

        /// <summary>
        /// Checks the current presentation's state. If it's new, unsaved, or changed,
        /// it resets the stored paths.
        /// </summary>
        /// <returns>True if paths can be managed (presentation is saved and identified), false otherwise.</returns>
        private bool UpdatePresentationContextAndPathStatus()
        {
            var app = Globals.ThisAddIn.Application;
            if (app?.ActivePresentation == null)
            {
                currentPresentationFullName = string.Empty;
                slideSavePaths.Clear();
                slideImageSavePaths.Clear();
                return false;
            }

            Presentation presentation = app.ActivePresentation;

            // If the presentation isn't saved, we can't reliably track its path or slide IDs persistently.
            if (presentation.Saved == Microsoft.Office.Core.MsoTriState.msoFalse || string.IsNullOrEmpty(presentation.FullName))
            {
                // Clear paths if the presentation becomes unsaved, as its identity is now ambiguous
                if (!string.IsNullOrEmpty(currentPresentationFullName))
                {
                    currentPresentationFullName = string.Empty;
                    slideSavePaths.Clear();
                    slideImageSavePaths.Clear();
                }
                return false;
            }

            // If the presentation's FullName has changed (e.g., saved to a new file, or different presentation opened)
            if (presentation.FullName != currentPresentationFullName)
            {
                currentPresentationFullName = presentation.FullName;
                slideSavePaths.Clear();
                slideImageSavePaths.Clear();
            }
            return true;
        }

        /// <summary>
        /// Tries to get the last saved export path for the currently active slide.
        /// </summary>
        /// <returns>The path if found, otherwise null.</returns>
        private string GetSavedPathForCurrentSlide()
        {
            if (!UpdatePresentationContextAndPathStatus())
            {
                return null;
            }

            Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow?.View?.Slide as Slide;
            if (currentSlide != null && slideSavePaths.TryGetValue(currentSlide.SlideID, out string savedPath))
            {
                return savedPath;
            }
            return null;
        }

        /// <summary>
        /// Stores the export path for the currently active slide.
        /// </summary>
        /// <param name="path">The path to save.</param>
        private void StorePathForCurrentSlide(string path)
        {
            if (!UpdatePresentationContextAndPathStatus() || string.IsNullOrEmpty(path))
            {
                return;
            }

            Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow?.View?.Slide as Slide;
            if (currentSlide != null)
            {
                slideSavePaths[currentSlide.SlideID] = path;
            }
        }

        private string GetSavedImagePathForCurrentSlide()
        {
            if (!UpdatePresentationContextAndPathStatus())
            {
                return null;
            }

            Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow?.View?.Slide as Slide;
            if (currentSlide != null && slideImageSavePaths.TryGetValue(currentSlide.SlideID, out string savedPath))
            {
                return savedPath;
            }
            return null;
        }

        private void StoreImagePathForCurrentSlide(string path)
        {
            if (!UpdatePresentationContextAndPathStatus() || string.IsNullOrEmpty(path))
            {
                return;
            }

            Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow?.View?.Slide as Slide;
            if (currentSlide != null)
            {
                slideImageSavePaths[currentSlide.SlideID] = path;
            }
        }

        private ImageFormatOption GetSelectedImageFormat()
        {
            string selectedLabel = imageFormatDropDown.SelectedItem?.Label ?? "PNG";
            foreach (ImageFormatOption format in ImageFormats)
            {
                if (string.Equals(format.Label, selectedLabel, StringComparison.OrdinalIgnoreCase))
                {
                    return format;
                }
            }
            return ImageFormats[0];
        }

        private bool TryGetImageDpi(out int dpi)
        {
            dpi = 0;
            string value = imageDpiComboBox.Text ?? string.Empty;
            Match match = Regex.Match(value, @"^\s*(\d+)\s*(?:dpi)?\s*$", RegexOptions.IgnoreCase);
            if (!match.Success ||
                !int.TryParse(match.Groups[1].Value, out dpi) ||
                dpi < 36 || dpi > 1200)
            {
                MessageBox.Show(
                    "Enter a DPI value from 36 to 1200, for example 300 DPI.",
                    "Invalid Image DPI",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        private bool ExportCurrentSlideImageToFile(bool cropToContent, bool forceNewPathSelection, out string exportedImagePath)
        {
            exportedImagePath = null;
            var addIn = Globals.ThisAddIn;
            Presentation presentation = addIn.Application.ActivePresentation;
            Slide currentSlide = addIn.Application.ActiveWindow?.View?.Slide as Slide;

            if (presentation == null || currentSlide == null)
            {
                MessageBox.Show("Open a presentation and select a slide, then try again.", "Nothing to Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!TryGetImageDpi(out int dpi))
            {
                return false;
            }

            Rect cropRect = new Rect();
            if (cropToContent && !addIn.GetCurrentSlideContentBoundingRect(out cropRect))
            {
                MessageBox.Show("This slide has no visible content to crop. Export the full slide, or add visible content and try again.", "Nothing to Crop", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            ImageFormatOption format = GetSelectedImageFormat();
            string targetPath = forceNewPathSelection ? null : GetSavedImagePathForCurrentSlide();
            if (!string.IsNullOrEmpty(targetPath))
            {
                targetPath = Path.ChangeExtension(targetPath, format.Extension);
            }

            if (string.IsNullOrEmpty(targetPath))
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = format.SaveDialogFilter;
                    saveFileDialog.Title = cropToContent ? "Save Cropped Slide as Image" : "Save Current Slide as Image";
                    saveFileDialog.DefaultExt = format.Extension;
                    saveFileDialog.AddExtension = true;
                    string pptName = Path.GetFileNameWithoutExtension(presentation.Name);
                    string suffix = cropToContent ? "_cropped" : string.Empty;
                    saveFileDialog.FileName = $"{pptName}_Slide{currentSlide.SlideIndex}{suffix}.{format.Extension}";

                    if (presentation.Saved == Microsoft.Office.Core.MsoTriState.msoTrue &&
                        !string.IsNullOrEmpty(presentation.Path) &&
                        !presentation.Path.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                    {
                        saveFileDialog.InitialDirectory = presentation.Path;
                    }

                    if (saveFileDialog.ShowDialog() != DialogResult.OK)
                    {
                        return false;
                    }

                    targetPath = Path.ChangeExtension(saveFileDialog.FileName, format.Extension);
                    StoreImagePathForCurrentSlide(targetPath);
                }
            }

            try
            {
                addIn.ExportCurrentSlideAsImage(targetPath, format.PowerPointFilter, dpi,
                    cropToContent ? (Rect?)cropRect : null);
                exportedImagePath = targetPath;
                StoreImagePathForCurrentSlide(targetPath);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Couldn't export this slide as an image.\n\n{ex.Message}", "Image Export Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// Handles the core logic of exporting the current slide to PDF.
        /// It prompts for a new path if Shift is pressed or no path is remembered.
        /// </summary>
        /// <param name="forceNewPathSelection">True to force the Save File Dialog, e.g., when Shift is pressed.</param>
        /// <param name="exportedPdfPath">The path where the PDF was saved.</param>
        /// <returns>True if export was successful, false otherwise.</returns>
        private bool ExportCurrentSlideToFile(bool forceNewPathSelection, out string exportedPdfPath)
        {
            exportedPdfPath = null;
            var addIn = Globals.ThisAddIn;
            Presentation presentation = addIn.Application.ActivePresentation;
            Slide currentSlide = addIn.Application.ActiveWindow?.View?.Slide as Slide;

            if (presentation == null || currentSlide == null)
            {
                MessageBox.Show("Open a presentation and select a slide, then try again.", "Nothing to Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            string targetPath = forceNewPathSelection ? null : GetSavedPathForCurrentSlide();

            if (string.IsNullOrEmpty(targetPath))
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "PDF files (*.pdf)|*.pdf";
                    saveFileDialog.Title = "Save Current Slide as PDF";
                    saveFileDialog.DefaultExt = "pdf";

                    string pptName = Path.GetFileNameWithoutExtension(presentation.Name);
                    // Suggest a filename like "PresentationName_Slide1.pdf"
                    saveFileDialog.FileName = $"{pptName}_Slide{currentSlide.SlideIndex}.pdf";

                    // Set initial directory if presentation is saved locally
                    if (presentation.Saved == Microsoft.Office.Core.MsoTriState.msoTrue &&
                        !string.IsNullOrEmpty(presentation.Path) && // Path is directory
                        !presentation.Path.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                    {
                        saveFileDialog.InitialDirectory = presentation.Path;
                    }

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        targetPath = saveFileDialog.FileName;
                        StorePathForCurrentSlide(targetPath);
                    }
                    else
                    {
                        return false; // User cancelled
                    }
                }
            }

            if (string.IsNullOrEmpty(targetPath)) // Should not happen if dialog wasn't cancelled
            {
                return false;
            }

            try
            {
                addIn.ExportCurrentSlideAsPdf(targetPath); 
                exportedPdfPath = targetPath;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Couldn't export this slide.\n\n{ex.Message}", "Export Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Optionally, clear the stored path if export failed, to force re-selection next time.
                // if (currentSlide != null) slideSavePaths.Remove(currentSlide.SlideID);
                return false;
            }
        }

        private void btnExportSlideToPdf_Click(object sender, RibbonControlEventArgs e)
        {
            bool forceNewPath = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            if (ExportCurrentSlideToFile(forceNewPath, out string outputPath))
            {
                MessageBox.Show("PDF exported to:\n" + outputPath, "PDF Exported", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            // If ExportCurrentSlideToFile returns false, it (or its sub-methods) should have already shown an error or handled cancellation.
        }

        private void btnExportContent_Click(object sender, RibbonControlEventArgs e)
        {
            var addIn = Globals.ThisAddIn; // Assuming Rect is defined, e.g. public struct Rect { public float X1, Y1, X2, Y2; }

            if (!addIn.GetCurrentSlideContentBoundingRect(out Rect rect)) // Replace var with your actual Rect type
            {
                MessageBox.Show("This slide has no visible content to crop. Use Export Full Slide, or add visible content and try again.", "Nothing to Crop", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            bool forceNewPath = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            if (ExportCurrentSlideToFile(forceNewPath, out string outputPath))
            {
                try
                {
                    addIn.CropPdf(outputPath, rect);
                    MessageBox.Show("Cropped PDF exported to:\n" + outputPath, "PDF Exported", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"The full slide was exported, but it couldn't be cropped.\n\nFile: {outputPath}\nError: {ex.Message}", "Couldn't Crop PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnExportSlideImage_Click(object sender, RibbonControlEventArgs e)
        {
            bool forceNewPath = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            if (ExportCurrentSlideImageToFile(false, forceNewPath, out string outputPath))
            {
                MessageBox.Show("Image exported to:\n" + outputPath, "Image Exported", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnExportContentImage_Click(object sender, RibbonControlEventArgs e)
        {
            bool forceNewPath = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            if (ExportCurrentSlideImageToFile(true, forceNewPath, out string outputPath))
            {
                MessageBox.Show("Cropped image exported to:\n" + outputPath, "Image Exported", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
