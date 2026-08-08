namespace Slide2Pdf
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnExportFullSlide = this.Factory.CreateRibbonButton();
            this.btnExportContent = this.Factory.CreateRibbonButton();
            this.imageExportGroup = this.Factory.CreateRibbonGroup();
            this.imageDpiComboBox = this.Factory.CreateRibbonComboBox();
            this.imageFormatDropDown = this.Factory.CreateRibbonDropDown();
            this.btnExportSlideImage = this.Factory.CreateRibbonButton();
            this.btnExportContentImage = this.Factory.CreateRibbonButton();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem dpi96 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem dpi150 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem dpi300 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem dpi600 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem formatPng = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem formatJpeg = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem formatTiff = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem formatBmp = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem formatGif = this.Factory.CreateRibbonDropDownItem();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.imageExportGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.imageExportGroup);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnExportFullSlide);
            this.group1.Items.Add(this.btnExportContent);
            this.group1.Label = "Export PDF";
            this.group1.Name = "group1";
            // 
            // btnExportFullSlide
            // 
            this.btnExportFullSlide.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportFullSlide.Image = global::Slide2Pdf.Properties.Resources.document_pdf_512x512;
            this.btnExportFullSlide.Label = "Export Full Slide";
            this.btnExportFullSlide.Name = "btnExportFullSlide";
            this.btnExportFullSlide.ScreenTip = "Export the current slide at full size";
            this.btnExportFullSlide.ShowImage = true;
            this.btnExportFullSlide.SuperTip = "Slide2Pdf remembers this slide's export location. Hold Shift while clicking " +
    "to choose a different location.";
            this.btnExportFullSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportSlideToPdf_Click);
            // 
            // btnExportContent
            // 
            this.btnExportContent.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportContent.Image = global::Slide2Pdf.Properties.Resources.crop_512x512;
            this.btnExportContent.Label = "Crop to Content";
            this.btnExportContent.Name = "btnExportContent";
            this.btnExportContent.ScreenTip = "Export the current slide, cropped to visible content";
            this.btnExportContent.ShowImage = true;
            this.btnExportContent.SuperTip = "Slide2Pdf remembers this slide's export location. Hold Shift while clicking " +
    "to choose a different location.";
            this.btnExportContent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportContent_Click);
            //
            // imageExportGroup
            //
            this.imageExportGroup.Items.Add(this.imageDpiComboBox);
            this.imageExportGroup.Items.Add(this.imageFormatDropDown);
            this.imageExportGroup.Items.Add(this.btnExportSlideImage);
            this.imageExportGroup.Items.Add(this.btnExportContentImage);
            this.imageExportGroup.Label = "Export Image";
            this.imageExportGroup.Name = "imageExportGroup";
            //
            // imageDpiComboBox
            //
            dpi96.Label = "96 DPI";
            dpi150.Label = "150 DPI";
            dpi300.Label = "300 DPI";
            dpi600.Label = "600 DPI";
            this.imageDpiComboBox.Items.Add(dpi96);
            this.imageDpiComboBox.Items.Add(dpi150);
            this.imageDpiComboBox.Items.Add(dpi300);
            this.imageDpiComboBox.Items.Add(dpi600);
            this.imageDpiComboBox.Label = "DPI";
            this.imageDpiComboBox.Name = "imageDpiComboBox";
            this.imageDpiComboBox.ScreenTip = "Image resolution in dots per inch";
            this.imageDpiComboBox.SizeString = "0000 DPI";
            this.imageDpiComboBox.Text = "300 DPI";
            this.imageDpiComboBox.SuperTip = "Choose a preset or enter a custom value from 36 to 1200 DPI. Pixel dimensions are calculated from the slide size.";
            //
            // imageFormatDropDown
            //
            formatPng.Label = "PNG";
            formatJpeg.Label = "JPEG";
            formatTiff.Label = "TIFF";
            formatBmp.Label = "BMP";
            formatGif.Label = "GIF";
            this.imageFormatDropDown.Items.Add(formatPng);
            this.imageFormatDropDown.Items.Add(formatJpeg);
            this.imageFormatDropDown.Items.Add(formatTiff);
            this.imageFormatDropDown.Items.Add(formatBmp);
            this.imageFormatDropDown.Items.Add(formatGif);
            this.imageFormatDropDown.Label = "Format";
            this.imageFormatDropDown.Name = "imageFormatDropDown";
            this.imageFormatDropDown.ScreenTip = "Output image format";
            //
            // btnExportSlideImage
            //
            this.btnExportSlideImage.Image = global::Slide2Pdf.Properties.Resources.document_pdf_512x512;
            this.btnExportSlideImage.Label = "Full Slide Image";
            this.btnExportSlideImage.Name = "btnExportSlideImage";
            this.btnExportSlideImage.ScreenTip = "Export the current slide as an image";
            this.btnExportSlideImage.ShowImage = true;
            this.btnExportSlideImage.SuperTip = "Uses the selected DPI and image format. Hold Shift while clicking to choose a different location.";
            this.btnExportSlideImage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportSlideImage_Click);
            //
            // btnExportContentImage
            //
            this.btnExportContentImage.Image = global::Slide2Pdf.Properties.Resources.crop_512x512;
            this.btnExportContentImage.Label = "Cropped Image";
            this.btnExportContentImage.Name = "btnExportContentImage";
            this.btnExportContentImage.ScreenTip = "Export visible slide content as an image";
            this.btnExportContentImage.ShowImage = true;
            this.btnExportContentImage.SuperTip = "Uses the selected DPI and image format, then crops to visible content. Hold Shift while clicking to choose a different location.";
            this.btnExportContentImage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportContentImage_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.imageExportGroup.ResumeLayout(false);
            this.imageExportGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportFullSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportContent;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup imageExportGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox imageDpiComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown imageFormatDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportSlideImage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportContentImage;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
