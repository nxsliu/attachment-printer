namespace AttachmentPrinter
{
    partial class AttachmentPrinterRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AttachmentPrinterRibbon()
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
            this.AttachmentPrinterTab = this.Factory.CreateRibbonTab();
            this.PrintGroup = this.Factory.CreateRibbonGroup();
            this.EmailNumberEditBox = this.Factory.CreateRibbonEditBox();
            this.PrintButton = this.Factory.CreateRibbonButton();
            this.AttachmentPrinterTab.SuspendLayout();
            this.PrintGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // AttachmentPrinterTab
            // 
            this.AttachmentPrinterTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.AttachmentPrinterTab.Groups.Add(this.PrintGroup);
            this.AttachmentPrinterTab.Label = "Attachment Printer";
            this.AttachmentPrinterTab.Name = "AttachmentPrinterTab";
            // 
            // PrintGroup
            // 
            this.PrintGroup.Items.Add(this.EmailNumberEditBox);
            this.PrintGroup.Items.Add(this.PrintButton);
            this.PrintGroup.Label = "Print";
            this.PrintGroup.Name = "PrintGroup";
            // 
            // EmailNumberEditBox
            // 
            this.EmailNumberEditBox.Label = "Scan Unread Emails";
            this.EmailNumberEditBox.Name = "EmailNumberEditBox";
            this.EmailNumberEditBox.Text = "50";
            // 
            // PrintButton
            // 
            this.PrintButton.Label = "Print Attachments";
            this.PrintButton.Name = "PrintButton";
            this.PrintButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PrintButton_Click);
            // 
            // AttachmentPrinterRibbon
            // 
            this.Name = "AttachmentPrinterRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.AttachmentPrinterTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AttachmentPrinterRibbon_Load);
            this.AttachmentPrinterTab.ResumeLayout(false);
            this.AttachmentPrinterTab.PerformLayout();
            this.PrintGroup.ResumeLayout(false);
            this.PrintGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab AttachmentPrinterTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup PrintGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EmailNumberEditBox;
        private Microsoft.Office.Tools.Ribbon.RibbonButton PrintButton;
    }

    partial class ThisRibbonCollection
    {
        internal AttachmentPrinterRibbon AttachmentPrinterRibbon
        {
            get { return this.GetRibbon<AttachmentPrinterRibbon>(); }
        }
    }
}
