
namespace MrpImport_Clicks.Excel.AddIn
{
    partial class ForecastImport : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ForecastImport()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ForecastImport));
            this.tabImportMrp = this.Factory.CreateRibbonTab();
            this.grpUpdate = this.Factory.CreateRibbonGroup();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnUpdate = this.Factory.CreateRibbonButton();
            this.btnComment = this.Factory.CreateRibbonButton();
            this.grpVersion = this.Factory.CreateRibbonGroup();
            this.btnVer = this.Factory.CreateRibbonButton();
            this.tabImportMrp.SuspendLayout();
            this.grpUpdate.SuspendLayout();
            this.group1.SuspendLayout();
            this.grpVersion.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabImportMrp
            // 
            this.tabImportMrp.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabImportMrp.Groups.Add(this.grpUpdate);
            this.tabImportMrp.Groups.Add(this.group1);
            this.tabImportMrp.Groups.Add(this.grpVersion);
            this.tabImportMrp.Label = "Clicks Mrp";
            this.tabImportMrp.Name = "tabImportMrp";
            // 
            // grpUpdate
            // 
            this.grpUpdate.Items.Add(this.btnUpdate);
            this.grpUpdate.Label = "Import Forecast";
            this.grpUpdate.Name = "grpUpdate";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnComment);
            this.group1.Label = "Instructions";
            this.group1.Name = "group1";
            // 
            // btnUpdate
            // 
            this.btnUpdate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdate.Image = ((System.Drawing.Image)(resources.GetObject("btnUpdate.Image")));
            this.btnUpdate.Label = "UPDATE MRP FORECAST";
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.ShowImage = true;
            this.btnUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdate_Click);
            // 
            // btnComment
            // 
            this.btnComment.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnComment.Image = ((System.Drawing.Image)(resources.GetObject("btnComment.Image")));
            this.btnComment.Label = "Select the Detailed Tab and press the Update button.";
            this.btnComment.Name = "btnComment";
            this.btnComment.ShowImage = true;
            // 
            // grpVersion
            // 
            this.grpVersion.Items.Add(this.btnVer);
            this.grpVersion.Label = "Version";
            this.grpVersion.Name = "grpVersion";
            // 
            // btnVer
            // 
            this.btnVer.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
            this.btnVer.Label = " Version Number";
            this.btnVer.Name = "btnVer";
            this.btnVer.ShowImage = true;
            // 
            // ForecastImport
            // 
            this.Name = "ForecastImport";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabImportMrp);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabImportMrp.ResumeLayout(false);
            this.tabImportMrp.PerformLayout();
            this.grpUpdate.ResumeLayout(false);
            this.grpUpdate.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.grpVersion.ResumeLayout(false);
            this.grpVersion.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabImportMrp;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnComment;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpVersion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVer;
    }

    partial class ThisRibbonCollection
    {
        internal ForecastImport Ribbon1
        {
            get { return this.GetRibbon<ForecastImport>(); }
        }
    }
}
