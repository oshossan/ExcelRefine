namespace ExcelRefineAddIn
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.excelRefineTab = this.Factory.CreateRibbonTab();
            this.csvFileOutputGrp = this.Factory.CreateRibbonGroup();
            this.charsetDrd = this.Factory.CreateRibbonDropDown();
            this.newLineDrd = this.Factory.CreateRibbonDropDown();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.saveAsCsvBtn = this.Factory.CreateRibbonButton();
            this.saveAsTsvBtn = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.chooseFolderTbt = this.Factory.CreateRibbonToggleButton();
            this.saveToBookFolderTbt = this.Factory.CreateRibbonToggleButton();
            this.excelRefineTab.SuspendLayout();
            this.csvFileOutputGrp.SuspendLayout();
            this.SuspendLayout();
            // 
            // excelRefineTab
            // 
            this.excelRefineTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.excelRefineTab.Groups.Add(this.csvFileOutputGrp);
            this.excelRefineTab.Label = "Excel Refine";
            this.excelRefineTab.Name = "excelRefineTab";
            // 
            // csvFileOutputGrp
            // 
            this.csvFileOutputGrp.Items.Add(this.charsetDrd);
            this.csvFileOutputGrp.Items.Add(this.newLineDrd);
            this.csvFileOutputGrp.Items.Add(this.separator1);
            this.csvFileOutputGrp.Items.Add(this.saveToBookFolderTbt);
            this.csvFileOutputGrp.Items.Add(this.chooseFolderTbt);
            this.csvFileOutputGrp.Items.Add(this.separator2);
            this.csvFileOutputGrp.Items.Add(this.saveAsCsvBtn);
            this.csvFileOutputGrp.Items.Add(this.saveAsTsvBtn);
            this.csvFileOutputGrp.Label = "CSV file output";
            this.csvFileOutputGrp.Name = "csvFileOutputGrp";
            // 
            // charsetDrd
            // 
            this.charsetDrd.Label = "Charset";
            this.charsetDrd.Name = "charsetDrd";
            this.charsetDrd.ShowLabel = false;
            // 
            // newLineDrd
            // 
            this.newLineDrd.Label = "dropDown1";
            this.newLineDrd.Name = "newLineDrd";
            this.newLineDrd.ShowLabel = false;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // saveAsCsvBtn
            // 
            this.saveAsCsvBtn.Label = "Save As CSV";
            this.saveAsCsvBtn.Name = "saveAsCsvBtn";
            this.saveAsCsvBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveAsCsvBtn_Click);
            // 
            // saveAsTsvBtn
            // 
            this.saveAsTsvBtn.Label = "Save As TSV";
            this.saveAsTsvBtn.Name = "saveAsTsvBtn";
            this.saveAsTsvBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveAsTsvBtn_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // chooseFolderTbt
            // 
            this.chooseFolderTbt.Label = "Choose folder to save";
            this.chooseFolderTbt.Name = "chooseFolderTbt";
            this.chooseFolderTbt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chooseFolderTbt_Click);
            // 
            // saveToBookFolderTbt
            // 
            this.saveToBookFolderTbt.Checked = true;
            this.saveToBookFolderTbt.Label = "Save to book\'s folder";
            this.saveToBookFolderTbt.Name = "saveToBookFolderTbt";
            this.saveToBookFolderTbt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveToBookFolderTbt_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.excelRefineTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.excelRefineTab.ResumeLayout(false);
            this.excelRefineTab.PerformLayout();
            this.csvFileOutputGrp.ResumeLayout(false);
            this.csvFileOutputGrp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab excelRefineTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup csvFileOutputGrp;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown charsetDrd;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveAsCsvBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveAsTsvBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown newLineDrd;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton chooseFolderTbt;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton saveToBookFolderTbt;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
