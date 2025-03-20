namespace CompanyDataAddIn
{
    partial class CompanyDataRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CompanyDataRibbon()
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
            this.tabAuditing = this.Factory.CreateRibbonTab();
            this.groupCompany = this.Factory.CreateRibbonGroup();
            this.CinTextBox = this.Factory.CreateRibbonEditBox();
            this.FetchDataButton = this.Factory.CreateRibbonButton();
            this.tabAuditing.SuspendLayout();
            this.groupCompany.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabAuditing
            // 
            this.tabAuditing.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabAuditing.Groups.Add(this.groupCompany);
            this.tabAuditing.Label = "Auditing";
            this.tabAuditing.Name = "tabAuditing";
            // 
            // groupCompany
            // 
            this.groupCompany.Items.Add(this.CinTextBox);
            this.groupCompany.Items.Add(this.FetchDataButton);
            this.groupCompany.Label = "Company";
            this.groupCompany.Name = "groupCompany";
            // 
            // CinTextBox
            // 
            this.CinTextBox.Label = "IČO:";
            this.CinTextBox.Name = "CinTextBox";
            this.CinTextBox.Text = "36206075";
            // 
            // FetchDataButton
            // 
            this.FetchDataButton.Label = "Fetch Data";
            this.FetchDataButton.Name = "FetchDataButton";
            this.FetchDataButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FetchDataButton_Click);
            // 
            // CompanyDataRibbon
            // 
            this.Name = "CompanyDataRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabAuditing);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CompanyDataRibbon_Load);
            this.tabAuditing.ResumeLayout(false);
            this.tabAuditing.PerformLayout();
            this.groupCompany.ResumeLayout(false);
            this.groupCompany.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabAuditing;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCompany;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox CinTextBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FetchDataButton;
    }

    partial class ThisRibbonCollection
    {
        internal CompanyDataRibbon CompanyDataRibbon
        {
            get { return this.GetRibbon<CompanyDataRibbon>(); }
        }
    }
}
