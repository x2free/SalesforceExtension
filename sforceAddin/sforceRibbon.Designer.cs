namespace sforceAddin
{
    partial class sforceRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public sforceRibbon()
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
            this.sfRibbonTab = this.Factory.CreateRibbonTab();
            this.grp_auth = this.Factory.CreateRibbonGroup();
            this.btn_login = this.Factory.CreateRibbonButton();
            this.btn_taskPane = this.Factory.CreateRibbonButton();
            this.grp_data = this.Factory.CreateRibbonGroup();
            this.btn_load = this.Factory.CreateRibbonButton();
            this.btn_upsert = this.Factory.CreateRibbonButton();
            this.sfRibbonTab.SuspendLayout();
            this.grp_auth.SuspendLayout();
            this.grp_data.SuspendLayout();
            this.SuspendLayout();
            // 
            // sfRibbonTab
            // 
            this.sfRibbonTab.Groups.Add(this.grp_auth);
            this.sfRibbonTab.Groups.Add(this.grp_data);
            this.sfRibbonTab.Label = "sforce";
            this.sfRibbonTab.Name = "sfRibbonTab";
            // 
            // grp_auth
            // 
            this.grp_auth.Items.Add(this.btn_login);
            this.grp_auth.Items.Add(this.btn_taskPane);
            this.grp_auth.Label = "Auth";
            this.grp_auth.Name = "grp_auth";
            // 
            // btn_login
            // 
            this.btn_login.Label = "login";
            this.btn_login.Name = "btn_login";
            this.btn_login.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_login_Click);
            // 
            // btn_taskPane
            // 
            this.btn_taskPane.Label = "Show/Hide Task Pane";
            this.btn_taskPane.Name = "btn_taskPane";
            this.btn_taskPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_taskPane_Click);
            // 
            // grp_data
            // 
            this.grp_data.Items.Add(this.btn_load);
            this.grp_data.Items.Add(this.btn_upsert);
            this.grp_data.Label = "Data";
            this.grp_data.Name = "grp_data";
            // 
            // btn_load
            // 
            this.btn_load.Label = "Load Data";
            this.btn_load.Name = "btn_load";
            this.btn_load.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_load_Click);
            // 
            // btn_upsert
            // 
            this.btn_upsert.Label = "Upsert Data";
            this.btn_upsert.Name = "btn_upsert";
            // 
            // sforceRibbon
            // 
            this.Name = "sforceRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.sfRibbonTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.sforceRibbon_Load);
            this.sfRibbonTab.ResumeLayout(false);
            this.sfRibbonTab.PerformLayout();
            this.grp_auth.ResumeLayout(false);
            this.grp_auth.PerformLayout();
            this.grp_data.ResumeLayout(false);
            this.grp_data.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab sfRibbonTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_auth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_login;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_taskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_data;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_load;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_upsert;
    }

    partial class ThisRibbonCollection
    {
        internal sforceRibbon sforceRibbon
        {
            get { return this.GetRibbon<sforceRibbon>(); }
        }
    }
}
