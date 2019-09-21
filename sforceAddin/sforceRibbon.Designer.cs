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
            this.grp_login = this.Factory.CreateRibbonGroup();
            this.btn_login = this.Factory.CreateRibbonButton();
            this.sfRibbonTab.SuspendLayout();
            this.grp_login.SuspendLayout();
            this.SuspendLayout();
            // 
            // sfRibbonTab
            // 
            this.sfRibbonTab.Groups.Add(this.grp_login);
            this.sfRibbonTab.Label = "sforce";
            this.sfRibbonTab.Name = "sfRibbonTab";
            // 
            // grp_login
            // 
            this.grp_login.Items.Add(this.btn_login);
            this.grp_login.Label = "Login";
            this.grp_login.Name = "grp_login";
            // 
            // btn_login
            // 
            this.btn_login.Label = "login";
            this.btn_login.Name = "btn_login";
            this.btn_login.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_login_Click);
            // 
            // sforceRibbon
            // 
            this.Name = "sforceRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.sfRibbonTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.sforceRibbon_Load);
            this.sfRibbonTab.ResumeLayout(false);
            this.sfRibbonTab.PerformLayout();
            this.grp_login.ResumeLayout(false);
            this.grp_login.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab sfRibbonTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_login;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_login;
    }

    partial class ThisRibbonCollection
    {
        internal sforceRibbon sforceRibbon
        {
            get { return this.GetRibbon<sforceRibbon>(); }
        }
    }
}
