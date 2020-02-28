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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.sfRibbonTab = this.Factory.CreateRibbonTab();
            this.grp_auth = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btn_login = this.Factory.CreateRibbonButton();
            this.orgType_cb = this.Factory.CreateRibbonComboBox();
            this.loadTable_btn = this.Factory.CreateRibbonButton();
            this.btn_taskPane = this.Factory.CreateRibbonButton();
            this.grp_data = this.Factory.CreateRibbonGroup();
            this.btn_load = this.Factory.CreateRibbonButton();
            this.btn_CommitChanges = this.Factory.CreateRibbonButton();
            this.sfRibbonTab.SuspendLayout();
            this.grp_auth.SuspendLayout();
            this.box1.SuspendLayout();
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
            this.grp_auth.Items.Add(this.box1);
            this.grp_auth.Items.Add(this.loadTable_btn);
            this.grp_auth.Items.Add(this.btn_taskPane);
            this.grp_auth.Label = "Auth";
            this.grp_auth.Name = "grp_auth";
            // 
            // box1
            // 
            this.box1.Items.Add(this.btn_login);
            this.box1.Items.Add(this.orgType_cb);
            this.box1.Name = "box1";
            // 
            // btn_login
            // 
            this.btn_login.Label = "login";
            this.btn_login.Name = "btn_login";
            this.btn_login.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_login_Click);
            // 
            // orgType_cb
            // 
            ribbonDropDownItemImpl1.Label = "Sandbox";
            ribbonDropDownItemImpl2.Label = "Production";
            this.orgType_cb.Items.Add(ribbonDropDownItemImpl1);
            this.orgType_cb.Items.Add(ribbonDropDownItemImpl2);
            this.orgType_cb.Label = " ";
            this.orgType_cb.Name = "orgType_cb";
            this.orgType_cb.ScreenTip = "Org type";
            this.orgType_cb.Text = null;
            this.orgType_cb.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.orgType_cb_TextChanged);
            // 
            // loadTable_btn
            // 
            this.loadTable_btn.Label = "Load Tables";
            this.loadTable_btn.Name = "loadTable_btn";
            this.loadTable_btn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.loadTable_btn_Click);
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
            this.grp_data.Items.Add(this.btn_CommitChanges);
            this.grp_data.Label = "Data";
            this.grp_data.Name = "grp_data";
            // 
            // btn_load
            // 
            this.btn_load.Label = "Load Data";
            this.btn_load.Name = "btn_load";
            this.btn_load.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_load_Click);
            // 
            // btn_CommitChanges
            // 
            this.btn_CommitChanges.Label = "Commit Changes";
            this.btn_CommitChanges.Name = "btn_CommitChanges";
            this.btn_CommitChanges.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_upsert_Click);
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
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_CommitChanges;
        public Microsoft.Office.Tools.Ribbon.RibbonComboBox orgType_cb;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loadTable_btn;
    }

    partial class ThisRibbonCollection
    {
        internal sforceRibbon sforceRibbon
        {
            get { return this.GetRibbon<sforceRibbon>(); }
        }
    }
}
