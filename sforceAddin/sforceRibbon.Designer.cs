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
            this.dropDown_org = this.Factory.CreateRibbonDropDown();
            this.loadTable_btn = this.Factory.CreateRibbonButton();
            this.btn_taskPane = this.Factory.CreateRibbonButton();
            this.grp_data = this.Factory.CreateRibbonGroup();
            this.btn_load = this.Factory.CreateRibbonButton();
            this.btn_CommitChanges = this.Factory.CreateRibbonButton();
            this.btn_CopySelection = this.Factory.CreateRibbonButton();
            this.grp_config = this.Factory.CreateRibbonGroup();
            this.editbox_APIVersion = this.Factory.CreateRibbonEditBox();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.gallery_addOrg = this.Factory.CreateRibbonGallery();
            this.sfRibbonTab.SuspendLayout();
            this.grp_auth.SuspendLayout();
            this.grp_data.SuspendLayout();
            this.grp_config.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.SuspendLayout();
            // 
            // sfRibbonTab
            // 
            this.sfRibbonTab.Groups.Add(this.grp_auth);
            this.sfRibbonTab.Groups.Add(this.grp_data);
            this.sfRibbonTab.Groups.Add(this.grp_config);
            this.sfRibbonTab.Label = "sforce";
            this.sfRibbonTab.Name = "sfRibbonTab";
            // 
            // grp_auth
            // 
            this.grp_auth.Items.Add(this.dropDown_org);
            this.grp_auth.Items.Add(this.loadTable_btn);
            this.grp_auth.Items.Add(this.btn_taskPane);
            this.grp_auth.Label = "Auth";
            this.grp_auth.Name = "grp_auth";
            // 
            // dropDown_org
            // 
            this.dropDown_org.Label = "Org";
            this.dropDown_org.Name = "dropDown_org";
            this.dropDown_org.SizeString = "MMMMMMMMMMM";
            this.dropDown_org.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown_org_SelectionChanged);
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
            this.grp_data.Items.Add(this.btn_CopySelection);
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
            this.btn_CommitChanges.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Commit_Changes);
            // 
            // btn_CopySelection
            // 
            this.btn_CopySelection.Label = "Copy Selection";
            this.btn_CopySelection.Name = "btn_CopySelection";
            this.btn_CopySelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_CopySelection_Click);
            // 
            // grp_config
            // 
            this.grp_config.Items.Add(this.editbox_APIVersion);
            this.grp_config.Items.Add(this.buttonGroup1);
            this.grp_config.Label = "Config";
            this.grp_config.Name = "grp_config";
            // 
            // editbox_APIVersion
            // 
            this.editbox_APIVersion.Label = "API Version:";
            this.editbox_APIVersion.Name = "editbox_APIVersion";
            this.editbox_APIVersion.Text = "48.0";
            this.editbox_APIVersion.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.apiVersion_TextChanged);
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.gallery_addOrg);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // gallery_addOrg
            // 
            ribbonDropDownItemImpl1.Label = "Sandbox";
            ribbonDropDownItemImpl2.Label = "Production";
            this.gallery_addOrg.Items.Add(ribbonDropDownItemImpl1);
            this.gallery_addOrg.Items.Add(ribbonDropDownItemImpl2);
            this.gallery_addOrg.Label = "+Auth an Org";
            this.gallery_addOrg.Name = "gallery_addOrg";
            this.gallery_addOrg.RowCount = 2;
            this.gallery_addOrg.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gallery_login_Click);
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
            this.grp_config.ResumeLayout(false);
            this.grp_config.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab sfRibbonTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_auth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_taskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_data;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_load;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_CommitChanges;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loadTable_btn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_config;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editbox_APIVersion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallery_addOrg;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_org;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_CopySelection;
    }

    partial class ThisRibbonCollection
    {
        internal sforceRibbon sforceRibbon
        {
            get { return this.GetRibbon<sforceRibbon>(); }
        }
    }
}
