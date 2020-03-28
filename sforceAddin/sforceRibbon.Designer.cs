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
            this.grp_config = this.Factory.CreateRibbonGroup();
            this.dropDown_TargetOrg = this.Factory.CreateRibbonDropDown();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.gallery_AuthOrg = this.Factory.CreateRibbonGallery();
            this.btn_Config = this.Factory.CreateRibbonButton();
            this.grp_Data = this.Factory.CreateRibbonGroup();
            this.btn_LoadTables = this.Factory.CreateRibbonButton();
            this.btn_loadData = this.Factory.CreateRibbonButton();
            this.btn_ShowHideSObList = this.Factory.CreateRibbonButton();
            this.btn_CommitChanges = this.Factory.CreateRibbonButton();
            this.btn_CloneSelection = this.Factory.CreateRibbonButton();
            this.sfRibbonTab.SuspendLayout();
            this.grp_config.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.grp_Data.SuspendLayout();
            this.SuspendLayout();
            // 
            // sfRibbonTab
            // 
            this.sfRibbonTab.Groups.Add(this.grp_config);
            this.sfRibbonTab.Groups.Add(this.grp_Data);
            this.sfRibbonTab.Label = "sforce";
            this.sfRibbonTab.Name = "sfRibbonTab";
            // 
            // grp_config
            // 
            this.grp_config.Items.Add(this.dropDown_TargetOrg);
            this.grp_config.Items.Add(this.buttonGroup1);
            this.grp_config.Items.Add(this.btn_Config);
            this.grp_config.Label = "Config";
            this.grp_config.Name = "grp_config";
            // 
            // dropDown_TargetOrg
            // 
            this.dropDown_TargetOrg.Label = "Target Org";
            this.dropDown_TargetOrg.Name = "dropDown_TargetOrg";
            this.dropDown_TargetOrg.ScreenTip = "Select an authorized Org to work against.";
            this.dropDown_TargetOrg.SizeString = "MMMMMMMMMMM";
            this.dropDown_TargetOrg.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown_TargetOrg_SelectionChanged);
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.gallery_AuthOrg);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // gallery_AuthOrg
            // 
            ribbonDropDownItemImpl1.Label = "Sandbox";
            ribbonDropDownItemImpl2.Label = "Production";
            this.gallery_AuthOrg.Items.Add(ribbonDropDownItemImpl1);
            this.gallery_AuthOrg.Items.Add(ribbonDropDownItemImpl2);
            this.gallery_AuthOrg.Label = "+Auth an Org";
            this.gallery_AuthOrg.Name = "gallery_AuthOrg";
            this.gallery_AuthOrg.RowCount = 2;
            this.gallery_AuthOrg.ScreenTip = "Login into an Org, production or sandbox.";
            this.gallery_AuthOrg.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gallery_AuthOrg_Click);
            // 
            // btn_Config
            // 
            this.btn_Config.Label = "Config";
            this.btn_Config.Name = "btn_Config";
            this.btn_Config.ScreenTip = "Basic configurations.";
            this.btn_Config.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Config_Click);
            // 
            // grp_Data
            // 
            this.grp_Data.Items.Add(this.btn_LoadTables);
            this.grp_Data.Items.Add(this.btn_loadData);
            this.grp_Data.Items.Add(this.btn_ShowHideSObList);
            this.grp_Data.Items.Add(this.btn_CommitChanges);
            this.grp_Data.Items.Add(this.btn_CloneSelection);
            this.grp_Data.Label = "Data";
            this.grp_Data.Name = "grp_Data";
            // 
            // btn_LoadTables
            // 
            this.btn_LoadTables.Enabled = false;
            this.btn_LoadTables.Label = "Load Tables";
            this.btn_LoadTables.Name = "btn_LoadTables";
            this.btn_LoadTables.ScreenTip = "Load sObjects";
            this.btn_LoadTables.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_LoadTables_Click);
            // 
            // btn_loadData
            // 
            this.btn_loadData.Enabled = false;
            this.btn_loadData.Label = "Load Data";
            this.btn_loadData.Name = "btn_loadData";
            this.btn_loadData.ScreenTip = "Load Data for current sheet";
            this.btn_loadData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_LoadData_Click);
            // 
            // btn_ShowHideSObList
            // 
            this.btn_ShowHideSObList.Enabled = false;
            this.btn_ShowHideSObList.Label = "Show/Hide sObject List";
            this.btn_ShowHideSObList.Name = "btn_ShowHideSObList";
            this.btn_ShowHideSObList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ShowHideTaskPane_Click);
            // 
            // btn_CommitChanges
            // 
            this.btn_CommitChanges.Enabled = false;
            this.btn_CommitChanges.Label = "Commit Changes";
            this.btn_CommitChanges.Name = "btn_CommitChanges";
            this.btn_CommitChanges.ScreenTip = "Commit all changes on current sheet, update, delete, create";
            this.btn_CommitChanges.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_CommitChanges_Click);
            // 
            // btn_CloneSelection
            // 
            this.btn_CloneSelection.Enabled = false;
            this.btn_CloneSelection.Label = "Clone Selection";
            this.btn_CloneSelection.Name = "btn_CloneSelection";
            this.btn_CloneSelection.ScreenTip = "Clone selected records into target org, may not be the org which load data from.";
            this.btn_CloneSelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_CopySelection_Click);
            // 
            // sforceRibbon
            // 
            this.Name = "sforceRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.sfRibbonTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.sforceRibbon_Load);
            this.sfRibbonTab.ResumeLayout(false);
            this.sfRibbonTab.PerformLayout();
            this.grp_config.ResumeLayout(false);
            this.grp_config.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.grp_Data.ResumeLayout(false);
            this.grp_Data.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab sfRibbonTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_Data;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ShowHideSObList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_loadData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_CommitChanges;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_LoadTables;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_config;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallery_AuthOrg;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_TargetOrg;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_CloneSelection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Config;
    }

    partial class ThisRibbonCollection
    {
        internal sforceRibbon sforceRibbon
        {
            get { return this.GetRibbon<sforceRibbon>(); }
        }
    }
}
