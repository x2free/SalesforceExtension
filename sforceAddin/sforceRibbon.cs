using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using sforceAddin.SFDC;
using System.Net;
using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace sforceAddin
{
    public partial class sforceRibbon
    {
        CustomTaskPane taskPane;
        sforce.SForceClient sfClient;
        // System.Data.DataTable dt;
        System.Data.DataSet ds = new System.Data.DataSet();

        private void sforceRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_login_Click(object sender, RibbonControlEventArgs e)
        {
            Cursor oldCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            string userName = "";
            string password = "";
            string secuToken = "";

            sfClient = new sforce.SForceClient();
            bool isSucess = sfClient.login(userName, password, secuToken);

            if (!isSucess)
            {
                Cursor.Current = oldCursor;
                return;
            }

            List<sforce.SObjectEntryBase> sobjectList = sfClient.getSObjects();

            //UI.sforceListViewControl lvControl = new UI.sforceListViewControl();
            //foreach (var item in sobjectList)
            //{
            //    lvControl.listview_sobjs.Items.Add(String.Format("{0}({1})", item.Label, item.Name));
            //}
            //lvControl.listview_sobjs.AutoResizeColumns(System.Windows.Forms.ColumnHeaderAutoResizeStyle.ColumnContent);

            //taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(lvControl, "SObject List");
            //taskPane.Visible = true;
            //taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;

            UI.SObjectTreeViewControl treeView = new UI.SObjectTreeViewControl();
            treeView.tv_sobjs.BeginUpdate();

            treeView.tv_sobjs.Nodes.Add("SObjects");
            treeView.tv_sobjs.Nodes.Add("Custom Settings");
            foreach (sforce.SObjectEntry item in sobjectList)
            {
                // TreeNode node = new UI.SObjectNodeBase(item.Name, item.Label, sfClient);
                // TreeNode node = new UI.SObjectNodeBase(item, sfClient);
                // node.Collapse();
                // treeView.tv_sobjs.Nodes[0].Nodes.Add(node);

                if (item.IsCustomSetting)
                {
                    treeView.tv_sobjs.Nodes[1].Nodes.Add(new UI.SObjectNode(item, null));
                }
                else
                {
                    treeView.tv_sobjs.Nodes[0].Nodes.Add(new UI.SObjectNode(item, null));
                }
            }

            treeView.tv_sobjs.NodeMouseDoubleClick += Tv_sobjs_NodeMouseDoubleClick;
            treeView.tv_sobjs.EndUpdate();

            taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(treeView, "SObject List");
            taskPane.Visible = true;
            taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;


            // taskPane.VisibleChanged += TaskPane_VisibleChanged;

            Cursor.Current = oldCursor;

            btn_taskPane.Enabled = true;
        }

        private void Tv_sobjs_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            Cursor curCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            UI.SObjectNodeBase node = e.Node as UI.SObjectNodeBase;

            if (node != null) //eg, root node
            {
                node.dbClick();
            }

            Cursor.Current = curCursor;
        }

        private void btn_taskPane_Click(object sender, RibbonControlEventArgs e)
        {
            if (taskPane == null)
            {
                btn_taskPane.Enabled = false;
                return;
            }

            taskPane.Visible = !taskPane.Visible;
            //UI.sforceListViewControl lvControl = taskPane.Control as UI.sforceListViewControl;
            //if (lvControl == null)
            //{
            //    lvControl.AutoSize = true;
            //    // btn_taskPane.Enabled = false;
            //    return;
            //}
        }

        private void btn_load_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;


            Microsoft.Office.Interop.Excel.ListObject listObj = null;
            foreach (Microsoft.Office.Interop.Excel.ListObject obj in sheet.ListObjects)
            {
                if (String.Equals(sheet.Name, obj.Name, StringComparison.InvariantCultureIgnoreCase))
                {
                    listObj = obj;
                }
            }

            if (listObj == null)
            {
                return;
            }

            // string tableName = listObj.DisplayName;
            string tableName = listObj.Name;
            StringBuilder sb = new StringBuilder();

            //foreach (Microsoft.Office.Interop.Excel.ListColumn col in listObj.ListColumns)
            //{
            //    sb.AppendFormat("{0},", col.Name);
            //}

            //foreach (Microsoft.Office.Interop.Excel.Range headerCell in listObj.HeaderRowRange.Cells)
            //{
            //    sb.AppendFormat("{0},", headerCell.Name.Name);
            //}

            // get text instead of API names
            //foreach (Microsoft.Office.Interop.Excel.ListColumn col in listObj.ListColumns)
            //{
            //    sb.AppendFormat("{0},", col.Name);
            //}

            //foreach (Microsoft.Office.Interop.Excel.Name item in sheet.Names)
            //{

            //}

            foreach (Microsoft.Office.Interop.Excel.Name item in Globals.ThisAddIn.Application.Names)
            {
                if (item.Name != null && item.Name.StartsWith(listObj.Name))
                {
                    sb.AppendFormat("{0},", item.Name.Substring(listObj.Name.Length + 1));
                }
            }

            //foreach (Microsoft.Office.Interop.Excel.Range cell in listObj.HeaderRowRange)
            //{
            //    Microsoft.Office.Interop.Excel.Name name = (Microsoft.Office.Interop.Excel.Name)cell.Name;
            //    sb.AppendFormat("{0},", name.Name.Substring(listObj.Name.Length + 1));
            //}


            sb.Remove(sb.Length - 1, 1);
            string queryStr = String.Format("SELECT {0} FROM {1}", sb.ToString(), tableName);

            System.Data.DataTable dt = (System.Data.DataTable)ds.Tables[tableName];
            bool isTableExist = dt != null;
            dt = sfClient.execQuery(queryStr, tableName, dt);
            dt.AcceptChanges();

            if (!isTableExist)
            {
                ds.Tables.Add(dt);
            }

            // Microsoft.Office.Tools.Excel.ApplicationFactory factory = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook).ActiveSheet;
            // Microsoft.Office.Tools.Excel.Worksheet sheet2 = (Microsoft.Office.Tools.Excel.Worksheet)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook).ActiveSheet;
            // sheet2.lis

            Microsoft.Office.Tools.Excel.ListObject hostListObject = Globals.Factory.GetVstoObject(listObj);
            hostListObject.SetDataBinding(dt);
            hostListObject.RefreshDataRows();
        }

        private void btn_upsert_Click(object sender, RibbonControlEventArgs e)
        {
            System.Data.DataTable dt = (System.Data.DataTable)ds.Tables[Globals.ThisAddIn.Application.ActiveSheet.Name];

            System.Data.DataTable updatedTable = dt.GetChanges(System.Data.DataRowState.Modified);
            System.Data.DataTable deletedTable = dt.GetChanges(System.Data.DataRowState.Deleted);
            System.Data.DataTable addedTable = dt.GetChanges(System.Data.DataRowState.Added);

            dt.AcceptChanges();

            updatedTable = dt.GetChanges(System.Data.DataRowState.Modified);
            deletedTable = dt.GetChanges(System.Data.DataRowState.Deleted);
            addedTable = dt.GetChanges(System.Data.DataRowState.Added);
        }
    }
}
