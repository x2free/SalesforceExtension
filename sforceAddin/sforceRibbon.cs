using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using sforceAddin.SFDC;
using System.Net;
using System.Windows.Forms;
using Microsoft.Office.Tools;

using Interop = Microsoft.Office.Interop.Excel;
using Tools = Microsoft.Office.Tools.Excel;

using sforceAddin.sforce;

namespace sforceAddin
{
    public partial class sforceRibbon
    {
        CustomTaskPane taskPane;
        // sforce.SForceClient sfClient;
        // System.Data.DataTable dt;
        // System.Data.DataSet ds = new System.Data.DataSet();
        private UI.SObjectTreeViewControl treeView;

        private void sforceRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            treeView = new UI.SObjectTreeViewControl();
            treeView.tv_sobjs.ImageList = UI.SObjectNodeBase.ImgList;
            treeView.tv_sobjs.NodeMouseDoubleClick += Tv_sobjs_NodeMouseDoubleClick;
            treeView.tv_sobjs.NodeMouseClick += Tv_sobjs_NodeMouseClick;
        }

        System.Net.Sockets.TcpListener myListener;

        private Cursor cursorState = null;

        private void Tv_sobjs_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            Cursor curCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            UI.SObjectNodeBase node = e.Node as UI.SObjectNodeBase;

            if (node != null) //eg, root node
            {
                node.LoadNode();
            }
            // root node wihout children
            else if (e.Node is TreeNode && e.Node.Parent == null && e.Node.Nodes.Count == 0)
            {
                TreeNode root = e.Node as TreeNode;
                List<sforce.SObjectEntryBase> sobjList =  SForceClient.Instance.GetSObjects();
                ExpandNode(root, sobjList, true);
            }

            Cursor.Current = curCursor;
        }

        private void btn_ShowHideTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            if (taskPane == null)
            {
                //btn_ShowHideTaskPane.Enabled = false;
                //return;
                taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(treeView, "SObject List");
                taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
                // taskPane.VisibleChanged += TaskPane_VisibleChanged;
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

        private void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            var sheets = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets;
            string tableName = null;
            if (taskPane.Visible)
            {
                foreach (Interop.Worksheet sheet in sheets)
                {
                    if (sheet.ListObjects != null)
                    {
                        SForceClient.Instance.SheetNameToTableNameMap.TryGetValue(sheet.Name, out tableName);
                        Tools.ListObject listObj = Globals.Factory.GetVstoObject(sheet.ListObjects[tableName]);
                        if (listObj != null && listObj.DataSource != null)
                        {
                            listObj.Disconnect();
                        }
                    }
                }
            }
            else
            {
                foreach (Interop.Worksheet sheet in sheets)
                {
                    if (sheet.ListObjects != null)
                    {
                        SForceClient.Instance.SheetNameToTableNameMap.TryGetValue(sheet.Name, out tableName);
                        // Tools.ListObject listObj = Globals.Factory.GetVstoObject(sheet.ListObjects[sheet.Name]);
                        Tools.ListObject listObj = Globals.Factory.GetVstoObject(sheet.ListObjects[tableName]);
                        if (listObj != null && listObj.DataSource == null)
                        {
                            listObj.SetDataBinding(sforce.SForceClient.Instance.DataSet.Tables[tableName]);
                        }
                    }
                }
            }
        }

        private void btn_LoadData_Click(object sender, RibbonControlEventArgs e)
        {
            Cursor curCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
                string tableName = null;
                SForceClient.Instance.SheetNameToTableNameMap.TryGetValue(sheet.Name, out tableName);

                Microsoft.Office.Interop.Excel.ListObject listObj = null;
                if (string.IsNullOrEmpty(tableName))
                {
                    listObj = sheet.ListObjects.Item[1];
                    tableName = listObj.Name;

                    System.Data.DataTable dt2 = new System.Data.DataTable(tableName);
                    foreach (Microsoft.Office.Interop.Excel.Range headerCell in listObj.HeaderRowRange.Cells)
                    {
                        string fieldName = headerCell.Name.Name.Substring(tableName.Length + 1);
                        dt2.Columns.Add(fieldName);
                    }

                    SForceClient.Instance.SheetNameToTableNameMap.Add(sheet.Name, listObj.Name);
                    SForceClient.Instance.DataSet.Tables.Add(dt2);
                }
                else
                {
                    foreach (Microsoft.Office.Interop.Excel.ListObject obj in sheet.ListObjects)
                    {
                        // if (String.Equals(sheet.Name, obj.Name, StringComparison.InvariantCultureIgnoreCase))
                        if (String.Equals(tableName, obj.Name, StringComparison.InvariantCultureIgnoreCase))
                        {
                            listObj = obj;
                            break;
                        }
                    }
                }

                if (listObj == null)
                {
                    return;
                }

                Microsoft.Office.Tools.Excel.ListObject hostListObject = Globals.Factory.GetVstoObject(listObj);

                // Disconnect the datasource temporarily, otherwise, it may be flashing due to row by row refresh
                if (hostListObject.DataSource != null)
                {
                    hostListObject.Disconnect();
                }

                // string tableName = listObj.DisplayName;
                // string tableName = hostListObject.Name;
                StringBuilder sb = new StringBuilder();

                //foreach (Microsoft.Office.Interop.Excel.ListColumn col in listObj.ListColumns)
                //{
                //    sb.AppendFormat("{0},", col.Name);
                //}
                List<string> columnNameList = new List<string>();

                foreach (Microsoft.Office.Interop.Excel.Range headerCell in hostListObject.HeaderRowRange.Cells)
                {
                    string fieldName = headerCell.Name.Name.Substring(hostListObject.Name.Length + 1);
                    // sb2.AppendFormat("{0},", headerCell.Name.Name);
                    var v1 = headerCell.Name;
                    var v2 = headerCell.Name.Name;
                    // sb.AppendFormat("{0},", headerCell.Name.Name.Substring(hostListObject.Name.Length + 1));
                    sb.AppendFormat("{0},", fieldName);

                    // columnNameList.Add(headerCell.Name.Name.Replace('.', '_'));
                    columnNameList.Add(fieldName);
                }

                // get text instead of API names
                //foreach (Microsoft.Office.Interop.Excel.ListColumn col in listObj.ListColumns)
                //{
                //    sb.AppendFormat("{0},", col.Name);
                //}

                //foreach (Microsoft.Office.Interop.Excel.Name item in sheet.Names)
                //{

                //}

                //List<string> columnNameList = new List<string>();
                //foreach (Microsoft.Office.Interop.Excel.Name item in Globals.ThisAddIn.Application.Names)
                //{
                //    if (item.Name != null && item.Name.StartsWith(listObj.Name))
                //    {
                //        sb.AppendFormat("{0},", item.Name.Substring(listObj.Name.Length + 1));

                //        columnNameList.Add(item.Name.Replace('.', '_'));
                //    }
                //}

                //foreach (Microsoft.Office.Interop.Excel.Range cell in listObj.HeaderRowRange)
                //{
                //    Microsoft.Office.Interop.Excel.Name name = (Microsoft.Office.Interop.Excel.Name)cell.Name;
                //    sb.AppendFormat("{0},", name.Name.Substring(listObj.Name.Length + 1));
                //}


                sb.Remove(sb.Length - 1, 1);
                string queryStr = String.Format("SELECT {0} FROM {1}", sb.ToString(), tableName);

                System.Data.DataTable dt = (System.Data.DataTable)SForceClient.Instance.DataSet.Tables[tableName];

                List<string> columnToRemove = new List<string>();
                foreach (System.Data.DataColumn col in dt.Columns)
                {
                    if (!columnNameList.Contains(col.ColumnName))
                    {
                        // dt.Columns.Remove(col); // // Collection was modified; enumeration operation may not execute.
                        columnToRemove.Add(col.ColumnName);
                    }
                }


                foreach (string colName in columnToRemove)
                {
                    dt.Columns.Remove(colName);
                }

                bool isTableExist = dt != null;
                dt = SForceClient.Instance.ExecQuery(queryStr, tableName, dt);

                if (dt == null)
                {
                    MessageBox.Show("No Data loaded", "sforce Addin", System.Windows.Forms.MessageBoxButtons.OK);
                    return;
                }

                dt.AcceptChanges();

                if (!isTableExist)
                {
                    SForceClient.Instance.DataSet.Tables.Add(dt);
                }

                // Microsoft.Office.Tools.Excel.ApplicationFactory factory = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook).ActiveSheet;
                // Microsoft.Office.Tools.Excel.Worksheet sheet2 = (Microsoft.Office.Tools.Excel.Worksheet)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook).ActiveSheet;
                // sheet2.lis

                // hostListObject.SetDataBinding(dt, "", sb.ToString().Split(','));
                hostListObject.SetDataBinding(dt, "", columnNameList.ToArray());
                //hostListObject.SetDataBinding(dt);
                hostListObject.RefreshDataRows();

                Cursor.Current = curCursor;
            }
            catch (Exception ex)
            {
                Cursor.Current = curCursor;
                MessageBox.Show(ex.Message);
            }
        }

        private void btn_CommitChanges_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Tools.Excel.Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveSheet);
            Microsoft.Office.Tools.Excel.ListObject listObj = null;
            string tableName = null;
            SForceClient.Instance.SheetNameToTableNameMap.TryGetValue(sheet.Name, out tableName);

            foreach (Microsoft.Office.Interop.Excel.ListObject item in sheet.ListObjects)
            {
                if (String.Equals(tableName, item.Name, StringComparison.InvariantCultureIgnoreCase))
                {
                    listObj = Globals.Factory.GetVstoObject(item);
                    break;
                }
            }

            // System.Data.DataTable dt = (System.Data.DataTable)SForceClient.Instance.DataSet.Tables[Globals.ThisAddIn.Application.ActiveSheet.Name];
            System.Data.DataTable dt = (System.Data.DataTable)SForceClient.Instance.DataSet.Tables[tableName];

            Cursor curCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                System.Data.DataTable updatedTable = dt.GetChanges(System.Data.DataRowState.Modified);
                System.Data.DataTable deletedTable = dt.GetChanges(System.Data.DataRowState.Deleted);
                System.Data.DataTable addedTable = dt.GetChanges(System.Data.DataRowState.Added);

                List<object> resultList = SForceClient.Instance.DoUpdate(dt);

                bool hasError = false;
                ProcessResult(resultList, out hasError);

                if (!hasError)
                {
                    dt.AcceptChanges();
                }

                if (listObj != null)
                {
                    if (listObj.DataSource == null)
                    {
                        listObj.SetDataBinding(dt);
                    }
                }

                Cursor.Current = curCursor;
            }
            catch (Exception ex)
            {
                Cursor.Current = curCursor;

                MessageBox.Show(ex.Message);
            }

            //updatedTable = dt.GetChanges(System.Data.DataRowState.Modified);
            //deletedTable = dt.GetChanges(System.Data.DataRowState.Deleted);
            //addedTable = dt.GetChanges(System.Data.DataRowState.Added);
        }

        private void btn_LoadTables_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                sforce.SFSession session = sforce.SFSessionManager.Instance.ActiveSession;
                if (session == null)
                {
                    string instanceName = this.dropDown_TargetOrg.SelectedItem.Label;
                    session = sforce.SFSessionManager.Instance.FindSession(instanceName);
                    session.IsActive = true;
                }
                SForceClient.Instance.SetSession(session);

                Cursor oldCursor = Cursor.Current;
                Cursor.Current = Cursors.WaitCursor;

                List<sforce.SObjectEntryBase> sobjectList = SForceClient.Instance.GetSObjects();

                //if (treeView == null)
                //{
                //    treeView = new UI.SObjectTreeViewControl();
                //}

                FufillTreeviewWithSObjectList(treeView, sobjectList);

                if (taskPane == null)
                {
                    taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(treeView, "SObject List");
                    taskPane.Visible = true;
                    taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;

                    // taskPane.VisibleChanged += TaskPane_VisibleChanged;
                }


                // taskPane.VisibleChanged += TaskPane_VisibleChanged;

                Cursor.Current = oldCursor;

                btn_ShowHideSObList.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ExpandNode(TreeNode node, List<sforce.SObjectEntryBase> sobjectList, bool expandChildren)
        {
            if (sobjectList == null)
            {
                return;
            }

            node.Nodes.Clear();

            foreach (var item in sobjectList)
            {
                UI.SObjectNodeBase subNode = UI.SObjectNodeBase.CreateNode(item, node);
                if (subNode == null)
                {
                    continue;
                }

                node.Nodes.Add(subNode);

                if (expandChildren)
                {
                    ExpandNode(subNode, item.Children, expandChildren);
                }
            }
        }

        private void FufillTreeviewWithSObjectList(UI.SObjectTreeViewControl treeView, List<sforce.SObjectEntryBase> sobjectList)
        {
            treeView.tv_sobjs.BeginUpdate();
            treeView.tv_sobjs.Nodes.Clear();

            TreeNode root = treeView.tv_sobjs.Nodes.Add("SObjects");
            // treeView.tv_sobjs.Nodes.Add("Custom Settings");

            //foreach (sforce.SObjectEntry item in sobjectList)
            //{
            //    // TreeNode node = new UI.SObjectNodeBase(item.Name, item.Label, sfClient);
            //    // TreeNode node = new UI.SObjectNodeBase(item, sfClient);
            //    // node.Collapse();
            //    // treeView.tv_sobjs.Nodes[0].Nodes.Add(node);

            //    UI.SObjectNode node = new UI.SObjectNode(item, parent);
            //    if (item.IsCustomSetting)
            //    {
            //        treeView.tv_sobjs.Nodes[1].Nodes.Add(node);
            //    }
            //    else
            //    {
            //        treeView.tv_sobjs.Nodes[0].Nodes.Add(node);
            //    }
            //}
            ExpandNode(root, sobjectList, true);

            treeView.tv_sobjs.EndUpdate();
        }

        private void Tv_sobjs_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
            {
                return;
            }

            TreeView tv = sender as TreeView;
            TreeNode node = tv.GetNodeAt(e.X, e.Y);
            tv.SelectedNode = node;

            ContextMenuStrip rightClickMenu = new ContextMenuStrip();
            rightClickMenu.Items.Add("Reload", null, (o, ev) =>
                    {
                        Cursor curCursor = Cursor.Current;
                        Cursor.Current = Cursors.WaitCursor;

                        //ToolStripItem item = o as ToolStripItem;
                        //if (item != null)
                        //{
                        //    ContextMenuStrip cxtMenuStrip = item.Owner as ContextMenuStrip;

                        //    if (cxtMenuStrip != null)
                        //    {
                        //        var obj = cxtMenuStrip.SourceControl;
                        //    }
                        //}

                        List<sforce.SObjectEntryBase> objList = null;

                        //if (node.Parent == null)
                        //{
                        //    objList = sfClient.getSObjects(true);
                        //    FufillTreeviewWithSObjectList(treeView, objList);
                        //}
                        //else if (node is UI.SObjectNode)
                        //{
                        //    (node as UI.SObjectNode).LoadNode(true);
                        //}

                        if (node.Parent == null)
                        {
                            objList = SForceClient.Instance.GetSObjects(true);
                        }
                        else if (node is UI.SObjectNode)
                        {
                            objList = SForceClient.Instance.DescribeSObject((node as UI.SObjectNode).SObjEntry);
                        }

                        ExpandNode(node, objList, true);

                        Cursor.Current = curCursor;
                    });
            rightClickMenu.Show(sender as Control, new System.Drawing.Point(e.X, e.Y));
        }

        private void gallery_AuthOrg_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonGallery gallery = sender as RibbonGallery;
            if (gallery == null)
            {
                return;
            }

            switch (gallery.SelectedItemIndex)
            {
                case 1: // production
                    Auth.AuthUtil.baseUrl = "https://login.salesforce.com";
                    break;
                default: // sandbox
                    Auth.AuthUtil.baseUrl = "https://test.salesforce.com";
                    break;
            }
            Auth.AuthUtil.doAuth(updateOrgList);
        }

        private bool updateOrgList(sforce.SFSession session)
        {
            RibbonDropDownItem newItem = this.dropDown_TargetOrg.Items.FirstOrDefault(item => item.Label == session.InstanceName);
            if (newItem != null)
            {
                session = sforce.SFSessionManager.Instance.FindSession(session.InstanceName);
            }
            else
            {
                newItem = Factory.CreateRibbonDropDownItem();
                newItem.Label = session.InstanceName;

                this.dropDown_TargetOrg.Items.Add(newItem);
            }

            // conn.Active();
            // this.dropDown_org.SelectedItem = newItem;

            if (this.dropDown_TargetOrg.Items.Count == 1) // If only one org
            {
                session.IsActive = true;
                SForceClient.Instance.SetSession(session);
            }

            // enable buttons
            // this.dropDown_TargetOrg.Enabled = true;
            this.btn_LoadTables.Enabled = true;
            this.btn_ShowHideSObList.Enabled = true;
            this.btn_loadData.Enabled = true;
            this.btn_CommitChanges.Enabled = true;
            this.btn_CloneSelection.Enabled = true;

            return true;
        }

        private void EnableButtons(bool IsEnable)
        {
            this.btn_LoadTables.Enabled = true;
            this.btn_ShowHideSObList.Enabled = true;
            this.btn_loadData.Enabled = true;
            this.btn_CommitChanges.Enabled = true;
            this.btn_CloneSelection.Enabled = true;
        }

        private void dropDown_TargetOrg_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            RibbonDropDown dropDown = sender as RibbonDropDown;
            if (dropDown == null)
            {
                return;
            }

            // Deactive current session
            // sforce.ConnectionManager.Instance.ActiveConnection.Deactive(); // session timeout
            sforce.SFSession session = sforce.SFSessionManager.Instance.ActiveSession;
            if (session != null)
            {
                session.IsActive = false;
            }

            string instanceName = dropDown.SelectedItem.Label;
            session = sforce.SFSessionManager.Instance.FindSession(instanceName);
            session.IsActive = true;

            SForceClient.Instance.SetSession(session);
            FufillTreeviewWithSObjectList(this.treeView, session.SObjects);

            if (this.treeView.tv_sobjs != null && this.treeView.tv_sobjs.TopNode != null)
            {
                this.treeView.tv_sobjs.TopNode.Expand();
            }
        }

        private bool apiVersion_Changed(string version)
        {
            // obsoleted since we are building the url dynamically?
            //double versionNum = 0;
            //bool isSuccess = double.TryParse(version, out versionNum);
            //if (isSuccess)
            //{
            //    Auth.AuthUtil.apiVersion = (int)versionNum;

            //    if (sfClient != null && sforce.ConnectionManager.Instance.ActiveConnection != null)
            //    {
            //        sfClient.init(sforce.ConnectionManager.Instance.ActiveConnection.Session);
            //    }

            //    return true;
            //}

            return false;
        }

        //private void apiVersion_TextChanged(object sender, RibbonControlEventArgs e)
        //{
        //    apiVersion_Changed(this.editbox_APIVersion.Text);
        //}

        private void btn_CopySelection_Click(object sender, RibbonControlEventArgs e)
        {
            Cursor curCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            string tableName = null;
            SForceClient.Instance.SheetNameToTableNameMap.TryGetValue(Globals.ThisAddIn.Application.ActiveSheet.Name, out tableName);

            // System.Data.DataTable dt = (System.Data.DataTable)SForceClient.Instance.DataSet.Tables[Globals.ThisAddIn.Application.ActiveSheet.Name];
            System.Data.DataTable dt = (System.Data.DataTable)SForceClient.Instance.DataSet.Tables[tableName];

            if (dt.GetChanges() != null)
            {
                DialogResult ret = MessageBox.Show("Save changes?", "sforce Addin", MessageBoxButtons.YesNoCancel);
                switch (ret)
                {
                    case DialogResult.Cancel:
                        return;
                    case DialogResult.Yes:
                        dt.AcceptChanges();
                        break;
                    case DialogResult.No:
                        dt.RejectChanges();
                        break;
                    default:
                        dt.RejectChanges(); // exception?
                        break;
                }
            }

            System.Data.DataTable changes = dt.GetChanges();
            if (changes != null)
            {
                throw new Exception("changes????");
            }

            try
            {
                foreach (Interop.Range item in ((Interop.Range)Globals.ThisAddIn.Application.Selection).Rows)
                {
                    // dt.Rows[item.Row - 2].SetModified(); // 2 = header + vsto starts from 1
                    dt.Rows[item.Row - 2].SetAdded(); // 2 = header + vsto starts from 1
                }

                List<object> resultList = SForceClient.Instance.DoUpdate(dt);

                bool hasError = false;
                ProcessResult(resultList, out hasError);

                if (!hasError)
                {
                    dt.AcceptChanges();
                }

                Cursor.Current = curCursor;
            }
            catch (Exception ex)
            {
                Cursor.Current = curCursor;
                MessageBox.Show(ex.Message);
            }
        }

        private UI.ConfigForm configForm;
        private void btn_Config_Click(object sender, RibbonControlEventArgs e)
        {
            if (configForm == null)
            {
                configForm = new UI.ConfigForm();
                configForm.APIVersionChnagedHandler += apiVersion_Changed;
            }

            configForm.ShowDialog();
        }

        private void ProcessResult(List<object> resultList, out bool hasError)
        {
            hasError = false;

            string resultTableName = "$$Result";
            System.Data.DataTable resultTable = SForceClient.Instance.DataSet.Tables[resultTableName];
            //// System.Data.DataTable errorTable = null;
            if (resultTable == null)
            {
                resultTable = new System.Data.DataTable();

                // resultTable.TableName = string.Format("{0}-{1}", tableName, DateTime.Today.ToString("yyMMddHHmmss"));
                resultTable.TableName = resultTableName;
                //// resultTable.Columns.Add("RecId", typeof(Guid));
                resultTable.Columns.Add("Id", typeof(string)); // Id or name (for insert)
                resultTable.Columns.Add("Operation", typeof(string)); // update/insert/delete
                //resultTable.Columns.Add("Status", typeof(bool)); // success or no
                resultTable.Columns.Add("Status", typeof(OpResultStatus));
                resultTable.Columns.Add("Errors", typeof(string)); // errors

                ////errorTable = new System.Data.DataTable("$$Errors");
                ////errorTable.Columns.Add("RecId", typeof(Guid));
                ////errorTable.Columns.Add("Error", typeof(string));
                ////errorTable.Columns.Add("Fields", typeof(string));
                ////errorTable.Columns.Add("Message", typeof(string));

                SForceClient.Instance.DataSet.Tables.Add(resultTable);
                SForceClient.Instance.SheetNameToTableNameMap.Add(resultTableName, resultTableName);
                //// SForceClient.Instance.DataSet.Tables.Add(errorTable);

                //// resultTable.ChildRelations.Add(new System.Data.DataRelation("R-E", resultTable.Columns["RecId"], errorTable.Columns["RecId"]));
            }

            ////if (errorTable == null)
            ////{
            ////    errorTable = resultTable.ChildRelations["R-E"].ChildTable;
            ////}

            foreach (var result in resultList)
            {
                System.Data.DataRow row = resultTable.NewRow();
                if (result is SFDC.UpsertResult)
                {
                    SFDC.UpsertResult ret = result as SFDC.UpsertResult;
                    ////Guid recId = Guid.NewGuid();
                    ////row["RecId"] = recId;
                    row["Id"] = ret.id;
                    row["Operation"] = ret.created ? "Insert" : "Update";
                    row["Status"] = ret.success ? OpResultStatus.Success : OpResultStatus.Failed;

                    if (!ret.success)
                    {
                        hasError = true;
                        row["Errors"] = ret.errors.Error2String();

                        ////System.Data.DataRow errRow = errorTable.NewRow();
                        ////foreach (var err in ret.errors)
                        ////{
                        ////    errRow["RecId"] = recId;
                        ////    errRow["Error"] = err.statusCode;
                        ////    errRow["Fields"] = string.Join(", ", err.fields);
                        ////    errRow["Message"] = err.message;
                        ////}
                    }
                }
                else if (result is SFDC.DeleteResult)
                {
                    SFDC.DeleteResult ret = result as SFDC.DeleteResult;
                    ////Guid recId = Guid.NewGuid();
                    ////row["RecId"] = recId;
                    row["Id"] = ret.id;
                    row["Operation"] = "Delete";
                    row["Status"] = ret.success ? OpResultStatus.Success : OpResultStatus.Failed;

                    if (!ret.success)
                    {
                        hasError = true;
                        row["Errors"] = ret.errors.Error2String();

                        ////System.Data.DataRow errRow = errorTable.NewRow();
                        ////foreach (var err in ret.errors)
                        ////{
                        ////    errRow["RecId"] = recId;
                        ////    errRow["Error"] = err.statusCode;
                        ////    errRow["Fields"] = string.Join(", ", err.fields);
                        ////    errRow["Message"] = err.message;
                        ////}
                    }
                }

                resultTable.Rows.Add(row);
            }

            Tools.Workbook workbook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);
            Tools.Worksheet worksheet = null;
            Tools.ListObject listObject = null;

            string tableName = null;
            foreach (Interop.Worksheet sheet in workbook.Sheets)
            {
                SForceClient.Instance.SheetNameToTableNameMap.TryGetValue(sheet.Name, out tableName);

                if (string.Equals(resultTableName, tableName, StringComparison.CurrentCultureIgnoreCase))
                {
                    worksheet = Globals.Factory.GetVstoObject(sheet);
                    break;
                }
            }

            if (worksheet == null)
            {
                worksheet = Globals.Factory.GetVstoObject(workbook.Sheets.Add());
                worksheet.Name = resultTableName;
            }

            foreach (Interop.ListObject listObj in worksheet.ListObjects)
            {
                if (string.Equals(resultTableName, listObj.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    listObject = Globals.Factory.GetVstoObject(listObj);
                    break;
                }
            }

            if (listObject == null)
            {
                listObject = Globals.Factory.GetVstoObject(worksheet.ListObjects.AddEx());
                listObject.Name = resultTableName;
                //// listObject.SetDataBinding(SForceClient.Instance.DataSet, resultTableName, "Id", "Operation", "Status", "Error", "Fields", "Message");
                //// listObject.SetDataBinding(SForceClient.Instance.DataSet, "", "Id", "Operation", "Status", "Error", "Fields", "Message");
                //// listObject.SetDataBinding(SForceClient.Instance.DataSet, resultTableName);
                listObject.SetDataBinding(SForceClient.Instance.DataSet, resultTableName, "Id", "Operation", "Status", "Errors");
                listObject.AutoSetDataBoundColumnHeaders = true; // else will show column1, column2, etc

                // listObject.ListColumns["Errors"].Range.AutoFit(); // first time will throw error
                // listObject.ListColumns["Errors"].Range.ColumnWidth = "Auto"; // unable to set ColumnWidth of Range clas
                // listObject.ListColumns["Errors"].Range.Columns.AutoFit();
                // listObject.ListColumns["Errors"].Range.EntireColumn.AutoFit();
                worksheet.Columns.AutoFit();

                //// set auto width
                //foreach (Interop.ListColumn col in listObject.ListColumns)
                //{
                //    col.Range.AutoFit();
                //}
            }

            listObject.RefreshDataRows();
            // listObject.ListColumns["Errors"].Range.AutoFit(); // always error

            worksheet.Activate();
        }
    }

    internal enum OpResultStatus {Failed, Success}
}
