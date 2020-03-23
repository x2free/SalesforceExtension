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


        public void StartListen()
        {
            int iStartPos = 0;
            String sRequest;
            String sDirName;
            String sRequestedFile;
            //String sErrorMessage;
            //String sLocalDir;
            //String sMyWebServerRoot = "C:\\MyWebServerRoot\\";
            //String sPhysicalFilePath = "";
            //String sFormattedMessage = "";
            //String sResponse = "";
            while (true)
            {
                //Accept a new connection  
                System.Net.Sockets.Socket mySocket = myListener.AcceptSocket();
                Console.WriteLine("Socket Type " + mySocket.SocketType);
                if (mySocket.Connected)
                {
                    Console.WriteLine(@"\nClient Connected!!\n==================\n CLient IP { 0}\n", mySocket.RemoteEndPoint);
                    //make a byte array and receive data from the client   
                    Byte[] bReceive = new Byte[1024];
                    int i = mySocket.Receive(bReceive, bReceive.Length, 0);
                    //Convert Byte to String  
                    string sBuffer = Encoding.ASCII.GetString(bReceive);
                    //At present we will only deal with GET type  
                    if (sBuffer.Substring(0, 3) != "GET")
                    {
                        Console.WriteLine("Only Get Method is supported..");
                        mySocket.Close();
                        return;
                    }
                    // Look for HTTP request  
                    iStartPos = sBuffer.IndexOf("HTTP", 1);
                    // Get the HTTP text and version e.g. it will return "HTTP/1.1"  
                    string sHttpVersion = sBuffer.Substring(iStartPos, 8);
                    // Extract the Requested Type and Requested file/directory  
                    sRequest = sBuffer.Substring(0, iStartPos - 1);
                    //Replace backslash with Forward Slash, if Any  
                    sRequest.Replace("\\", "/");
                    //If file name is not supplied add forward slash to indicate   
                    //that it is a directory and then we will look for the   
                    //default file name..  
                    if ((sRequest.IndexOf(".") < 1) && (!sRequest.EndsWith("/")))
                    {
                        sRequest = sRequest + "/";
                    }
                    //Extract the requested file name  
                    iStartPos = sRequest.LastIndexOf("/") + 1;
                    sRequestedFile = sRequest.Substring(iStartPos);
                    //Extract The directory Name  
                    sDirName = sRequest.Substring(sRequest.IndexOf("/"), sRequest.LastIndexOf("/") - 3);
                }
            }
        }

        System.Net.Sockets.TcpListener myListener;
        int port = 5050;

        private Cursor cursorState = null;
        private void btn_login_Click(object sender, RibbonControlEventArgs e)
        {
            cursorState = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            // Auth.AuthUtil.doAuth(initSFClient);

            ////start listing on the given port  
            //myListener = new System.Net.Sockets.TcpListener(IPAddress.Parse("127.0.0.1"), port);
            //myListener.Start();
            //Console.WriteLine("Web Server Running... Press ^C to Stop...");
            ////start the thread which calls the method 'StartListen'  
            //System.Threading.Thread th = new System.Threading.Thread(new System.Threading.ThreadStart(StartListen));
            //th.Start();

            /*

            Cursor oldCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            string userName = "";
            string password = "";
            string secuToken = "";

            sfClient = new sforce.SForceClient();
            // bool isSucess = sfClient.login(userName, password, secuToken);

            sforce.SFSession sfSession = sforce.SFSession.GetSession();
            bool isSucess = sfClient.login(sfSession);

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
            */
        }

        private bool initSFClient(sforce.SFSession session)
        {
            //if (session == null || !session.IsValid)
            //{
            //    return false;
            //}

            //sfClient = new sforce.SForceClient();
            //sfClient.init(session);

            //Cursor.Current = this.cursorState;

            return true;
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

        private void btn_ShowHideTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            if (taskPane == null)
            {
                //btn_ShowHideTaskPane.Enabled = false;
                //return;
                taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(treeView, "SObject List");
                taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
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

        private void btn_LoadData_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;


            Microsoft.Office.Interop.Excel.ListObject listObj = null;
            foreach (Microsoft.Office.Interop.Excel.ListObject obj in sheet.ListObjects)
            {
                if (String.Equals(sheet.Name, obj.Name, StringComparison.InvariantCultureIgnoreCase))
                {
                    listObj = obj;
                    break;
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
            List<string> columnNameList = new List<string>();

            foreach (Microsoft.Office.Interop.Excel.Range headerCell in listObj.HeaderRowRange.Cells)
            {
                // sb2.AppendFormat("{0},", headerCell.Name.Name);
                var v1 = headerCell.Name;
                var v2 = headerCell.Name.Name;
                sb.AppendFormat("{0},", headerCell.Name.Name.Substring(listObj.Name.Length + 1));

                columnNameList.Add(headerCell.Name.Name.Replace('.', '_'));
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

            System.Data.DataTable dt = (System.Data.DataTable)ds.Tables[tableName];
            bool isTableExist = dt != null;
            dt = sfClient.execQuery(queryStr, tableName, dt);

            if (dt == null)
            {
                MessageBox.Show("No Data loaded", "sforce Addin", System.Windows.Forms.MessageBoxButtons.OK);
                return;
            }

            dt.AcceptChanges();

            if (!isTableExist)
            {
                ds.Tables.Add(dt);
            }

            // Microsoft.Office.Tools.Excel.ApplicationFactory factory = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook).ActiveSheet;
            // Microsoft.Office.Tools.Excel.Worksheet sheet2 = (Microsoft.Office.Tools.Excel.Worksheet)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook).ActiveSheet;
            // sheet2.lis

            Microsoft.Office.Tools.Excel.ListObject hostListObject = Globals.Factory.GetVstoObject(listObj);
            // hostListObject.SetDataBinding(dt, "", sb.ToString().Split(','));
            // hostListObject.SetDataBinding(dt, "", columnNameList.ToArray());
            hostListObject.SetDataBinding(dt);
            hostListObject.RefreshDataRows();
        }

        private void btn_CommitChanges_Click(object sender, RibbonControlEventArgs e)
        {
            System.Data.DataTable dt = (System.Data.DataTable)ds.Tables[Globals.ThisAddIn.Application.ActiveSheet.Name];

            System.Data.DataTable updatedTable = dt.GetChanges(System.Data.DataRowState.Modified);
            System.Data.DataTable deletedTable = dt.GetChanges(System.Data.DataRowState.Deleted);
            System.Data.DataTable addedTable = dt.GetChanges(System.Data.DataRowState.Added);

            sfClient.doUpdate(dt);

            dt.AcceptChanges();

            //updatedTable = dt.GetChanges(System.Data.DataRowState.Modified);
            //deletedTable = dt.GetChanges(System.Data.DataRowState.Deleted);
            //addedTable = dt.GetChanges(System.Data.DataRowState.Added);
        }

        private void orgType_cb_TextChanged(object sender, RibbonControlEventArgs e)
        {
            RibbonComboBox cb = sender as RibbonComboBox;

            switch (cb.Text.ToLower())
            {
                case "production":
                    Auth.AuthUtil.baseUrl = "https://login.salesforce.com";
                    break;
                case "sandbox":
                default:
                    Auth.AuthUtil.baseUrl = "https://test.salesforce.com";
                    break;
            }
        }


        private UI.SObjectTreeViewControl treeView = new UI.SObjectTreeViewControl();
        private void btn_LoadTables_Click(object sender, RibbonControlEventArgs e)
        {
            sforce.Connection curConn = sforce.ConnectionManager.Instance.ActiveConnection;
            if (curConn == null)
            {
                curConn = sforce.ConnectionManager.Instance.Connections.First();

                curConn.Active();
            }

            if (sfClient == null)
            {
                sfClient = new sforce.SForceClient();
                sfClient.init(curConn.Session);
            }

            Cursor oldCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            //sfClient = new sforce.SForceClient();
            //sforce.SFSession sfSession = sforce.SFSession.GetSession();
            //bool isSucess = sfClient.login(sfSession);

            List<sforce.SObjectEntryBase> sobjectList = sfClient.getSObjects();

            //if (treeView == null)
            //{
            //    treeView = new UI.SObjectTreeViewControl();
            //}

            FufillTreeviewWithSObjectList(treeView, sobjectList);

            if (taskPane == null)
            {
                taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(treeView, "SObject List");
                taskPane.Visible = true;
            }
            // taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;


            // taskPane.VisibleChanged += TaskPane_VisibleChanged;

            Cursor.Current = oldCursor;

            btn_ShowHideTaskPane.Enabled = true;
        }

        private void FufillTreeviewWithSObjectList(UI.SObjectTreeViewControl treeView, List<sforce.SObjectEntryBase> sobjectList)
        {
            treeView.tv_sobjs.BeginUpdate();

            treeView.tv_sobjs.Nodes.Clear();

            treeView.tv_sobjs.Nodes.Add("SObjects");
            treeView.tv_sobjs.Nodes.Add("Custom Settings");
            foreach (sforce.SObjectEntry item in sobjectList)
            {
                // TreeNode node = new UI.SObjectNodeBase(item.Name, item.Label, sfClient);
                // TreeNode node = new UI.SObjectNodeBase(item, sfClient);
                // node.Collapse();
                // treeView.tv_sobjs.Nodes[0].Nodes.Add(node);

                UI.SObjectNode node = new UI.SObjectNode(item, null);
                if (item.IsCustomSetting)
                {
                    treeView.tv_sobjs.Nodes[1].Nodes.Add(node);
                }
                else
                {
                    treeView.tv_sobjs.Nodes[0].Nodes.Add(node);
                }
            }

            treeView.tv_sobjs.NodeMouseDoubleClick += Tv_sobjs_NodeMouseDoubleClick;
            treeView.tv_sobjs.NodeMouseClick += Tv_sobjs_NodeMouseClick;
            treeView.tv_sobjs.EndUpdate();
        }

        private void Tv_sobjs_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenuStrip rightClickMenu = new ContextMenuStrip();
                rightClickMenu.Items.Add("Reload", null, (o, ev) =>
                            {
                                ToolStripItem item = o as ToolStripItem;
                                if (item != null)
                                {
                                    ContextMenuStrip cxtMenuStrip = item.Owner as ContextMenuStrip;

                                    if (cxtMenuStrip != null)
                                    {
                                        var obj = cxtMenuStrip.SourceControl;
                                        TreeView tv = obj as TreeView;

                                        TreeNode node = tv.GetNodeAt(e.X, e.Y);

                                    }
                                }
                                MessageBox.Show(o.ToString() + "====" + ev.ToString());

                            });
                rightClickMenu.Show(sender as Control, new System.Drawing.Point(e.X, e.Y));
                return;
            }
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

        private bool updateOrgList(sforce.Connection conn)
        {
            RibbonDropDownItem newItem = this.dropDown_TargetOrg.Items.FirstOrDefault(item => item.Label == conn.InstanceName);
            if (newItem != null)
            {
                conn = sforce.ConnectionManager.Instance.FindConnection(conn.InstanceName);
            }
            else
            {
                newItem = Factory.CreateRibbonDropDownItem();
                newItem.Label = conn.InstanceName;

                this.dropDown_TargetOrg.Items.Add(newItem);
            }

            // conn.Active();
            // this.dropDown_org.SelectedItem = newItem;

            if (this.dropDown_TargetOrg.Items.Count == 1)
            {
                conn.Active((con) => {
                    if (sfClient == null)
                    {
                        sfClient = new sforce.SForceClient();
                    }

                    sfClient.init(con.Session);
                });

            }

            // enable buttons
            // this.dropDown_TargetOrg.Enabled = true;
            this.btn_LoadTables.Enabled = true;
            this.btn_ShowHideTaskPane.Enabled = true;
            this.btn_loadData.Enabled = true;
            this.btn_CommitChanges.Enabled = true;
            this.btn_CopySelection.Enabled = true;

            return true;
        }

        private void dropDown_TargetOrg_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            RibbonDropDown dropDown = sender as RibbonDropDown;
            if (dropDown == null)
            {
                return;
            }

            sforce.ConnectionManager.Instance.ActiveConnection.Deactive();

            string orgName = dropDown.SelectedItem.Label;
            sforce.Connection conn = sforce.ConnectionManager.Instance.FindConnection(orgName);
            conn.Active((con) => {
                if (sfClient == null)
                {
                    sfClient = new sforce.SForceClient();
                }

                sfClient.init(con.Session);

                FufillTreeviewWithSObjectList(this.treeView, con.SObjects);
            });
        }

        private bool apiVersion_Changed(string version)
        {
            double versionNum = 0;
            bool ret = double.TryParse(version, out versionNum);
            if (ret)
            {
                Auth.AuthUtil.apiVersion = (int)versionNum;
                // this.editbox_APIVersion.Text = string.Format("{0}.0", Auth.AuthUtil.apiVersion);


                //if (sfClient == null)
                //{
                //    sfClient = new sforce.SForceClient();

                //    if (sforce.ConnectionManager.Instance.ActiveConnection != null)
                //    {
                //        sfClient.init(sforce.ConnectionManager.Instance.ActiveConnection.Session);
                //    }
                //}

                if (sfClient != null && sforce.ConnectionManager.Instance.ActiveConnection != null)
                {
                    sfClient.init(sforce.ConnectionManager.Instance.ActiveConnection.Session);
                }

                return true;
            }

            return false;
        }

        //private void apiVersion_TextChanged(object sender, RibbonControlEventArgs e)
        //{
        //    apiVersion_Changed(this.editbox_APIVersion.Text);
        //}

        private void btn_CopySelection_Click(object sender, RibbonControlEventArgs e)
        {
            System.Data.DataTable dt = (System.Data.DataTable)ds.Tables[Globals.ThisAddIn.Application.ActiveSheet.Name];

            System.Data.DataTable updatedTable = dt.GetChanges(System.Data.DataRowState.Modified);
            System.Data.DataTable deletedTable = dt.GetChanges(System.Data.DataRowState.Deleted);
            System.Data.DataTable addedTable = dt.GetChanges(System.Data.DataRowState.Added);

            Tools.Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveSheet);
            Tools.ListObject listObj = Globals.Factory.GetVstoObject(sheet.ListObjects.Item[sheet.Name]);

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

            // List<int> indexList = new List<int>();
            foreach (Interop.Range item in ((Interop.Range)Globals.ThisAddIn.Application.Selection).Rows)
            {
                // indexList.Add(item.Row);

                // dt.Rows[item.Row - 2].SetModified(); // 2 = header + vsto starts from 1
                dt.Rows[item.Row - 2].SetAdded(); // 2 = header + vsto starts from 1
            }


            // updatedTable = dt.GetChanges(System.Data.DataRowState.Added);
            // sfClient.doUpdate(updatedTable);
            sfClient.doUpdate(dt);

            dt.AcceptChanges();

            //foreach (Interop.Range range in ranges)
            //{

            //}

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
    }
}
