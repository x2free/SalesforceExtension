using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Tool = Microsoft.Office.Tools.Excel;
using Microsoft.Office.Core;
using System.Reflection;

namespace sforceAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //AppDomain.CurrentDomain.UnhandledException +=
            //    (s, ev) => { System.Windows.Forms.MessageBox.Show(ev.ToString(), "sforce Addin Unhandled Exception", System.Windows.Forms.MessageBoxButtons.OK); };

            //AppDomain.CurrentDomain.FirstChanceException +=
            //    (obj, ex) => {
            //            System.Windows.Forms.MessageBox.Show(ex.Exception.Message, "sforce Addin Unhandled Exception"
            //                    , System.Windows.Forms.MessageBoxButtons.OK);
            //    };

            // load all listobject/table into memory
            // assume that one sheet has only one list object
            //if (this.Application == null || this.Application.Worksheets == null) // cannot visit this.Application while starting, why?
            //{
            //    return;
            //}

            //foreach (Excel.Worksheet sheet in this.Application.Worksheets)
            //{
            //    if (sheet.ListObjects == null || sheet.ListObjects.Count == 0)
            //    {
            //        continue;
            //    }

            //    Excel.ListObject listObj = sheet.ListObjects.Item[1];

            //    System.Data.DataTable dt2 = new System.Data.DataTable(listObj.Name);
            //    foreach (Microsoft.Office.Interop.Excel.Range headerCell in listObj.HeaderRowRange.Cells)
            //    {
            //        string fieldName = headerCell.Name.Name.Substring(listObj.Name.Length + 1);
            //        dt2.Columns.Add(fieldName);
            //    }

            //    sforce.SForceClient.Instance.SheetNameToTableNameMap.Add(sheet.Name, listObj.Name);
            //    sforce.SForceClient.Instance.DataSet.Tables.Add(dt2);
            //}
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (sforce.SFSessionManager.Instance.Sessions == null || sforce.SFSessionManager.Instance.Sessions.Count() == 0)
            {
                return;
            }


            foreach (sforce.SFSession session in sforce.SFSessionManager.Instance.Sessions)
            {
                sforce.SForceClient.Instance.SetSession(session);
                sforce.SForceClient.Instance.Logout();
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            this.Application.SheetBeforeRightClick += OnSheetRightClick;
        }

        private void OnSheetRightClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            CommandBar cxtMenu = Globals.ThisAddIn.Application.CommandBars["Cell"];
            CommandBarButton cmdItem = (CommandBarButton)cxtMenu.FindControl(MsoControlType.msoControlButton, 0, "sfMenuItem", Missing.Value, Missing.Value);
            if (cmdItem == null)
            {
                // add the button
                cmdItem = (CommandBarButton)cxtMenu.Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, cxtMenu.Controls.Count, true);
                cmdItem.Caption = "sfMenuItems";
                cmdItem.BeginGroup = true;
                cmdItem.Tag = "sfMenuItem";
                cmdItem.Click += OnClickCxtMenu;
                //cmdItem.Visible = true;
                //cxtMenu.Reset();
            }

            //cxtMenu.ShowPopup();
        }

        private void OnClickCxtMenu(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            
        }

        #endregion
    }
}
