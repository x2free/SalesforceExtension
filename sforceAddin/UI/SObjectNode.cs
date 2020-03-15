using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using sforceAddin.sforce;
using Interop = Microsoft.Office.Interop.Excel;
using Tools = Microsoft.Office.Tools.Excel;

namespace sforceAddin.UI
{
    abstract class SObjectNodeBase : TreeNode
    {
        /// <summary>
        /// SForceClient object, to integration with Salesforce
        /// </summary>
        // protected sforce.SForceClient sfClient;
        /// <summary>
        /// Sub nodes, eg, fields of an SOject
        /// </summary>
        /// private List<SObjectNodeBase> subNodes;

        /// <summary>
        /// SObject binding to this node
        /// </summary>
        protected sforce.SObjectEntryBase sobjEntry;

        /// <summary>
        /// Parent node
        /// </summary>
        protected TreeNode parent;

        //public SObjectNodeBase(String name, String label, sforce.SForceClient sfClient)
        //{
        //    this.Name = name;
        //    this.Text = label;

        //    this.sfClient = sfClient;
        //}

        public SObjectNodeBase(sforce.SObjectEntryBase sobj/*, sforce.SForceClient sfClient*/, TreeNode parent)
        {
            this.Name = sobj.Name;
            this.Text = sobj.Label;
            this.parent = parent;

            // this.sfClient = sfClient;

            //if (sobj is sforce.SObjectEntry) {

            //}

            this.sobjEntry = sobj;
        }

        /// <summary>
        /// Determine if <sobjeEntry> has sub entrys
        /// </summary>
        abstract public void dbClick();

    }

    class SObjectNode : SObjectNodeBase
    {
        public SObjectNode(sforce.SObjectEntryBase sobj/*, sforce.SForceClient sfClient*/, TreeNode parent)
            : base(sobj/*, sfClient*/, parent) { }


        public override void dbClick()
        {
            if (this.Nodes != null && this.Nodes.Count > 0)
            {
                return;
            }

            var nodes = sobjEntry.getChildren();
            foreach (var item in nodes)
            {
                this.Nodes.Add(new FieldNode(item, this));
            }
        }
    }

    class FieldNode : SObjectNodeBase
    {
        public FieldNode(sforce.SObjectEntryBase sobj/*, sforce.SForceClient sfClient*/, TreeNode parent)
            : base(sobj/*, sfClient*/, parent)
        {
        }

        public override void dbClick()
        {
            // string tableName = parent.Text;
            string tableName = parent.Name;

            Microsoft.Office.Interop.Excel.Worksheet iWorksheet = null;
            // Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            // Microsoft.Office.Interop.Excel.Worksheets sheets = (Microsoft.Office.Interop.Excel.Worksheets)excelApp.Worksheets;

            //Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            //Microsoft.Office.Interop.Excel.Workbook wb = excelApp.ActiveWorkbook;
            Interop.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Interop.Sheets sheets = activeWorkbook.Sheets;

            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {
                if (String.Equals(tableName, sheet.Name, StringComparison.InvariantCultureIgnoreCase))
                {
                    iWorksheet = sheet;
                    break;
                }
            }

            if (iWorksheet == null)
            {
                iWorksheet = activeWorkbook.Sheets.Add();
                iWorksheet.Name = tableName;
            }

            iWorksheet.Activate();


            Tools.Worksheet worksheet = Globals.Factory.GetVstoObject(iWorksheet);

            //Microsoft.Office.Interop.Excel.QueryTable workTable = null;
            //var tables = worksheet.QueryTables;

            //foreach (Microsoft.Office.Interop.Excel.QueryTable  table in tables)
            //{
            //    if (String.Equals(tableName, table.Name, StringComparison.InvariantCultureIgnoreCase))
            //    {
            //        workTable = table;
            //        break;
            //    }
            //}

            //worksheet.QueryTables.Add()

            //Microsoft.Office.Interop.Excel.Range tableRange = worksheet.Range[worksheet.Cells[1, 0], worksheet.Cells[1, 2]];
            //tableRange.Name = tableName;
            //tableRange.Show();

            //Microsoft.Office.Interop.Excel.Range tableRange = worksheet.UsedRange;
            //if (tableRange == null)
            //{
            //    tableRange = worksheet.Range["A1", "A3"];
            //    tableRange.Name = tableName;
            //}

            //var r = worksheet.UsedRange;
            //int c = r.Cells.Count;
            //r.Value2 = "adadad";
            //r.Value = "111";
            // c = r.Cells.Count;

            //Microsoft.Office.Interop.Excel.Range tableRange = worksheet.UsedRange;
            //Microsoft.Office.Interop.Excel.ListObject listObj = worksheet.ListObjects.Add(Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange, tableRange,
            //                                     Type.Missing, Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes);
            //listObj.Name = tableName;
            //listObj.TableStyle = "TableStyleDark10";

            //Microsoft.Office.Interop.Excel.ListRow row = listObj.ListRows.AddEx(Type.Missing, Type.Missing);

            Microsoft.Office.Interop.Excel.ListObject listObj = null;

            foreach (Microsoft.Office.Interop.Excel.ListObject item in iWorksheet.ListObjects)
            {
                if (String.Equals(tableName, item.Name, StringComparison.InvariantCultureIgnoreCase))
                {
                    listObj = item;
                    break;
                }
            }

            if (listObj == null)
            {
                //// Microsoft.Office.Interop.Excel.Range curRange = worksheet.Cells.CurrentRegion;
                Microsoft.Office.Interop.Excel.Range curRange = Globals.ThisAddIn.Application.ActiveCell;
                // curRange.Name = this.Name;
                curRange.Name = string.Format("{0}.{1}", parent.Name, this.Name);
                // curRange.Value = this.Name;
                curRange.Value2 = this.Text;

                listObj = iWorksheet.ListObjects.Add(Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange, curRange/*worksheet.UsedRange*/,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes);

                // listObj.DisplayName = parent.Name;
                listObj.Name = tableName;
                listObj.TableStyle = "TableStyleMedium23";
                Microsoft.Office.Tools.Excel.ListObject hostedListObj = Globals.Factory.GetVstoObject(listObj);
                hostedListObj.Change += ListObject_Change;
                hostedListObj.OriginalDataRestored += HostedListObj_OriginalDataRestored;

                //listObj = worksheet.Controls.AddListObject(Globals.ThisAddIn.Application.ActiveCell, parent.Name).InnerObject;
                //// field header
                //Microsoft.Office.Interop.Excel.Range headerRange = listObj.HeaderRowRange.Cells[1, listObj.ListColumns.Count];
                //// headerRange.Name = this.Name;
                //headerRange.Name = string.Format("{0}.{1}", parent.Name, this.Name);
                //// headerRange.Value = this.Name;
                //headerRange.Value2 = this.Text;

            }
            else
            {
                Microsoft.Office.Tools.Excel.ListObject hostedListObj = Globals.Factory.GetVstoObject(listObj);

                // To add a column, must disconnect from the binded datasource, or the column cannot be added,
                // so that the consequent operation on this column will fail with exception.
                if (hostedListObj.DataSource != null)
                {
                    hostedListObj.Disconnect();
                }
                int cnt = hostedListObj.ListColumns.Count;
                Interop.ListColumn column = hostedListObj.ListColumns.Add(cnt + 1);
                column.Range.NumberFormat = "@"; // format as string
                cnt = hostedListObj.ListColumns.Count;
                column = hostedListObj.ListColumns[cnt];

                // int cnt = listObj.ListColumns.Count;
                // Microsoft.Office.Interop.Excel.ListColumn column = listObj.ListColumns.Add(cnt + 1);
                // cnt = listObj.ListColumns.Count;
                // column.Name = this.Name; // throwing exception if loaded data then add column again.
                column.Name = string.Format("{0}_{1}", parent.Name, this.Name);
                // Microsoft.Office.Interop.Excel.Range r = column.Range; // this won't get the header, it gets 2nd row
                //r.Value = this.Name;
                // r.Value2 = this.Text;

                // field header
                Microsoft.Office.Interop.Excel.Range headerRange = listObj.HeaderRowRange.Cells[1, listObj.ListColumns.Count];
                // headerRange.Name = this.Name;
                headerRange.Name = string.Format("{0}.{1}", parent.Name, this.Name);
                // headerRange.Value = this.Name;
                headerRange.Value2 = this.Text;

                //Microsoft.Office.Interop.Excel.Range curRange = worksheet.Cells[1, listObj.ListColumns.Count];
                //curRange.Value = this.Name;
                //curRange.Value2 = this.Text;

                //listObj.ListColumns.Add(curRange);
            }

            // remove this node once add to sheet
            // parent.Nodes.Remove(this);
        }

        private void HostedListObj_OriginalDataRestored(object sender, Tools.OriginalDataRestoredEventArgs e)
        {
            // throw new NotImplementedException();
        }

        private void ListObject_Change(Interop.Range targetRange, Tools.ListRanges changedRanges)
        {
            // throw new NotImplementedException();
            var entireColAddres = targetRange.EntireColumn.Address.Count();
            var entireRowAddres = targetRange.EntireRow.Address;
            var cellAddress = targetRange.Address.Count();

            bool isColDeleting = entireColAddres == cellAddress;

            if (changedRanges == (Tools.ListRanges.DataBodyRange | Tools.ListRanges.HeaderRowRange))
            {
                Tools.ListObject hostedListObj = Globals.Factory.GetVstoObject(targetRange.ListObject);
                if (hostedListObj.DataSource != null)
                {
                    System.Data.DataTable dt = hostedListObj.DataSource as System.Data.DataTable;
                    if (dt != null)
                    {
                        int count1 = hostedListObj.ListColumns.Count;
                        int count2 = dt.Columns.Count;

                        if (count1 < count2)
                        {
                            hostedListObj.Disconnect();
                        }
                    }
                }
                //if (hostedListObj.DataSource != null)
                //{
                //    hostedListObj.Disconnect();
                //}
            }
        }
    }
}
