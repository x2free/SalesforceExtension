using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using sforceAddin.sforce;
using Interop = Microsoft.Office.Interop.Excel;
using Tools = Microsoft.Office.Tools.Excel;
using System.Data;

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
        // protected List<SObjectNodeBase> subNodes;

        /// <summary>
        /// SObject binding to this node
        /// </summary>
        public sforce.SObjectEntryBase SObjEntry { get; private set; }
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

        private static ImageList imgList = null;
        public static ImageList ImgList {
            get
            {
                if (imgList == null)
                {
                    imgList = new ImageList();
                    // imgList.Images.Add(System.Drawing.Icon.ExtractAssociatedIcon(@"Resources\Required Icon.ico"));
                    // imgList.Images.Add("RedStar", System.Drawing.Image.FromFile("sforceAddin.Resources.RedStar.png"));
                    imgList.Images.Add("NonImg", Properties.Resources.NonImg);
                    imgList.Images.Add("RedStar4P", Properties.Resources.RedStar4P);
                    imgList.Images.Add("RedStar6P", Properties.Resources.RedStar6P);
                }

                return imgList;
            }
        }

        public SObjectNodeBase(sforce.SObjectEntryBase sobj/*, sforce.SForceClient sfClient*/, TreeNode parent)
        {
            this.Name = sobj.Name;
            this.Text = string.Format("{0}({1})", sobj.Label, sobj.Name);
            this.parent = parent;

            // this.sfClient = sfClient;

            //if (sobj is sforce.SObjectEntry) {

            //}

            this.SObjEntry = sobj;
        }

        /// <summary>
        /// Determine if <sobjeEntry> has sub entrys
        /// </summary>
        abstract public void LoadNode(bool force = false);

        public static SObjectNodeBase CreateNode(sforce.SObjectEntryBase entry, TreeNode parent)
        {
            SObjectNodeBase node = null;
            if (entry is sforce.SObjectEntry)
            {
                node = new SObjectNode(entry, parent);
            }
            else if (entry is sforce.FieldEntry)
            {
                node = new FieldNode(entry, parent);
            }

            return node;
        }
    }

    class SObjectNode : SObjectNodeBase
    {
        public SObjectNode(sforce.SObjectEntryBase sobj/*, sforce.SForceClient sfClient*/, TreeNode parent)
            : base(sobj/*, sfClient*/, parent) { }


        public override void LoadNode(bool force = false)
        {
            if (this.Nodes.Count != 0)
            {
                this.Nodes.Clear();
            }

            foreach (var entry in SObjEntry.LoadChildren(force))
            {
                this.Nodes.Add(new FieldNode(entry, this));
            }
        }
    }

    class FieldNode : SObjectNodeBase
    {
        public FieldNode(sforce.SObjectEntryBase sobj/*, sforce.SForceClient sfClient*/, TreeNode parent)
            : base(sobj/*, sfClient*/, parent)
        {
            bool isRequired = (sobj as sforce.FieldEntry).IsRequired;

            if (isRequired)
            {
                this.ImageKey = "RedStar4P";
                this.SelectedImageKey = "RedStar6P";
            }
        }

        public override void LoadNode(bool force = false)
        {
            // string tableName = parent.Text;
            string tableName = parent.Name;
            string sheetName = tableName;
            if (sheetName.Length >= 32)
            {
                sheetName = sheetName.Substring(0, 27) + "$" + sheetName.Length.ToString();
            }

            if (!SForceClient.Instance.SheetNameToTableNameMap.ContainsKey(sheetName))
            {
                SForceClient.Instance.SheetNameToTableNameMap.Add(sheetName, tableName);
            }

            Microsoft.Office.Interop.Excel.Worksheet iWorksheet = null;
            // Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            // Microsoft.Office.Interop.Excel.Worksheets sheets = (Microsoft.Office.Interop.Excel.Worksheets)excelApp.Worksheets;

            //Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            //Microsoft.Office.Interop.Excel.Workbook wb = excelApp.ActiveWorkbook;
            Interop.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Interop.Sheets sheets = activeWorkbook.Sheets;

            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {
                if (String.Equals(sheetName, sheet.Name, StringComparison.InvariantCultureIgnoreCase))
                {
                    iWorksheet = sheet;
                    break;
                }
            }

            if (iWorksheet == null)
            {
                iWorksheet = activeWorkbook.Sheets.Add();
                iWorksheet.Name = sheetName;
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

            Microsoft.Office.Tools.Excel.ListObject hostedListObj = null;
            DataTable dt = SForceClient.Instance.DataSet.Tables[tableName];
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
                hostedListObj = Globals.Factory.GetVstoObject(listObj);
                // hostedListObj.Change += ListObject_Change;
                hostedListObj.OriginalDataRestored += HostedListObj_OriginalDataRestored;
                hostedListObj.BeforeRightClick += HostedListObj_BeforeRightClick;
                hostedListObj.DataMemberChanged += HostedListObj_DataMemberChanged;
                hostedListObj.Change += HostedListObj_Change;

                //listObj = worksheet.Controls.AddListObject(Globals.ThisAddIn.Application.ActiveCell, parent.Name).InnerObject;
                //// field header
                //Microsoft.Office.Interop.Excel.Range headerRange = listObj.HeaderRowRange.Cells[1, listObj.ListColumns.Count];
                //// headerRange.Name = this.Name;
                //headerRange.Name = string.Format("{0}.{1}", parent.Name, this.Name);
                //// headerRange.Value = this.Name;
                //headerRange.Value2 = this.Text;

                if (dt != null)
                {
                    SForceClient.Instance.DataSet.Tables.Remove(dt);
                }

                dt = new DataTable(tableName);

                SForceClient.Instance.DataSet.Tables.Add(dt);
            }
            else
            {
                hostedListObj = Globals.Factory.GetVstoObject(listObj);

                // To add a column, must disconnect from the binded datasource, or the column cannot be added,
                // so that the consequent operation on this column will fail with exception.
                if (hostedListObj.DataSource != null)
                {
                    hostedListObj.Disconnect();
                }

                string headerRangeName = string.Format("{0}.{1}", parent.Name, this.Name);
                for (int i = 1; i <= listObj.ListColumns.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.Range range = listObj.HeaderRowRange.Cells[1, i];
                    if (string.Equals(headerRangeName, range.Name.Name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        range.Select();

                        return;
                    }
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

            // Do not add duplicate column for a table
            if (!dt.Columns.Contains(this.Name))
            {
                dt.Columns.Add(this.Name, typeof(string));
            }

            hostedListObj.SetDataBinding(dt);
            //hostedListObj.DataBoundFormat = Interop.XlRangeAutoFormat.xlRangeAutoFormat3DEffects1;
            //hostedListObj.DefaultDataSourceUpdateMode = DataSourceUpdateMode.OnPropertyChanged;
        }

        private void HostedListObj_Change(Interop.Range targetRange, Tools.ListRanges changedRanges)
        {
            // throw new NotImplementedException();
        }

        private void HostedListObj_DataMemberChanged(object sender, EventArgs e)
        {
            // throw new NotImplementedException();
        }

        private void HostedListObj_BeforeRightClick(Interop.Range Target, ref bool Cancel)
        {
            // throw new NotImplementedException();
        }

        private void HostedListObj_OriginalDataRestored(object sender, Tools.OriginalDataRestoredEventArgs e)
        {
            // throw new NotImplementedException();

            Microsoft.Office.Interop.Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            string tableName = null;
            SForceClient.Instance.SheetNameToTableNameMap.TryGetValue(sheet.Name, out tableName);

            Microsoft.Office.Interop.Excel.ListObject listObj = null;
            foreach (Microsoft.Office.Interop.Excel.ListObject obj in sheet.ListObjects)
            {
                // if (String.Equals(sheet.Name, obj.Name, StringComparison.InvariantCultureIgnoreCase))
                if (String.Equals(tableName, obj.Name, StringComparison.InvariantCultureIgnoreCase))
                {
                    listObj = obj;
                    break;
                }
            }

            if (listObj == null)
            {
                return;
            }

            Microsoft.Office.Tools.Excel.ListObject hostListObject = Globals.Factory.GetVstoObject(listObj);

            if (hostListObject.DataSource != null)
            {
                hostListObject.Disconnect();
            }
        }

        private void ListObject_Change(Interop.Range targetRange, Tools.ListRanges changedRanges)
        {
            // throw new NotImplementedException();
            var entireColAddres = targetRange.EntireColumn.Address.Count();
            var entireRowAddres = targetRange.EntireRow.Address;
            var cellAddress = targetRange.Address.Count();

            // bool isColDeleting = (entireColAddres == cellAddress) && (changedRanges == (Tools.ListRanges.DataBodyRange | Tools.ListRanges.HeaderRowRange));
            bool isColumnsChanged = changedRanges == (Tools.ListRanges.DataBodyRange | Tools.ListRanges.HeaderRowRange);

            // if (isColDeleting)
            if (isColumnsChanged)
            {
                Tools.ListObject hostedListObj = Globals.Factory.GetVstoObject(targetRange.ListObject);
                if (hostedListObj.DataSource != null)
                {
                    hostedListObj.Disconnect();
                }
            }

            //if (changedRanges == (Tools.ListRanges.DataBodyRange | Tools.ListRanges.HeaderRowRange))
            //{
            //    Tools.ListObject hostedListObj = Globals.Factory.GetVstoObject(targetRange.ListObject);
            //    if (hostedListObj.DataSource != null)
            //    {
            //        System.Data.DataTable dt = hostedListObj.DataSource as System.Data.DataTable;
            //        if (dt != null)
            //        {
            //            int count1 = hostedListObj.ListColumns.Count;
            //            int count2 = dt.Columns.Count;

            //            if (count1 < count2)
            //            {
            //                hostedListObj.Disconnect();
            //            }
            //        }
            //    }
            //    //if (hostedListObj.DataSource != null)
            //    //{
            //    //    hostedListObj.Disconnect();
            //    //}
            //}
        }
    }
}
