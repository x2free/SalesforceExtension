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

        private void sforceRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_login_Click(object sender, RibbonControlEventArgs e)
        {
            Cursor currentCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            string userName = "@163.com";
            string password = "@";
            string secuToken = "";

            sforce.SForceClient sfClient = new sforce.SForceClient();
            bool isSucess = sfClient.login(userName, password, secuToken);

            if (!isSucess)
            {
                return;
            }

            List<sforce.SObjectEntry> sobjectList = sfClient.getSObjects();

            UI.sforceListViewControl lvControl = new UI.sforceListViewControl();
            foreach (var item in sobjectList)
            {
                lvControl.listview_sobjs.Items.Add(String.Format("{0}({1})", item.Label, item.Name));
            }
            lvControl.listview_sobjs.AutoResizeColumns(System.Windows.Forms.ColumnHeaderAutoResizeStyle.ColumnContent);

            taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(lvControl, "SObject List");
            taskPane.Visible = true;
            taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;

            // taskPane.VisibleChanged += TaskPane_VisibleChanged;

            Cursor.Current = currentCursor;

            btn_taskPane.Enabled = true;
        }

        private void btn_taskPane_Click(object sender, RibbonControlEventArgs e)
        {
            if (taskPane == null)
            {
                btn_taskPane.Enabled = false;
                return;
            }

            taskPane.Visible = !taskPane.Visible;
            UI.sforceListViewControl lvControl = taskPane.Control as UI.sforceListViewControl;
            if (lvControl == null)
            {
                lvControl.AutoSize = true;
                // btn_taskPane.Enabled = false;
                return;
            }

        }
    }
}
