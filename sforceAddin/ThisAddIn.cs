using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

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
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
        }
        
        #endregion
    }
}
