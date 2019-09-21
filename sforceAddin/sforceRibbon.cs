using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using sforceAddin.SFDC;
using System.Net;

namespace sforceAddin
{
    public partial class sforceRibbon
    {
        private void sforceRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_login_Click(object sender, RibbonControlEventArgs e)
        {
            string userName = "";
            string password = "@";
            string secuToken = "";

            sforce.SForceClient sfClient = new sforce.SForceClient();
            bool isSucess = sfClient.login(userName, password, secuToken);

            if (!isSucess)
            {
                return;
            }

            List<sforce.SObjectEntry> sobjectList = sfClient.getSObjects();
        }
    }
}
