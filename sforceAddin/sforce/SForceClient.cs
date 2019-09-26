using sforceAddin.SFDC;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace sforceAddin.sforce
{
    class SForceClient
    {
        private SforceService sfSvc;
        private String oldAuthUrl;
        public String serverUrl;

        public List<SObjectEntry> sobjectList;

        public bool login(String userName, String password, String securityToken)
        {
            // To enable SSL/TLS
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

            sfSvc = new SFDC.SforceService();

            // sfSvc.Url = "https://login.salesforce.com"; // Do not set this, use default SOAP login URL in config file
            SFDC.LoginResult lr;
            try
            {
                lr = sfSvc.login(userName, password + securityToken);

                // save old authenticaton endpoint
                oldAuthUrl = sfSvc.Url;

                // Get the session ID from the login result and set it for the
                // session header that will be used for all subsequent calls.
                sfSvc.SessionHeaderValue = new SFDC.SessionHeader();
                sfSvc.SessionHeaderValue.sessionId = lr.sessionId;

                sfSvc.Url = lr.serverUrl;
                this.serverUrl = sfSvc.Url;

                return true;
            }
            catch (Exception ex)
            {
            }

            return false;
        }

        public List<SObjectEntryBase> getSObjects()
        {
            List<sforce.SObjectEntryBase> sobjects = new List<sforce.SObjectEntryBase>();
            // get SObjects
            // Make the describeGlobal() call 
            DescribeGlobalResult globalDesc = sfSvc.describeGlobal();

            // Get the sObjects from the describe global result
            DescribeGlobalSObjectResult[] sObjResults = globalDesc.sobjects;

            foreach (var sobj in globalDesc.sobjects)
            {
                if (sobj.queryable && sobj.createable && sobj.updateable && sobj.deletable)
                {
                    sobjects.Add(new SObjectEntry(sobj.name, sobj.label, sobj.keyPrefix, sobj.custom, sobj.customSetting, this, sobj.labelPlural));
                }
            }

            sobjects.Sort();

            return sobjects;
        }

        public List<sforce.SObjectEntryBase> describeSObject(SObjectEntryBase sobj)
        {
            List<sforce.SObjectEntryBase> fields = new List<SObjectEntryBase>();

            DescribeSObjectResult result =  this.sfSvc.describeSObject(sobj.Name);
            if (result == null)
            {
                return fields;
            }

            foreach (var field in result.fields)
            {
                fields.Add(new FieldEntry(field.name, field.label, field.custom, this, sobj));
            }

            // var relation = result.childRelationships;
            // var cr = result.childRelationships;

            return fields;
        }

        private bool IsValidSession
        {
            get
            {
                // this.sfSvc
                return false;
            }
        }
    }
}
