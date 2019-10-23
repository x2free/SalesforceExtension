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

        public System.Data.DataTable execQuery(string query)
        {
            QueryResult ret = this.sfSvc.query(query);

            System.Data.DataTable dt = new System.Data.DataTable("Table01");

            if (ret != null && ret.records.Count<sObject>() > 0)
            {
                sObject rec = ret.records.First<sObject>();
                // create column info based on 1st row
                foreach (System.Xml.XmlElement col in rec.Any)
                {
                    if (col.FirstChild != null)
                    {
                        switch (col.FirstChild.NodeType)
                        {
                            case System.Xml.XmlNodeType.None:
                                break;
                            case System.Xml.XmlNodeType.Element:
                                break;
                            case System.Xml.XmlNodeType.Attribute:
                                break;
                            case System.Xml.XmlNodeType.Text:
                                dt.Columns.Add(col.LocalName, typeof(string));
                                break;
                            case System.Xml.XmlNodeType.CDATA:
                                break;
                            case System.Xml.XmlNodeType.EntityReference:
                                break;
                            case System.Xml.XmlNodeType.Entity:
                                break;
                            case System.Xml.XmlNodeType.ProcessingInstruction:
                                break;
                            case System.Xml.XmlNodeType.Comment:
                                break;
                            case System.Xml.XmlNodeType.Document:
                                break;
                            case System.Xml.XmlNodeType.DocumentType:
                                break;
                            case System.Xml.XmlNodeType.DocumentFragment:
                                break;
                            case System.Xml.XmlNodeType.Notation:
                                break;
                            case System.Xml.XmlNodeType.Whitespace:
                                break;
                            case System.Xml.XmlNodeType.SignificantWhitespace:
                                break;
                            case System.Xml.XmlNodeType.EndElement:
                                break;
                            case System.Xml.XmlNodeType.EndEntity:
                                break;
                            case System.Xml.XmlNodeType.XmlDeclaration:
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        dt.Columns.Add(col.LocalName, typeof(string));
                    }
                }
            }

            foreach (sObject rec in ret.records)
            {
                System.Data.DataRow dr = dt.NewRow();

                foreach (System.Xml.XmlElement col in rec.Any)
                {
                    dr[col.LocalName] = col.InnerText;
                }

                dt.Rows.Add(dr);
            }

            return dt;
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
