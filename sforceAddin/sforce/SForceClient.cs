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

        public bool login(sforce.SFSession sfSession)
        {
            sfSvc = new SFDC.SforceService();
            sfSvc.SessionHeaderValue = new SessionHeader();
            sfSvc.SessionHeaderValue.sessionId = sfSession.SessionId;
            // sfSvc.Url = sfSession.InstanceUrl;
            sfSvc.Url = sfSession.SoapPartnerUrl;
            this.serverUrl = sfSession.SoapPartnerUrl;

            return true;
        }

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
                Console.Write(ex.ToString());
                throw ex;
            }

            //return false;
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

        public System.Data.DataTable execQuery(string query, string tableName, System.Data.DataTable dt)
        {
            QueryResult ret = this.sfSvc.query(query);

            if (ret == null || ret.records == null)
            {
                throw new Exception("No data loaded!");
            }

            bool isChanged = false;
            if (dt == null)
            {
                dt = new System.Data.DataTable(tableName);

                dt.ColumnChanged += Dt_ColumnChanged;
                // dt.ColumnChanging += Dt_ColumnChanging;
                dt.RowChanged += Dt_RowChanged;
                dt.RowDeleted += Dt_RowDeleted;

                isChanged = true;
            }

            if (ret.records.Count<sObject>() > 0) {
                sObject rec = ret.records.First<sObject>();

                if (dt.Columns.Count != rec.Any.Count())
                {
                    isChanged = true;
                }
            }

            if (isChanged)
            {
                dt.Clear();
                dt.Columns.Clear();
            }

            if (isChanged)
            {
                sObject rec = ret.records.First<sObject>();
                // create column info based on 1st row
                foreach (System.Xml.XmlElement col in rec.Any)
                {
                    string fieldName = string.Format("{0}_{1}", tableName, col.LocalName);
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
                                dt.Columns.Add(fieldName, typeof(string));
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
                        dt.Columns.Add(fieldName, typeof(string));
                    }
                }
            }

                // In case of that add/remove columns when reload
                // dt.Columns.Clear();

                // clear rows then re-bind them.
                dt.Rows.Clear();

            foreach (sObject rec in ret.records)
            {
                System.Data.DataRow dr = dt.NewRow();

                foreach (System.Xml.XmlElement col in rec.Any)
                {
                    // dr[col.LocalName] = col.InnerText;
                    string fieldName = string.Format("{0}_{1}", tableName, col.LocalName);
                    dr[fieldName] = col.InnerText;
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }

        public void doUpdate(System.Data.DataTable table)
        {
            System.Data.DataTable updatedTable = table.GetChanges(System.Data.DataRowState.Modified);
            System.Data.DataTable deletedTable = table.GetChanges(System.Data.DataRowState.Deleted);
            System.Data.DataTable addedTable = table.GetChanges(System.Data.DataRowState.Added);

            List<sObject> upsertList = new List<sObject>();

            // refer to https://developer.salesforce.com/forums/?id=906F00000008sJ3IAI
            // to create objects
            foreach (var item in updatedTable.Rows)
            {
                sObject obj = new sObject();
            }

            /*
             
col
{Element, Name="sf:Id"}
    Attributes: {System.Xml.XmlAttributeCollection}
    BaseURI: ""
    ChildNodes: {System.Xml.XmlChildNodes}
    FirstChild: {Text, Value="0016F00002cseUHQAY"}
    HasAttributes: false
    HasChildNodes: true
    InnerText: "0016F00002cseUHQAY"
    InnerXml: "0016F00002cseUHQAY"
    IsEmpty: false
    IsReadOnly: false
    LastChild: {Text, Value="0016F00002cseUHQAY"}
    LocalName: "Id"
    Name: "sf:Id"
    NamespaceURI: "urn:sobject.partner.soap.sforce.com"
    NextSibling: null
    NodeType: Element
    OuterXml: "<sf:Id xmlns:sf=\"urn:sobject.partner.soap.sforce.com\">0016F00002cseUHQAY</sf:Id>"
    OwnerDocument: {Document}
    ParentNode: null
    Prefix: "sf"
    PreviousSibling: null
    PreviousText: null
    SchemaInfo: {System.Xml.XmlName}
    Value: null
    Results View: Expanding the Results View will enumerate the IEnumerable
col.ChildNodes
{System.Xml.XmlChildNodes}
    Count: 1
    Results View: Expanding the Results View will enumerate the IEnumerable
col.ChildNodes[0]
{Text, Value="0016F00002cseUHQAY"}
    Attributes: null
    BaseURI: ""
    ChildNodes: {System.Xml.XmlChildNodes}
    Data: "0016F00002cseUHQAY"
    FirstChild: null
    HasChildNodes: false
    InnerText: "0016F00002cseUHQAY"
    InnerXml: ""
    IsReadOnly: false
    LastChild: null
    Length: 18
    LocalName: "#text"
    Name: "#text"
    NamespaceURI: ""
    NextSibling: null
    NodeType: Text
    OuterXml: "0016F00002cseUHQAY"
    OwnerDocument: {Document}
    ParentNode: {Element, Name="sf:Id"}
    Prefix: ""
    PreviousSibling: null
    PreviousText: null
    SchemaInfo: {System.Xml.Schema.XmlSchemaInfo}
    Value: "0016F00002cseUHQAY"
    Results View: Expanding the Results View will enumerate the IEnumerable

             
             */
        }

        private void Dt_RowDeleted(object sender, System.Data.DataRowChangeEventArgs e)
        {

        }

        private void Dt_RowChanged(object sender, System.Data.DataRowChangeEventArgs e)
        {

        }

        private void Dt_ColumnChanging(object sender, System.Data.DataColumnChangeEventArgs e)
        {

        }

        private void Dt_ColumnChanged(object sender, System.Data.DataColumnChangeEventArgs e)
        {

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
