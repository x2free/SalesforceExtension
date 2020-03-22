using sforceAddin.SFDC;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace sforceAddin.sforce
{
    class SForceClient
    {
        private SforceService sfSvc;
        private String oldAuthUrl;
        public String serverUrl;

        public List<SObjectEntry> sobjectList;

        public bool init(sforce.SFSession sfSession)
        {
            if (sfSvc == null)
            {
                sfSvc = new SFDC.SforceService();
                sfSvc.SessionHeaderValue = new SessionHeader();
            }

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

        public void logout()
        {
            sfSvc.logout();
        }

        public List<SObjectEntryBase> getSObjects(bool force = false)
        {
            // cache objects
            List<sforce.SObjectEntryBase> sobjects = ConnectionManager.Instance.ActiveConnection.SObjects == null
                        ? new List<SObjectEntryBase>() : ConnectionManager.Instance.ActiveConnection.SObjects;
            if (!force && sobjects.Any())
            {
                return sobjects;
            }

            sobjects.Clear();

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

            sobjects.Sort((a, b) => { return string.Compare(a.Label, b.Label); });
            ConnectionManager.Instance.ActiveConnection.SObjects = sobjects;

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
                //throw new Exception("No data loaded!");
                return null;
            }

            // column
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

            if (ret.records.Count<sObject>() > 0)
            {
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
                    // string fieldName = string.Format("{0}_{1}", tableName, col.LocalName);
                    string fieldName = col.LocalName;
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

            do
            {
                foreach (sObject rec in ret.records)
                {
                    System.Data.DataRow dr = dt.NewRow();

                    foreach (System.Xml.XmlElement col in rec.Any)
                    {
                        // dr[col.LocalName] = col.InnerText;
                        // string fieldName = string.Format("{0}_{1}", tableName, col.LocalName);
                        string fieldName = col.LocalName;
                        dr[fieldName] = col.InnerText;
                    }

                    dt.Rows.Add(dr);
                }

                if (ret.done)
                {
                    break;
                }

                ret = sfSvc.queryMore(ret.queryLocator);
            } while (true);

            dt.AcceptChanges();

            return dt;
        }

        public void doUpdate(System.Data.DataTable table)
        {
            System.Data.DataTable updatedTable = table.GetChanges(System.Data.DataRowState.Modified);
            System.Data.DataTable deletedTable = table.GetChanges(System.Data.DataRowState.Deleted);
            System.Data.DataTable addedTable = table.GetChanges(System.Data.DataRowState.Added);

            //DataTable upsertTable = table.GetChanges(DataRowState.Modified | DataRowState.Added);

            List<sObject> upsertList = new List<sObject>();

            // refer to https://developer.salesforce.com/forums/?id=906F00000008sJ3IAI
            // and https://developer.salesforce.com/forums/?id=906F00000008sErIAI
            // to create objects

            XmlDocument doc = new XmlDocument();

            if (updatedTable != null)
            {
                foreach (System.Data.DataRow row in updatedTable.Rows)
                {
                    IEnumerable<DataColumn> changedCols = DataRowExtensions.GetChangedColumns(row);
                    sObject obj = new sObject();
                    obj.type = updatedTable.TableName;
                    bool isChanged = false;

                    List<XmlElement> fieldElements = new List<XmlElement>();
                    List<String> fields2Null = null;

                    foreach (System.Data.DataColumn column in updatedTable.Columns)
                    {
                        var oldValue = row[column, DataRowVersion.Original];
                        var curValue = row[column, DataRowVersion.Current];
                        object fieldValue = null;

                        // DateTime? dt = row.Field<DateTime?>(column);

                        XmlElement field = null;
                        // if (row.IsNull(column))
                        if (curValue == DBNull.Value)
                        {
                            if (oldValue == DBNull.Value || string.IsNullOrEmpty(oldValue as string)) // empty but not changed, ignore this field for this row
                            {
                                continue;
                            }
                            else // this field gets deleted
                            {
                                // field = doc.CreateElement(column.ColumnName);
                                // field.InnerText = null;
                                // fieldValue = string.Empty;
                                // fieldValue = DBNull.Value;

                                if (fields2Null == null)
                                {
                                    fields2Null = new List<string>();
                                }
                                fields2Null.Add(column.ColumnName);

                                isChanged |= true;
                            }
                        }
                        else if ("id".Equals(column.ColumnName, StringComparison.InvariantCultureIgnoreCase))
                        {
                            // field = doc.CreateElement(column.ColumnName);
                            // field.InnerText = (string)curValue;

                            fieldValue = curValue;
                        }
                        else if (curValue != oldValue)
                        {
                            // field = doc.CreateElement(column.ColumnName);
                            // field.InnerText = (string)row[column];
                            // field.InnerText = (string)curValue;
                            fieldValue = curValue;

                            isChanged |= true;
                        }

                        // if (fieldValue != null || isChanged)
                        if (fieldValue != null)
                        {
                            field = doc.CreateElement(column.ColumnName);
                            field.InnerText = fieldValue as string;

                            fieldElements.Add(field);
                        }
                    }

                    if (isChanged)
                    {
                        obj.Any = fieldElements.ToArray();
                        obj.fieldsToNull = fields2Null == null ? null : fields2Null.ToArray();
                        upsertList.Add(obj);
                    }
                }
            }

            if (addedTable != null)
            {
                foreach (DataRow row in addedTable.Rows)
                {
                    sObject obj = new sObject();
                    obj.type = addedTable.TableName;

                    List<XmlElement> fieldElements = new List<XmlElement>();

                    foreach (System.Data.DataColumn column in addedTable.Columns)
                    {
                        var curValue = row[column, DataRowVersion.Current];

                        XmlElement field = null;
                        // if (row.IsNull(column))
                        if (curValue == DBNull.Value || string.IsNullOrEmpty(curValue as string) || "id".Equals(column.ColumnName, StringComparison.InvariantCultureIgnoreCase))
                        {
                            continue;
                        }
                        //else if ("id".Equals(column.ColumnName, StringComparison.InvariantCultureIgnoreCase))
                        //{
                        //    field = doc.CreateElement(column.ColumnName);
                        //    field.InnerText = null;
                        //}
                        else
                        {
                            field = doc.CreateElement(column.ColumnName);
                            field.InnerText = (string)curValue;
                        }

                        fieldElements.Add(field);
                    }

                    obj.Any = fieldElements.ToArray();
                    upsertList.Add(obj);
                }
            }

            if (upsertList.Count > 0)
            {
                UpsertResult[] results =  sfSvc.upsert("Id", upsertList.ToArray());

                foreach (UpsertResult ret in results)
                {
                    if (!ret.success)
                    {
                        Error[] errors = ret.errors;
                    }

                }
            }

            List<string> idsToDelete = new List<string>();

            if (deletedTable != null)
            {
                foreach (DataRow row in deletedTable.Rows)
                {
                    string id = row["Id", DataRowVersion.Original] as string;

                    if (!string.IsNullOrEmpty(id))
                    {
                        idsToDelete.Add(id);
                    }
                }
            }

            if (idsToDelete.Count > 0)
            {
                DeleteResult[] results = sfSvc.delete(idsToDelete.ToArray());

                foreach (DeleteResult result in results)
                {
                    if (!result.success)
                    {
                        Error[] errors = result.errors;
                    }
                }
            }
        }

        private static bool hasCellChanged(DataRow row, DataColumn col)
        {
            if (!row.HasVersion(DataRowVersion.Original))
            {
                // Row has been added. All columns have changed. 
                return true;
            }
            if (!row.HasVersion(DataRowVersion.Current))
            {
                // Row has been removed. No columns have changed.
                return false;
            }
            var originalVersion = row[col, DataRowVersion.Original];
            var currentVersion = row[col, DataRowVersion.Current];
            if (originalVersion == DBNull.Value && currentVersion == DBNull.Value)
            {
                return false;
            }
            else if (originalVersion != DBNull.Value && currentVersion != DBNull.Value)
            {
                return !originalVersion.Equals(currentVersion);
            }
            return true;
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

    public static class DataRowExtensions
    {
        private static bool hasCellChanged(DataRow row, DataColumn col)
        {
            if (!row.HasVersion(DataRowVersion.Original))
            {
                // Row has been added. All columns have changed. 
                return true;
            }
            if (!row.HasVersion(DataRowVersion.Current))
            {
                // Row has been removed. No columns have changed.
                return false;
            }
            var originalVersion = row[col, DataRowVersion.Original];
            var currentVersion = row[col, DataRowVersion.Current];
            if (originalVersion == DBNull.Value && currentVersion == DBNull.Value)
            {
                return false;
            }
            else if (originalVersion != DBNull.Value && currentVersion != DBNull.Value)
            {
                return !originalVersion.Equals(currentVersion);
            }

            return true;
        }

        public static IEnumerable<DataColumn> GetChangedColumns(this DataRow row)
        {
            return row.Table.Columns.Cast<DataColumn>()
                .Where(col => hasCellChanged(row, col));
        }

        public static IEnumerable<DataColumn> GetChangedColumns(this IEnumerable<DataRow> rows)
        {
            return rows.SelectMany(row => row.GetChangedColumns())
                .Distinct();
        }

        public static IEnumerable<DataColumn> GetChangedColumns(this DataTable table)
        {
            return table.GetChanges().Rows
                .Cast<DataRow>()
                .GetChangedColumns();
        }
    }
}
