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
        private sforce.SFSession sfSession;
        private static SForceClient instance;
        // public String serverUrl;

        public System.Data.DataSet DataSet { get; private set; }
        public Dictionary<string, string> SheetNameToTableNameMap { get; private set; }

        public void SetSession(sforce.SFSession session)
        {
            this.sfSession = session;

            sfSvc.SessionHeaderValue.sessionId = sfSession.SessionId;
            // sfSvc.Url = sfSession.InstanceUrl;
            sfSvc.Url = sfSession.SoapPartnerUrl;
            // this.serverUrl = sfSession.SoapPartnerUrl;
        }

        private SForceClient()
        {
            sfSvc = new SFDC.SforceService();
            sfSvc.SessionHeaderValue = new SessionHeader();

            DataSet = new System.Data.DataSet();
            SheetNameToTableNameMap = new Dictionary<string, string>();

            DataSet.Tables.CollectionChanging += (o, e) =>
            {
                if (e.Action == System.ComponentModel.CollectionChangeAction.Add)
                {
                    DataTable dt = e.Element as DataTable;
                    if (dt == null)
                    {
                        return;
                    }

                    string sheetName = dt.TableName;
                    if (sheetName.Length > 32)
                    {
                        sheetName = sheetName.Substring(0, 27) + "$" + sheetName.Length.ToString();
                    }

                    // dt.DisplayExpression = sheetName;
                    // SheetNameToTableNameMap.Add(sheetName, dt.TableName);
                }
            };
        }

        public static SForceClient Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (new object())
                    {
                        if (instance == null)
                        {
                            instance = new SForceClient();
                        }
                    }
                }

                return instance;
            }
            private set { }
        }

        public bool Login(String userName, String password, String securityToken)
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
                // this.serverUrl = sfSvc.Url;

                return true;
            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
                throw ex;
            }

            //return false;
        }

        public void Logout()
        {
            sfSvc.logout();
        }

        public List<SObjectEntryBase> GetSObjects(bool force = false)
        {
            // cache objects
            List<sforce.SObjectEntryBase> sobjects = this.sfSession.SObjects == null ? new List<SObjectEntryBase>() : this.sfSession.SObjects;

            if (!force && sobjects.Any())
            {
                return sobjects;
            }

            sobjects.Clear();

            // get SObjects
            // Make the describeGlobal() call 

            DescribeGlobalResult globalDesc = null;
            try
            {
                globalDesc = sfSvc.describeGlobal();
            }
            catch ( System.Web.Services.Protocols.SoapException ex)
            {
                if (string.Equals(ex.Code.Name, "INVALID_SESSION_ID", StringComparison.InvariantCultureIgnoreCase)) {
                    Auth.AuthServer.RefreshAccessToken(this.sfSession);
                    this.SetSession(this.sfSession);

                    globalDesc = sfSvc.describeGlobal();
                }
                else
                {
                    throw;
                }
            }

            // Get the sObjects from the describe global result
            DescribeGlobalSObjectResult[] sObjResults = globalDesc.sobjects;

            foreach (var sobj in globalDesc.sobjects)
            {
                // if (sobj.queryable && sobj.createable && sobj.updateable && sobj.deletable)
                // if (sobj.queryable && sobj.createable && sobj.deletable)
                if (sobj.queryable && sobj.updateable)
                {
                    sobjects.Add(new SObjectEntry(sobj.name, sobj.label, sobj.keyPrefix, sobj.custom, sobj.customSetting, this, sobj.labelPlural));
                }
            }

            sobjects.Sort((a, b) => { return string.Compare(a.Label, b.Label); });
            this.sfSession.SObjects = sobjects;

            return sobjects;
        }

        public List<sforce.SObjectEntryBase> DescribeSObject(SObjectEntryBase sobj)
        {
            List<sforce.SObjectEntryBase> fields = new List<SObjectEntryBase>();
            DescribeSObjectResult result = null;
            try
            {
                result = this.sfSvc.describeSObject(sobj.Name);
            }
            catch (System.Web.Services.Protocols.SoapException ex)
            {
                if (string.Equals(ex.Code.Name, "INVALID_SESSION_ID", StringComparison.InvariantCultureIgnoreCase))
                {
                    Auth.AuthServer.RefreshAccessToken(this.sfSession);
                    this.SetSession(this.sfSession);

                    result = this.sfSvc.describeSObject(sobj.Name);
                }
                else
                {
                    throw;
                }
            }

            if (result == null)
            {
                return fields;
            }

            // field types: https://developer.salesforce.com/docs/atlas.en-us.api.meta/api/field_types.htm
            // field.type
            foreach (var field in result.fields)
            {
                FieldEntry entry = new FieldEntry(field.name, field.label, field.custom, this, sobj);

                entry.IsRequired = !field.nillable;
                entry.IsReadonly = field.autoNumber // auto-number name field
                    || field.calculated // formula field
                    || field.type == fieldType.id // Id field
                    || (!field.updateable && field.defaultedOnCreate); // created date, created by Id, etc

                fields.Add(entry);

                if (field.referenceTo != null) // lookup/master-detail field
                {
                    string fieldName = string.Format("{0}.Name", field.relationshipName); // assume Name is on the related object
                    FieldEntry nameField = new FieldEntry(fieldName, field.label, field.custom, this, sobj);
                    nameField.IsRequired = false;
                    nameField.IsReadonly = true;

                    fields.Add(nameField);
                }
            }

            // var relation = result.childRelationships;
            // var cr = result.childRelationships;

            return fields;
        }

        public System.Data.DataTable ExecQuery(string query, string tableName, System.Data.DataTable dt)
        {
            QueryResult ret = null;

            try
            {
                ret = this.sfSvc.query(query);
            }
            catch (System.Web.Services.Protocols.SoapException ex)
            {
                if (string.Equals(ex.Code.Name, "INVALID_SESSION_ID", StringComparison.InvariantCultureIgnoreCase))
                {
                    Auth.AuthServer.RefreshAccessToken(this.sfSession);
                    this.SetSession(this.sfSession);

                    ret = this.sfSvc.query(query);
                }
                else
                {
                    throw;
                }
            }

            if (ret == null || ret.records == null)
            {
                //throw new Exception("No data loaded!");
                return null;
            }

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
                        string value = col.InnerText;

                        // if (col.HasAttributes && !string.IsNullOrEmpty(col.GetAttribute("xsi:type"))) // relationship field
                        if (col.HasAttributes) // relationship field
                        {
                            if (!string.IsNullOrEmpty(col.GetAttribute("xsi:type")))
                            {
                                fieldName = string.Format("{0}.{1}", col.LocalName, col.LastChild.LocalName);
                                value = col.LastChild.InnerText;
                            }
                            else
                            {
                                continue; // relationship field but no value returned
                            }
                        }

                        dr[fieldName] = value;
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

        public List<object> DoUpdate(System.Data.DataTable table)
        {
            List<object> resultList = new List<object>();

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

            if (upsertList.Count > 0)
            {
                UpsertResult[] results =  sfSvc.upsert("Id", upsertList.ToArray());
                resultList.AddRange(results);

                foreach (UpsertResult ret in results)
                {
                    if (!ret.success)
                    {
                        Error[] errors = ret.errors;
                    }

                }
            }

            if (idsToDelete.Count > 0)
            {
                DeleteResult[] results = sfSvc.delete(idsToDelete.ToArray());
                resultList.AddRange(results);

                foreach (DeleteResult result in results)
                {
                    if (!result.success)
                    {
                        Error[] errors = result.errors;
                    }
                }
            }

            return resultList;
        }

        public List<object> DoUpdate2(System.Data.DataTable table)
        {
            List<object> resultList = new List<object>();

            // Dictionary<int, int> updateIndexMap = new Dictionary<int, int>();
            // Dictionary<int, int> insertIndexMap = new Dictionary<int, int>();
            Dictionary<int, int> indexMap = new Dictionary<int, int>();
            int updateIdx = 0, insertIdx = 0;
            System.Data.DataTable updatedTable = table.GetChanges(System.Data.DataRowState.Modified);
            System.Data.DataTable deletedTable = table.GetChanges(System.Data.DataRowState.Deleted);
            System.Data.DataTable addedTable = table.GetChanges(System.Data.DataRowState.Added);

            for (int idx = 0; idx < table.Rows.Count; idx++)
            {
                if (table.Rows[idx].RowState == DataRowState.Modified)
                {
                    indexMap.Add(updateIdx++, idx);
                }
                else if (table.Rows[idx].RowState == DataRowState.Added)
                {
                    indexMap.Add((updatedTable == null ? 0 : updatedTable.Rows.Count) + insertIdx++, idx);
                }
            }

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
                        if (column.ReadOnly)
                        {
                            continue;
                        }

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
                        if (column.ReadOnly)
                        {
                            continue;
                        }

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

            if (upsertList.Count > 0)
            {
                UpsertResult[] results = sfSvc.upsert("Id", upsertList.ToArray());
                resultList.AddRange(results);

                for (int idx = 0; idx < results.Length; idx++)
                {
                    if (results[idx].success)
                    {
                        // if (results[idx].created) // insert
                        {
                            table.Rows[indexMap[idx]].AcceptChanges();
                        }
                    }
                }
            }

            if (idsToDelete.Count > 0)
            {
                DeleteResult[] results = sfSvc.delete(idsToDelete.ToArray());
                resultList.AddRange(results);

                foreach (DeleteResult result in results)
                {
                    if (!result.success)
                    {
                        Error[] errors = result.errors;
                    }
                }
            }

            return resultList;
        }

        private static bool HasCellChanged(DataRow row, DataColumn col)
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
