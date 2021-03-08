using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sforceAddin.sforce
{
    public class DataColumnX : System.Data.DataColumn
    {
        public DataColumnX() : base() { }
        public DataColumnX(string columnName) : base(columnName) { }
        public DataColumnX(string columnName, Type dataType) : base(columnName, dataType) { }
        public DataColumnX(string columnName, Type dataType, string expr) : base(columnName, dataType, expr) { }
        public DataColumnX(string columnName, Type dataType, string expr, MappingType type) : base(columnName, dataType, expr, type) { }

        public bool IsReadonly { get; set; }
    }
}
