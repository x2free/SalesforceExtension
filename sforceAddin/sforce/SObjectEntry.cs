using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sforceAddin.sforce
{
    /*
     SObject -- SObject and Custom Setting
     Field
     Relation
    */

    /// <summary>
    /// Indicates an SObject/Custom Setting
    /// </summary>
    class SObjectEntryBase : IComparable
    {
        /// <summary>
        /// Label
        /// </summary>
        public String Label { get; private set; }
        /// <summary>
        /// API name
        /// </summary>
        public String Name { get; private set; }
        /// <summary>
        /// Is custom SObject or not
        /// </summary>
        public bool IsCustom { get; private set; }

        /// <summary>
        /// Related tables. Only used for SObject, not for Custom Setting
        /// </summary>
        public List<SObjectEntryBase> SubSObjects;

        public SObjectEntryBase(String name, String label, bool isCustom)
        {
            this.Name = name;
            this.Label = label;
            // this.LabelPlural = pluralLabel;
            // this.KeyPrefix = keyPrefix;
            this.IsCustom = isCustom;
            // this.isCustomSetting = isCustomSetting;
        }

        public int CompareTo(object obj)
        {
            SObjectEntryBase sobj = obj as SObjectEntryBase;

            if (sobj != null)
            {
                // Just used to sort our entries, so use label instead of name
                return string.Compare(this.Label, sobj.Label, StringComparison.CurrentCultureIgnoreCase);
            }

            return -1;
        }
    }

    class SObjectEntry : SObjectEntryBase
    {
        /// <summary>
        /// Is custom setting or SObject
        /// </summary>
        public bool IsCustomSetting { get; private set; }

        /// <summary>
        /// Id prefix for SObject
        /// </summary>
        public String KeyPrefix { get; private set; }

        /// <summary>
        /// Plural label
        /// </summary>
        public String LabelPlural { get; private set; }

        public SObjectEntry(String name, String label, String keyPrefix, bool isCustom, bool isCustomSetting, String pluralLabel = null)
            : base(name, label, isCustom)
        {
            this.IsCustomSetting = isCustomSetting;
            this.KeyPrefix = keyPrefix;
            this.LabelPlural = pluralLabel;

            // SubSObjects = new List<SObjectEntryBase>();
        }
    }

    class FieldEntry : SObjectEntryBase
    {
        public FieldEntry(String name, String label, bool isCustom)
            : base(name, label, isCustom)
        {

        }
    }
}
