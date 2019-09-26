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

        private SForceClient sfClient;

        public SObjectEntryBase(String name, String label, bool isCustom, SForceClient sfClient)
        {
            this.Name = name;
            this.Label = label;
            // this.LabelPlural = pluralLabel;
            // this.KeyPrefix = keyPrefix;
            this.IsCustom = isCustom;
            // this.isCustomSetting = isCustomSetting;
            this.sfClient = sfClient;
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

        public List<SObjectEntryBase> getChildren()
        {
            return this.sfClient.describeSObject(this);
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

        public SObjectEntry(String name, String label, String keyPrefix, bool isCustom, bool isCustomSetting, SForceClient sfClient, String pluralLabel = null)
            : base(name, label, isCustom, sfClient)
        {
            this.IsCustomSetting = isCustomSetting;
            this.KeyPrefix = keyPrefix;
            this.LabelPlural = pluralLabel;

            // SubSObjects = new List<SObjectEntryBase>();
        }
    }

    class FieldEntry : SObjectEntryBase
    {
        /// <summary>
        /// To indicate which table/custom setting this field belongs to
        /// </summary>
        private SObjectEntryBase parent;

        public FieldEntry(String name, String label, bool isCustom, SForceClient sfClient, SObjectEntryBase parent)
            : base(name, label, isCustom, sfClient)
        {
            this.parent = parent;
        }
    }
}
