using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sforceAddin.sforce
{
    /// <summary>
    /// Indicates an SObject/Custom Setting
    /// </summary>
    class SObjectEntry
    {
        /// <summary>
        /// Label
        /// </summary>
        public String Label { get; private set; }
        /// <summary>
        /// Plural label
        /// </summary>
        public String LabelPlural { get; private set; }
        /// <summary>
        /// API name
        /// </summary>
        public String Name { get; private set; }
        /// <summary>
        /// Id prefix for SObject
        /// </summary>
        public String KeyPrefix { get; private set; }
        /// <summary>
        /// Is custom SObject or not
        /// </summary>
        public bool IsCustom { get; private set; }
        /// <summary>
        /// Is custom setting or SObject
        /// </summary>
        public bool isCustomSetting { get; private set; }


        public SObjectEntry(String name, String label, String pluralLabel, String keyPrefix, bool isCustom, bool isCustomSetting) {
            this.Name = name;
            this.Label = label;
            this.LabelPlural = pluralLabel;
            this.KeyPrefix = keyPrefix;
            this.IsCustom = isCustom;
            this.isCustomSetting = isCustomSetting;
        }
    }
}
