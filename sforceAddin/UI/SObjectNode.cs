using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace sforceAddin.UI
{
    class SObjectNode : TreeNode
    {
        /// <summary>
        /// SForceClient object, to integration with Salesforce
        /// </summary>
        private sforce.SForceClient sfClient;
        /// <summary>
        /// Sub nodes, eg, fields of an SOject
        /// </summary>
        /// private List<SObjectNode> subNodes;

        /// <summary>
        /// SObject binding to this node
        /// </summary>
        private sforce.SObjectEntryBase sobjEntry;

        public SObjectNode(String name, String label, sforce.SForceClient sfClient)
        {
            this.Name = name;
            this.Text = label;

            this.sfClient = sfClient;
        }

        public SObjectNode(sforce.SObjectEntryBase sobj, sforce.SForceClient sfClient)
        {
            this.Name = sobj.Name;
            this.Text = sobj.Label;

            this.sfClient = sfClient;

            if (sobj is sforce.SObjectEntry) {

            }
        }

        /// <summary>
        /// Determine if <sobjeEntry> has sub entrys
        /// </summary>
        public void expand()
        {
            var nodes = this.sfClient.describeSObject(this.Name);
            foreach (var item in nodes)
            {
                this.Nodes.Add(new SObjectNode(item, this.sfClient));
            }
        }
    }
}
