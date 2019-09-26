using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using sforceAddin.sforce;

namespace sforceAddin.UI
{
    abstract class SObjectNodeBase : TreeNode
    {
        /// <summary>
        /// SForceClient object, to integration with Salesforce
        /// </summary>
        // protected sforce.SForceClient sfClient;
        /// <summary>
        /// Sub nodes, eg, fields of an SOject
        /// </summary>
        /// private List<SObjectNodeBase> subNodes;

        /// <summary>
        /// SObject binding to this node
        /// </summary>
        protected sforce.SObjectEntryBase sobjEntry;

        //public SObjectNodeBase(String name, String label, sforce.SForceClient sfClient)
        //{
        //    this.Name = name;
        //    this.Text = label;

        //    this.sfClient = sfClient;
        //}

        public SObjectNodeBase(sforce.SObjectEntryBase sobj/*, sforce.SForceClient sfClient*/)
        {
            this.Name = sobj.Name;
            this.Text = sobj.Label;

            // this.sfClient = sfClient;

            //if (sobj is sforce.SObjectEntry) {

            //}

            this.sobjEntry = sobj;
        }

        /// <summary>
        /// Determine if <sobjeEntry> has sub entrys
        /// </summary>
        abstract public void expand();
        
    }

    class SObjectNode : SObjectNodeBase
    {
        public SObjectNode(sforce.SObjectEntryBase sobj/*, sforce.SForceClient sfClient*/)
            : base(sobj/*, sfClient*/) { }


        public override void expand()
        {
            if (this.Nodes != null && this.Nodes.Count > 0)
            {
                return;
            }

            var nodes = sobjEntry.getChildren();
            foreach (var item in nodes)
            {
                this.Nodes.Add(new FieldNode(item));
            }
        }
    }

    class FieldNode : SObjectNodeBase
    {
        public FieldNode(sforce.SObjectEntryBase sobj/*, sforce.SForceClient sfClient*/)
            : base(sobj/*, sfClient*/)
        {
        }

        public override void expand()
        {
            // Microsoft.Office.Interop.Excel.Worksheet sheet;

        }
    }
}
