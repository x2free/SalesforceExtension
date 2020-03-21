using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace sforceAddin.sforce
{
    /*
     * Connection to an org, holds an Session
     */
    class Connection
    {
        private SFSession session;
        private string instanceName;
        private bool isActive;
        public List<sforce.SObjectEntryBase> SObjects { get; set; }

        public Connection(SFSession session)
        {
            if (session == null)
            {
                throw new ArgumentNullException("session", "Connection: session cannot be null");
            }

            this.session = session;
            // this.InstanceName = this.session.InstanceUrl;
            // this.instanceName = this.session.InstanceUrl;

            if (!string.IsNullOrEmpty(this.Session.InstanceUrl))
            {
                // get instance name from instance url
                // Regex reg = new Regex(@"://(?<ins>.*).cs"); // sandbox only
                Regex reg = new Regex(@"://(?<ins>\S*?)\..*");
                Match match = reg.Match(this.Session.InstanceUrl);
                if (match.Success)
                {
                    this.instanceName = match.Groups["ins"].Success ? match.Groups["ins"].Value : this.Session.InstanceUrl;
                }
            }
        }

        public void Active(Action<Connection> callback = null)
        {
            this.isActive = true;
            if (callback != null)
            {
                callback(this);
            }
        }

        public void Deactive(Action<Connection> callback = null)
        {
            this.IsActive = false;
            if (callback != null)
            {
                callback(this);
            }
        }

        public SFSession Session
        {
            get { return this.session; }
            private set { }
        }

        public string InstanceName
        {
            get { return this.instanceName; }
            private set { }
        }

        public bool IsActive
        {
            get { return this.isActive; }
            private set { }
        }
    }
}
