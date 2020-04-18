using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sforceAddin.sforce
{
    class SFSessionManager
    {
        private static SFSessionManager sessionManger;
        public List<SFSession> Sessions;

        private SFSessionManager()
        {
            this.Sessions = new List<SFSession>();
        }

        public static SFSessionManager Instance
        {
            get
            {
                if (sessionManger == null)
                {
                    lock (new object())
                    {
                        if (sessionManger == null)
                        {
                            sessionManger = new SFSessionManager();
                        }
                    }
                }

                return sessionManger;
            }

            private set { }
        }

        public SFSession FindSession(string instanceName)
        {
            return this.Sessions.Find(session => session.InstanceName.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase));
        }

        public void AddSession(SFSession session)
        {
            this.Sessions.RemoveAll(sess => sess.InstanceName.Equals(session.InstanceName, StringComparison.CurrentCultureIgnoreCase));
            this.Sessions.Add(session);
        }

        public SFSession ActiveSession
        {
            get { return this.Sessions.FirstOrDefault(s => s.IsActive); }
            private set { }
        }
    }
}
