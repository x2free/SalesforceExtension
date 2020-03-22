using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sforceAddin.sforce
{
    class ConnectionManager
    {
        private static ConnectionManager connectionManger;
        public List<Connection> Connections;

        //private ActiveConnectionChanged callback;
        //public delegate void ActiveConnectionChanged(Connection con);

        private ConnectionManager()
        {
            this.Connections = new List<Connection>();
        }

        // public bool AddConnection(Connection connection, ActiveConnectionChanged callback = null)
        public bool AddConnection(Connection connection)
        {
            //List<Connection> conns = this.Connections.FindAll(con => con.InstanceName.Equals(connection.InstanceName, StringComparison.CurrentCultureIgnoreCase));

            //if (conns != null && conns.Count > 0)
            //{
            this.Connections.RemoveAll(con => con.InstanceName.Equals(connection.InstanceName, StringComparison.CurrentCultureIgnoreCase));
            this.Connections.Add(connection);
            // }

            //if (connection.IsActive)
            //{
            //    foreach (var con in this.Connections)
            //    {
            //        con.Deactive();
            //    }
            //}

            //this.Connections.Add(connection);

            //if (callback != null)
            //{
            //    this.callback += callback;
            //}

            return true;
        }

        public Connection FindConnection(string instanceName)
        {
            return this.Connections.Find(con => con.InstanceName.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase));
        }

        public Connection ActiveConnection {
            get {
                return this.Connections.FirstOrDefault(c => c.IsActive);
            }

            set {
                foreach (Connection con in Connections.FindAll(c => c.IsActive))
                {
                    //con.IsActive = false;
                    con.Deactive();
                }

                Connection connection = Connections.Find(c => String.Equals(c.InstanceName, value.InstanceName, StringComparison.CurrentCultureIgnoreCase));
                //connection.IsActive = true;
                //callback(connection);
                connection.Active();
            }
        }

        public static ConnectionManager Instance {
            get {
                if (connectionManger == null)
                {
                    lock (new object())
                    {
                        if (connectionManger == null)
                        {
                            connectionManger = new ConnectionManager();
                        }
                    }
                }

                return connectionManger;
            }
            private set { }
        }
    }
}
