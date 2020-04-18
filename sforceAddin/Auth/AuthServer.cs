using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace sforceAddin.Auth
{
    class AuthServer
    {
        private TcpListener server;
        private Func<sforce.SFSession, bool> callback;
        private bool isStarted = false;

        private static AuthServer _instance;

        private AuthServer(Func<sforce.SFSession, bool> callback)
        {
            this.callback = callback;
        }

        public static AuthServer GetAuthServer(Func<sforce.SFSession, bool> callback)
        {
            if (_instance == null)
            {
                lock (new object())
                {
                    if (_instance == null)
                    {
                        _instance = new AuthServer(callback);
                    }
                }
            }

            return _instance;
        }

        public void startServer(int port)
        {
            if (this.server == null)
            {
                IPEndPoint endPoint = new IPEndPoint(IPAddress.Loopback, port);
                this.server = new TcpListener(endPoint);
            }
            if (!this.isStarted)
            {
                this.server.Start();

                this.isStarted = true;
            }
            // Console.WriteLine("Listening on port " + port);
        }

        public void handleRequest()
        {
            while (this.isStarted)
            {
                TcpClient client = server.AcceptTcpClient();
                client.ReceiveTimeout = 1000 * 3;
                
                NetworkStream netStream = client.GetStream();
                byte[] buffer = new byte[1024 * 4];
                int length = netStream.Read(buffer, 0, 4096);

                // Nothing in this request, ignore it
                if (length == 0)
                {
                    continue;
                }

                string reqString = Encoding.UTF8.GetString(buffer);
                // get code and state
                // string[] headerLines = reqString.Split(new string[] {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);
                // var newLinesRegex = new Regex(@"\r\n|\n|\r", RegexOptions.Singleline);
                // var lines = newLinesRegex.Split(reqString);

                // string pattern = @"code=(?<code>\S*?)( |&state=(?<state>\S*)).+?\sReferer: (?<referer>\S*)"; // On IE, no referer???
                string pattern = @"code=(?<code>\S*?)( |&state=(?<state>\S*)).+?";
                Regex rx = new Regex(pattern, RegexOptions.Singleline);
                Match m = rx.Match(reqString);
                string code = string.Empty, state = string.Empty, referer = string.Empty;
                if (m.Success)
                {
                    code = m.Groups["code"].Success ? m.Groups["code"].Value : "";
                    state = m.Groups["state"].Success ? m.Groups["state"].Value : "";
                    // referer = m.Groups["referer"].Success ? m.Groups["referer"].Value : "";
                }
                else
                {
                    continue;
                }

                // if (string.Compare(state, "one", true) == 0)

                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls
                       | System.Net.SecurityProtocolType.Tls11
                       | System.Net.SecurityProtocolType.Tls12;

                //string postData = string.Format("grant_type=authorization_code&code={0}&client_id={1}&client_secret={2}&redirect_uri={3}"
                //                    , System.Net.WebUtility.UrlEncode(code), client_id, client_secret, System.Net.WebUtility.UrlEncode(redirect_url));
                string postData = string.Format("grant_type=authorization_code&code={0}&&redirect_uri={1}&state=two"
                                   , code, System.Net.WebUtility.UrlEncode(AuthUtil.redirect_url));
                var request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(AuthUtil.baseUrl + AuthUtil.token_url);
                // var request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(domainInstanceUri + token_url);

                var data = Encoding.UTF8.GetBytes(postData);

                //foreach (var key in context.Request.Headers.AllKeys)
                //{
                //    request.Headers[key] = context.Request.Headers[key];
                //}
                 
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";// grant type not supported
                request.ContentLength = data.Length;
                request.Accept = "application/xml;charset=UTF-8";
                // request.Accept = "application/x-www-form-urlencoded";

                // request.Headers.Add("Authorization", string.Format("Basic client_id={0}&client_secret={1}", client_id, client_secret));
                string basicCredential = string.Format("Basic {0}", Convert.ToBase64String(Encoding.UTF8.GetBytes(string.Format("{0}:{1}", AuthUtil.client_id, AuthUtil.client_secret))));
                request.Headers.Add("Authorization", basicCredential);

                using (Stream stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }

                try
                {
                    using (var response = (System.Net.HttpWebResponse)request.GetResponse())
                    {
                        var responseString = new System.IO.StreamReader(response.GetResponseStream()).ReadToEnd();

                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml(responseString);

                        XmlSerializer xs = new XmlSerializer(typeof(OAuth));
                        OAuth oAuthObj = (OAuth)xs.Deserialize(new StringReader(responseString));

                        // sforce.SFSession sfSession = sforce.SFSession.GetSession();
                        //sforce.SFSession sfSession = new sforce.SFSession();
                        //sfSession.SessionId = oAuthObj.access_token;
                        //sfSession.Signature = oAuthObj.signature;
                        //sfSession.Id = oAuthObj.id;
                        //// sfSession.IdToken = oAuthObj.id_token;
                        //sfSession.refreshToken = oAuthObj.refresh_token;
                        //sfSession.InstanceUrl = oAuthObj.instance_url;
                        //sfSession.IssuedAt = oAuthObj.issued_at;
                        //sfSession.Scope = oAuthObj.scope;
                        //sfSession.TokenType = oAuthObj.token_type;
                        //sfSession.ApiVersion = AuthUtil.apiVersion;
                        //sfSession.IsValid = true;

                        sforce.SFSession sfSession = new sforce.SFSession(oAuthObj);
                        //sforce.Connection connection = new sforce.Connection(sfSession);
                        //sforce.ConnectionManager.Instance.AddConnection(connection);
                        sforce.SFSessionManager.Instance.AddSession(sfSession);

                        // response
                        string statusLine = "HTTP/1.1 200 OK\r\n";
                        string resBody = string.Format(@"<html>
                                            <head>
                                                <title>Authenticiation</title>
                                            </head>
                                            <body>Login successfuly, you may close browser now.</body>
                                        </html>", DateTime.Now);
                        string resHeader = string.Format("Content-type:text/html;charset=utf-8\r\nContent-Lenght:{0}\r\n", resBody.Length);
                        byte[] resBodyBytes = Encoding.UTF8.GetBytes(resBody);
                        byte[] statusLineBytes = Encoding.UTF8.GetBytes(statusLine);
                        byte[] resHeaderBytes = Encoding.UTF8.GetBytes(resHeader);

                        netStream.Write(statusLineBytes, 0, statusLineBytes.Length);
                        netStream.Write(resHeaderBytes, 0, resHeaderBytes.Length);
                        netStream.Write(new byte[] { 13, 10 }, 0, 2); // ??
                        netStream.Write(resBodyBytes, 0, resBodyBytes.Length);

                        if (this.callback != null)
                        {
                            this.callback(sfSession);
                        }

                        netStream.Close();
                        netStream.Dispose();
                        client.Close();

                        this.server.Stop();
                        this.isStarted = false;
                        break;
                    }
                }
                catch (System.Net.WebException ex)
                {
                    if (ex.Response != null)
                    {
                        string content = new System.IO.StreamReader(ex.Response.GetResponseStream()).ReadToEnd();
                        Console.WriteLine(content);
                    }
                    //throw ex;
                }
            }
        }

        public static void RefreshAccessToken(sforce.SFSession expiredSession)
        {
            //string postData = string.Format("grant_type=refresh_token&refresh_token={0}&&client_id={1}&client_secret={2}"
            //                      , expiredSession.refreshToken, Auth.AuthUtil.client_id, Auth.AuthUtil.client_secret);
            string postData = string.Format("grant_type=refresh_token&refresh_token={0}", expiredSession.RefreshToken);
            var request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(AuthUtil.baseUrl + AuthUtil.token_url);

            string basicCredential = string.Format("Basic {0}", Convert.ToBase64String(Encoding.UTF8.GetBytes(string.Format("{0}:{1}", AuthUtil.client_id, AuthUtil.client_secret))));
            request.Headers.Add("Authorization", basicCredential);

            var data = Encoding.UTF8.GetBytes(postData);

            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";// grant type not supported
            request.ContentLength = data.Length;
            request.Accept = "application/xml;charset=UTF-8";
            // request.Accept = "application/x-www-form-urlencoded";

            // request.Headers.Add("Authorization", string.Format("Basic client_id={0}&client_secret={1}", client_id, client_secret));
            //string basicCredential = string.Format("Basic {0}", Convert.ToBase64String(Encoding.UTF8.GetBytes(string.Format("{0}:{1}", AuthUtil.client_id, AuthUtil.client_secret))));
            //request.Headers.Add("Authorization", basicCredential);

            using (Stream stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            try
            {
                using (var response = (System.Net.HttpWebResponse)request.GetResponse())
                {
                    var responseString = new System.IO.StreamReader(response.GetResponseStream()).ReadToEnd();

                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(responseString);

                    XmlSerializer xs = new XmlSerializer(typeof(OAuth));
                    OAuth oAuthObj = (OAuth)xs.Deserialize(new StringReader(responseString));

                    // sforce.SFSession sfSession = sforce.SFSession.GetSession();
                    // sforce.SFSession sfSession = new sforce.SFSession();
                    //expiredSession.SessionId = oAuthObj.access_token;
                    //expiredSession.Signature = oAuthObj.signature;
                    //expiredSession.Id = oAuthObj.id;
                    //// expiredSession.IdToken = oAuthObj.id_token;
                    //if (!string.IsNullOrEmpty(oAuthObj.refresh_token))
                    //{
                    //    expiredSession.refreshToken = oAuthObj.refresh_token;
                    //}
                    //expiredSession.InstanceUrl = oAuthObj.instance_url;
                    //expiredSession.IssuedAt = oAuthObj.issued_at;
                    //expiredSession.Scope = oAuthObj.scope;
                    //expiredSession.TokenType = oAuthObj.token_type;
                    //expiredSession.ApiVersion = AuthUtil.apiVersion;
                    //expiredSession.IsValid = true;
                    expiredSession.RefreshSession(oAuthObj);
                    //sforce.Connection connection = new sforce.Connection(sfSession);
                    //sforce.ConnectionManager.Instance.AddConnection(connection);
                }
            }
            catch (System.Net.WebException ex)
            {
                if (ex.Response != null)
                {
                    string content = new System.IO.StreamReader(ex.Response.GetResponseStream()).ReadToEnd();
                    Console.WriteLine(content);
                }

                throw new Exception("Failed to refresh access token. Please login manually. Error: " + ex.Message);
            }
        }

        public void stopServer()
        {
            server.Stop();
        }
    }
}
