using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sforceAddin.Auth
{
    class AuthUtil
    {
        public static string authorization_url = "/services/oauth2/authorize";
        public static string token_url = "/services/oauth2/token";
        public static string revoke_url = "/services/oauth2/revoke";
        public static string baseUrl = "https://test.salesforce.com";
        // public static string client_id = "3MVG9YDQS5WtC11o5Mbm9Am1IBP7MyithezCXauojL8lCuh42psSRB4CRxCxQ8BcWpzZMOvvnPi6oQioIO8Ot";
        public static string client_id = "3MVG9YDQS5WtC11o5Mbm9Am1IBOEpngXdHJ1srU2WtFQHzjNx3PE4Ray.jp8CqHweBVEyscilAeWnoI3dG1VD";
        // public static int port = 9286;
        public static int port = 15506;
        public static string redirect_url = string.Format("http://localhost:{0}", port);
        // public static string client_secret = "158FF1F4FBE35220BB658C5BFF30771CE2D9FF7F6CAF11925984956C184C20F8";
        public static string client_secret = "C23E3781FFF6AF8419362F0FA819C0F8358B57598A8065DB83DC2F216EC1E3BC";
        public static int apiVersion = 48;

        public static void doAuth(Func<sforce.SFSession, bool> callback)
        {
            /*
                https://login.salesforce.com/services/oauth2/authorize?response_type=code
                &client_id=3MVG9lKcPoNINVBIPJjdw1J9LLM82HnFVVX19KY1uA5mu0QqEWhqKpoW3svG3X
                HrXDiCQjK1mdgAvhCscA9GE&redirect_uri=https%3A%2F%2Fwww.mysite.com%2F
                code_callback.jsp&state=mystate 
             */

            //theServer server = new theServer();
            //server.startServer();
            //System.Threading.Thread thread = new System.Threading.Thread(server.StartListen);
            //thread.Start();

            // start callback server
            AuthServer authSvr = AuthServer.GetAuthServer(callback);
            authSvr.startServer(AuthUtil.port);
            System.Threading.Thread thread = new System.Threading.Thread(authSvr.handleRequest);
            thread.Start();

            string getTookenReqUrl = string.Format("{0}{1}?response_type=code&client_id={2}&redirect_uri={3}&state={4}&prompt=select_account" // &scope=refresh_token // If add this parameter, cannot use the returned session Id
                                                                                                                                              // to retrieve sobjects with error: INVALID_SESSION_ID: This session is not valid for use with the API
                                                                                                                                              // but it does return refresh_token which can be used to get access token (but cannot use)
                        , baseUrl, authorization_url
                        , client_id, redirect_url
                        , "one");

            System.Diagnostics.Process.Start(getTookenReqUrl);
        }
    }
}
