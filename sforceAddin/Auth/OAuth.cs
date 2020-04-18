using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sforceAddin.Auth
{
    /// <summary>
    /// Entity for OAuth2 response
    /// </summary>
    public class OAuth
    {
        public string access_token;
        public string signature;
        public string scope;
        // public string id_token;
        public string instance_url;
        public string id; // eg, https://login.salesforce.com/id/00D6F000002WOdlUAG/0056F00000ACTvNQAX
        public string token_type;
        public string issued_at;
        public string refresh_token;
    }
}
