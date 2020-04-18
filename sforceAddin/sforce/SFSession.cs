using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace sforceAddin.sforce
{
    class SFSession
    {
        // private static SFSession _session;
        private Auth.OAuth oAuth2Obj;
        private string instanceName;
        public string RefreshToken { get; private set; } // Generally, this is set when do auth for 1st time only. When use it to refresh access token, it won't be changed
        // public SFSession() { }

        public SFSession(Auth.OAuth oAuth2Obj)
        {
            RefreshSession(oAuth2Obj);
            this.IsActive = false;
        }

        public void RefreshSession(Auth.OAuth oAuth2Obj)
        {
            if (oAuth2Obj == null)
            {
                throw new ArgumentNullException("oAuth2Obj");
            }

            this.oAuth2Obj = oAuth2Obj;
            this.instanceName = null;

            if (!string.IsNullOrEmpty(this.oAuth2Obj.refresh_token))
            {
                this.RefreshToken = this.oAuth2Obj.refresh_token;
            }

            string idUrl = this.oAuth2Obj.id; // eg, https://login.salesforce.com/id/00D6F000002WOdlUAG/0056F00000ACTvNQAX
            string pattern = @"id/(?<oId>\w+)*?/(?<uId>\S+)";
            Regex regex = new Regex(pattern);
            Match match = regex.Match(idUrl);

            if (match.Success)
            {
                this.OrgId = match.Groups["oId"].Success ? match.Groups["oId"].Value : string.Empty;
                this.UserId = match.Groups["uId"].Success ? match.Groups["uId"].Value : string.Empty;
            }
        }

        // public bool IsValid;
        public string SessionId
        {
            get { return this.oAuth2Obj.access_token; }
            private set { }
        }
        // public string Scope;
        // public string Signature;
        // public string IdToken;
        // public string refreshToken;
        public string InstanceUrl
        {
            get { return this.oAuth2Obj.instance_url; }
            private set { }
        }

        //public string TokenType;
        //public string IssuedAt;
        public string UserId;
        public string OrgId;
        // public int ApiVersion;

        public string SoapPartnerUrl
        {
            get
            {
                //if (string.IsNullOrEmpty(this.soapPartnerUrl))
                //{
                //    this.soapPartnerUrl = string.Format("{0}/services/Soap/u/{1}.0/{2}", this.InstanceUrl, this.ApiVersion, this.OrgId);
                //}

                //return this.soapPartnerUrl;

                return string.Format("{0}/services/Soap/u/{1}.0/{2}"
                        , this.InstanceUrl
                        , Auth.AuthUtil.apiVersion
                        , this.OrgId); ;
            }
        }
        public string SoapMetadataUrl
        {
            get
            {
                //if (string.IsNullOrEmpty(this.soapMetadataUrl))
                //{
                //    this.soapMetadataUrl = string.Format("{0}/services/Soap/m/{1}.0", this.InstanceUrl, this.ApiVersion);
                //}

                //return this.soapMetadataUrl;

                return string.Format("{0}/services/Soap/m/{1}.0"
                        , this.InstanceUrl
                        , Auth.AuthUtil.apiVersion);
            }
        }

        public string InstanceName
        {
            get
            {
                if (string.IsNullOrEmpty(this.instanceName))
                {
                    // get instance name from instance url
                    // Regex reg = new Regex(@"://(?<ins>.*).cs"); // sandbox only
                    Regex reg = new Regex(@"://(?<ins>\S*?)\..*");
                    Match match = reg.Match(this.InstanceUrl);
                    if (match.Success)
                    {
                        this.instanceName = match.Groups["ins"].Success ? match.Groups["ins"].Value : this.InstanceUrl;
                    }
                }

                return this.instanceName;
            }
            private set { }
        }

        /// <summary>
        /// To indicate which org we are working against, not indicate if is validate(expired, logged out) or not
        /// </summary>
        public bool IsActive;
        public List<SObjectEntryBase> SObjects { get; set; }

        //private string soapPartnerUrl;
        //private string soapMetadataUrl;
    }
}
