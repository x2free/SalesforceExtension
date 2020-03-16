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

        public SFSession() { }

        //public static SFSession GetSession()
        //{
        //    if (_session == null)
        //    {
        //        lock (new object())
        //        {
        //            if (_session == null)
        //            {
        //                _session = new SFSession();
        //            }
        //        }
        //    }

        //    return _session;
        //}

        public bool IsValid;
        public string SessionId;
        public string Scope;
        public string Signature;
        public string IdToken;
        public string InstanceUrl;
        public string Id
        {
            get
            {
                return this.id;
            }
            set
            {
                if (string.Compare(this.Id, value, true) != 0)
                {
                    this.id = value;
                    string pattern = @"id/(?<oId>\w+)*?/(?<uId>\S+)";
                    Regex regex = new Regex(pattern);
                    Match match = regex.Match(id);

                    if (match.Success)
                    {
                        this.OrgId = match.Groups["oId"].Success ? match.Groups["oId"].Value : string.Empty;
                        this.UserId = match.Groups["uId"].Success ? match.Groups["uId"].Value : string.Empty;
                    }
                }
            }
        }
        public string TokenType;
        public string IssuedAt;
        public string UserId;
        public string OrgId;
        public int ApiVersion;

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

        private string id; // eg, https://login.salesforce.com/id/00D6F000002WOdlUAG/0056F00000ACTvNQAX
        private string soapPartnerUrl;
        private string soapMetadataUrl;
    }
}
