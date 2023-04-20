using SalesforceSyncJobs.wsSalesforce;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SalesforceSyncJobs
{
    class MarcusMillichap_APEX_API
    {
        public wsSalesforceAPEX.SObjectServiceService ConnectToApexAPI(SforceService sf)
        {
            wsSalesforceAPEX.SObjectServiceService apex = null;
            string currentSessionId = string.Empty;
            if (sf != null)
            {
                currentSessionId = sf.SessionHeaderValue.sessionId;
                if (!string.IsNullOrEmpty(currentSessionId))
                {
                    apex = new wsSalesforceAPEX.SObjectServiceService();
                    apex.SessionHeaderValue = new wsSalesforceAPEX.SessionHeader();
                    apex.SessionHeaderValue.sessionId = currentSessionId;
                }
                else
                {
                    //no session id available login again
                    MarcusMillichap_SF_API api = new MarcusMillichap_SF_API();
                    sf = api.LogIn();
                    currentSessionId = sf.SessionHeaderValue.sessionId;
                    if (!string.IsNullOrEmpty(currentSessionId))
                    {
                        apex = new wsSalesforceAPEX.SObjectServiceService();
                        apex.SessionHeaderValue = new wsSalesforceAPEX.SessionHeader();
                        apex.SessionHeaderValue.sessionId = currentSessionId;
                    }
                }
            }
            else
            {
                MarcusMillichap_SF_API api = new MarcusMillichap_SF_API();
                sf = api.LogIn();
                currentSessionId = sf.SessionHeaderValue.sessionId;
                if (!string.IsNullOrEmpty(currentSessionId))
                {
                    apex = new wsSalesforceAPEX.SObjectServiceService();
                    apex.SessionHeaderValue = new wsSalesforceAPEX.SessionHeader();
                    apex.SessionHeaderValue.sessionId = currentSessionId;
                }
            }
            return apex;
        }
    }
}
