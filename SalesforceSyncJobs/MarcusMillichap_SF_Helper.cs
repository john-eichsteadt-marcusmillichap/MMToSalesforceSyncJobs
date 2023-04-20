using SalesforceSyncJobs.wsSalesforce;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Web.Services.Protocols;

namespace SalesforceSyncJobs
{
    class MarcusMillichap_SF_Helper
    {
        /// <summary>
        /// Log into salesforce
        /// </summary>
        /// <param name="service"></param>
        /// <returns></returns>
        public SforceService LogIn()
        {
            //create service object
            SforceService service = new SforceService();
            //set timeout to 1 min
            //service.Timeout = 60000;

            //logging in
            LoginResult lr;
            try
            {
                lr = service.login(ConfigurationManager.AppSettings["username"].ToString(), ConfigurationManager.AppSettings["password"].ToString() + ConfigurationManager.AppSettings["token"].ToString());
            }
            catch (SoapException ex)
            {
                //Email e = new Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                //                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                //                    "SalesForce API Error",
                //                    "Error Message: <br /><br />" + ex.Message +
                //                    "Inner Exception: <br /><br />" + ex.InnerException +
                //                    "Stack Trace: <br /><br />" + ex.StackTrace,
                //                    true);
                //e.SendEmail();
                return null;
            }

            //check if password expired
            if (lr.passwordExpired)
            {
                Console.WriteLine("An error has occurred. Your password has expired.");
                return null;
            }
            /** Once the client application has logged in successfully, it will use
            * the results of the login call to reset the endpoint of the service
            * to the virtual server instance that is servicing your organization
            */
            // Save old authentication end point URL
            string authEndPoint = service.Url;
            // Set returned service endpoint URL
            service.Url = lr.serverUrl;
            /** The sample client application now has an instance of the SforceService
              * that is pointing to the correct endpoint. Next, the sample client
              * application sets a persistent SOAP header (to be included on all
              * subsequent calls that are made with SforceService) that contains the
              * valid sessionId for our login credentials. To do this, the sample
              * client application creates a new SessionHeader object and persist it to
              * the SforceService. Add the session ID returned from the login to the
              * session header
              */
            service.SessionHeaderValue = new SessionHeader();
            service.SessionHeaderValue.sessionId = lr.sessionId;

            //for debugging purpose,  writes out user info to make sure it's working correctly
            //printUserInfo(lr, authEndPoint);

            // Return true to indicate that we are logged in, pointed  
            // at the right URL and have our security token in place.     
            return service;
        }
    }
}
