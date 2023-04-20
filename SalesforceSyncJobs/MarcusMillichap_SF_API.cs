using SalesforceSyncJobs.wsSalesforce;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Web.Services.Protocols;

namespace SalesforceSyncJobs
{
    class MarcusMillichap_SF_API
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
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error",
                                    "Error Message: <br /><br />" + ex.Message +
                                    "Inner Exception: <br /><br />" + ex.InnerException +
                                    "Stack Trace: <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
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
            printUserInfo(lr, authEndPoint);

            // Return true to indicate that we are logged in, pointed  
            // at the right URL and have our security token in place.     
            return service;
        }

        /// <summary>
        /// Log out of salesforce
        /// </summary>
        /// <param name="service"></param>
        public void LogOut(ref SforceService service)
        {
            try
            {
                service.logout();
            }
            catch { }
        }

        /// <summary>
        /// Upsert salesforce
        /// </summary>
        /// <param name="s"></param>
        /// <param name="service"></param>
        public bool Upsert(string externalId, sObject[] s, ref SforceService service, out List<MarcusMillichap_Errors> error)
        {
            bool results = true;
            error = new List<MarcusMillichap_Errors>();
            int index = 0;
            try
            {
                UpsertResult[] upsertResults = service.upsert(externalId, s);

                foreach (UpsertResult sr in upsertResults)
                {
                    if (!sr.success)
                    {
                        if (!sr.errors[0].message.Equals("invalid cross reference id", StringComparison.InvariantCultureIgnoreCase) &&
                            !sr.errors[0].message.Equals("entity is deleted", StringComparison.InvariantCultureIgnoreCase))
                        {
                            error.Add(new MarcusMillichap_Errors { SalesforceId = sr.id, ErrorMessage = sr.errors[0].message, ExternalId = s[index].Id });
                            results = false;
                        }
                    }
                    else
                    {
                        Console.WriteLine("\nItem Upserted ......" + sr.id + "\n");

                    }
                    index++;
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error - Upsert",
                                    "Error Message: <br /><br />" + ex.Message +
                                    "Inner Exception: <br /><br />" + ex.InnerException +
                                    "Stack Trace: <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                results = false;
            }
            if (error != null && error.Count > 0)
            {
                //send error emails
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                    "SalesForce API Error - Upsert",
                    MarcusMillichap_Email.GenerateErrorBulkMessage(error, externalId),
                    true);
                e.SendEmail();
            }
            return results;
        }

        /// <summary>
        /// Upsert salesforce
        /// </summary>
        /// <param name="s"></param>
        /// <param name="service"></param>
        public bool UpsertUser(string externalId, sObject[] s, ref SforceService service, out List<MarcusMillichap_Errors> error)
        {
            bool results = true;
            error = new List<MarcusMillichap_Errors>();
            int index = 0;
            try
            {
                UpsertResult[] upsertResults = service.upsert(externalId, s);

                foreach (UpsertResult sr in upsertResults)
                {
                    if (!sr.success)
                    {
                        if (!sr.errors[0].message.Equals("invalid cross reference id", StringComparison.InvariantCultureIgnoreCase))
                        {
                            error.Add(new MarcusMillichap_Errors { SalesforceId = sr.id, ErrorMessage = sr.errors[0].message, EmailAddress = ((User)s[index]).Email, FirstName = ((User)s[index]).FirstName, LastName = ((User)s[index]).LastName });
                            results = false;
                        }
                    }
                    else
                    {
                        Console.WriteLine("\nItem Upserted ......" + sr.id + "\n");

                    }
                    index++;
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error - Upsert",
                                    "Error Message: <br /><br />" + ex.Message +
                                    "Inner Exception: <br /><br />" + ex.InnerException +
                                    "Stack Trace: <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                results = false;
            }
            if (error != null && error.Count > 0)
            {
                //send error emails
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                    "SalesForce API Error - Upsert",
                    MarcusMillichap_Email.GenerateErrorBulkMessage(error, externalId),
                    true);
                e.SendEmail();
            }
            return results;
        }

        /// <summary>
        /// Update salesforce
        /// </summary>
        /// <param name="s"></param>
        /// <param name="service"></param>
        public bool Update(sObject[] s, ref SforceService service, out Dictionary<string, sObject> success, out List<MarcusMillichap_Errors> error)
        {
            error = new List<MarcusMillichap_Errors>();
            bool results = true;
            success = new Dictionary<string, sObject>();
            int index = 0;
            try
            {
                SaveResult[] saveResults = service.update(s);
                foreach (SaveResult sr in saveResults)
                {
                    if (!sr.success)
                    {
                        error.Add(new MarcusMillichap_Errors { SalesforceId = sr.id, ErrorMessage = sr.errors[0].message, ExternalId = s[index].Id });
                        results = false;
                    }
                    else
                    {
                        Console.WriteLine("\nItem Updates ......" + sr.id + "\n");
                        if (!success.ContainsKey(sr.id))
                        {
                            success.Add(sr.id, s[index]);
                        }
                    }
                    index++;
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error - Update",
                                    "Error Message: <br /><br />" + ex.Message +
                                    "Inner Exception: <br /><br />" + ex.InnerException +
                                    "Stack Trace: <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                results = false;
            }
            if (error != null && error.Count > 0)
            {
                //send error emails
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                    "SalesForce API Error - Update",
                    MarcusMillichap_Email.GenerateErrorBulkMessage(error),
                    true);
                e.SendEmail();
            }
            return results;
        }

        /// <summary>
        /// Create salesforce
        /// </summary>
        /// <param name="s"></param>
        /// <param name="service"></param>
        public bool Create(sObject[] s, ref SforceService service, out Dictionary<string, sObject> success, out List<MarcusMillichap_Errors> error)
        {
            error = new List<MarcusMillichap_Errors>();
            success = new Dictionary<string, sObject>();
            bool results = true;
            int index = 0;
            try
            {
                SaveResult[] saveResults = service.create(s);
                foreach (SaveResult sr in saveResults)
                {
                    if (!sr.success)
                    {
                        error.Add(new MarcusMillichap_Errors { SalesforceId = sr.id, ErrorMessage = sr.errors[0].message, ExternalId = s[index].Id });
                        results = false;
                    }
                    else
                    {
                        Console.WriteLine("\nItem Created ......" + sr.id + "\n");
                        if (!success.ContainsKey(sr.id))
                        {
                            success.Add(sr.id, s[index]);
                        }
                    }
                    index++;
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error - Create",
                                    "Error Message: <br /><br />" + ex.Message +
                                    "Inner Exception: <br /><br />" + ex.InnerException +
                                    "Stack Trace: <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                results = false;
            }
            if (error != null && error.Count > 0)
            {
                //send error emails
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                    "SalesForce API Error - Create",
                    MarcusMillichap_Email.GenerateErrorBulkMessage(error),
                    true);
                e.SendEmail();
            }
            return results;
        }

        /// <summary>
        /// Create salesforce
        /// </summary>
        /// <param name="s"></param>
        /// <param name="service"></param>
        public bool CreateGroups(sObject[] s, ref SforceService service, out Dictionary<string, sObject> success, out List<MarcusMillichap_Errors> error)
        {
            error = new List<MarcusMillichap_Errors>();
            success = new Dictionary<string, sObject>();
            bool results = true;
            int index = 0;
            try
            {
                SaveResult[] saveResults = service.create(s);
                foreach (SaveResult sr in saveResults)
                {
                    if (!sr.success)
                    {
                        error.Add(new MarcusMillichap_Errors { SalesforceId = sr.id, ErrorMessage = sr.errors[0].message, ExternalId = ((wsSalesforce.Group)s[index]).DeveloperName });
                        results = false;
                    }
                    else
                    {
                        Console.WriteLine("\nItem Created ......" + sr.id + "\n");
                        if (!success.ContainsKey(sr.id))
                        {
                            success.Add(sr.id, s[index]);
                        }
                    }
                    index++;
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error - Create Groups",
                                    "Error Message: <br /><br />" + ex.Message +
                                    "Inner Exception: <br /><br />" + ex.InnerException +
                                    "Stack Trace: <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                results = false;
            }
            if (error != null && error.Count > 0)
            {
                //send error emails
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                    "SalesForce API Error - Create",
                    MarcusMillichap_Email.GenerateErrorBulkMessage(error),
                    true);
                e.SendEmail();
            }
            return results;
        }

        /// <summary>
        /// Delete salesforce
        /// </summary>
        /// <param name="s"></param>
        /// <param name="service"></param>
        public bool Delete(string[] ids, ref SforceService service, out List<string> success, out List<MarcusMillichap_Errors> error)
        {
            error = new List<MarcusMillichap_Errors>();
            success = new List<string>();
            bool results = true;
            int index = 0;
            try
            {
                DeleteResult[] deleteResults = service.delete(ids);
                foreach (DeleteResult sr in deleteResults)
                {
                    if (!sr.success)
                    {
                        if (!sr.errors[0].message.Contains("insufficient access rights on object id"))
                        {
                            error.Add(new MarcusMillichap_Errors { SalesforceId = sr.id, ErrorMessage = sr.errors[0].message });
                            results = false;
                        }
                    }
                    else
                    {
                        Console.WriteLine("\nItem Deleted ......" + sr.id + "\n");
                        success.Add(sr.id);
                    }
                    index++;
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error - Delete",
                                    "Error Message: <br /><br />" + ex.Message +
                                    "Inner Exception: <br /><br />" + ex.InnerException +
                                    "Stack Trace: <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                results = false;
            }
            if (error != null && error.Count > 0)
            {
                //send error emails
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                    "SalesForce API Error - Delete",
                    MarcusMillichap_Email.GenerateErrorBulkMessage(error),
                    true);
                e.SendEmail();
            }
            return results;
        }


        private void printUserInfo(LoginResult lr, String authEP)
        {
            try
            {
                GetUserInfoResult userInfo = lr.userInfo;

                Console.WriteLine("\nLogging in ...\n");
                Console.WriteLine("UserID: " + userInfo.userId);
                Console.WriteLine("User Full Name: " +
                    userInfo.userFullName);
                Console.WriteLine("User Email: " +
                    userInfo.userEmail);
                Console.WriteLine();
                Console.WriteLine("SessionID: " +
                    lr.sessionId);
                Console.WriteLine("Auth End Point: " +
                    authEP);
                Console.WriteLine("Service End Point: " +
                    lr.serverUrl);
                Console.WriteLine();
            }
            catch (SoapException e)
            {
                Console.WriteLine("An unexpected error has occurred: " + e.Message +
                    " Stack trace: " + e.StackTrace);
            }
        }
    }
}
