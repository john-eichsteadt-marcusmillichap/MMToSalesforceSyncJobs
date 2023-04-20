using SalesforceSyncJobs.Mappers;
using SalesforceSyncJobs.Models;
using SalesforceSyncJobs.wsSalesforce;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;

namespace SalesforceSyncJobs
{
    class ActivityService
    {
        private bool Success = true;
        private SforceService SforceService;
        private readonly Dictionary<string, MarcusMillichap_User> UserList;

        public ActivityService(SforceService sforceService, Dictionary<string, MarcusMillichap_User> userList)
        {
            SforceService = sforceService;
            UserList = userList;
        }

        /// <summary>
        /// Public Function: Call to Process Activity Sync to SalesForce
        /// </summary>
        /// <param name="service"></param>
        public bool SyncToSalesforce(DateTime lastRunTime)
        {
            try
            {
                Dictionary<string, sObject> temp = new Dictionary<string, sObject>();
                List<MarcusMillichap_Errors> error = new List<MarcusMillichap_Errors>();

                var modifiedRecordsFromActivity = GetRecordsFromMNet(lastRunTime);
                var matchingModifiedSalesforceRecords = GetRecordsFromSalesForce(modifiedRecordsFromActivity);

                Update(matchingModifiedSalesforceRecords);

                //update activities in salesforce that was not modified by the sync, this captures the adds done manually by the agents
                SyncModifiedActivitiesInSalesforce();

                return Success;
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                   "SalesForce API Error - Activity.cs - SyncToSalesforce()",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                return false;
            }
        }


        /// <summary>
        /// Step 1: Get records that was last changed from MNet
        /// </summary>
        /// <param name="lastRunTime"></param>
        /// <returns></returns>
        private List<MMActivityModel> GetRecordsFromMNet(DateTime lastRunTime)
        {
            DataProvider dp = new DataProvider(ConfigurationManager.ConnectionStrings["MNet3"].ToString());
            NameValueCollection nv = new NameValueCollection();
            NameValueCollection nvOut = new NameValueCollection();
            NameValueCollection results = new NameValueCollection();
            DataSet ds = new DataSet();
            List<MMActivityModel> activities = new List<MMActivityModel>();
            try
            {
                nv.Add("@Time_PST", lastRunTime.ToString());
                nvOut.Add("@ErrCode", "4");
                nvOut.Add("@ErrMessage", "2048");
                ds = dp.GetDataSetMultipleOutputParam("[ActivityDetail].[spActivity_Search_ModifiedAfterTime]", CommandType.StoredProcedure, nv, nvOut, out results);

                // if there was an error, throw exception
                if (!string.IsNullOrEmpty(results["@ErrMessage"].ToString()))
                {
                    throw new Exception(results["@ErrMessage"].ToString());
                }

                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        activities.Add(new MMActivityModel
                        {
                            ActivityId = dr["ActivityID"].ToString()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                   "SalesForce API Error - Activity.cs - GetRecordsFromMNet()",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                Success = false;
            }
            finally {
                if (ds != null) { ds.Dispose(); }
                dp.Dispose();
                nv = null;
            }
            return activities;
        }

        /// <summary>
        /// Step 2: Based on records from MNet, query SalesForce to get list of matching records that was changed
        /// </summary>
        /// <param name="modifiedRecordsFromActivity"></param>
        /// <returns></returns>
        private List<SFActivityModel> GetRecordsFromSalesForce(List<MMActivityModel> modifiedRecordsFromActivity)
        {
            ArrayList queryList = new ArrayList();
            List<SFActivityModel> sfActivities = new List<SFActivityModel>();
            try
            {
                if (modifiedRecordsFromActivity.Count > 0)
                {
                    QueryResult qResult = null;
                    string queryParams = string.Empty;
                    int count = 0;
                    bool queryComplete = false;
                    //build sql query
                    string queryString = "SELECT id, MNet_Activity_ID__c, Owner.Id FROM Deal__c ";
                    queryString += "WHERE MNet_Activity_ID__c IN (";
                    foreach (var item in modifiedRecordsFromActivity)
                    {
                        if (queryParams.Equals(string.Empty)) queryParams += "'" + item.ActivityId + "'";
                        else queryParams += ",'" + item.ActivityId + "'";
                        count++;

                        if (queryParams.Length > 19000)
                        {
                            queryList.Add(queryString + queryParams + ")");
                            queryParams = string.Empty;
                        }
                    }

                    if (queryParams.Length > 0)
                    {
                        queryList.Add(queryString + queryParams + ")");
                        queryParams = string.Empty;
                    }


                    foreach (string query in queryList)
                    {
                        queryComplete = false;
                        //run query
                        qResult = SforceService.query(query);
                        //check results, if there are results
                        if (qResult.size > 0)
                        {
                            while (!queryComplete)
                            {
                                sObject[] records = qResult.records;
                                for (int i = 0; i < records.Length; i++)
                                {
                                    Deal__c d = (Deal__c)records[i];
                                    SFActivityModel sfActivity = new SFActivityModel
                                    {
                                        ActivityId = d.MNet_Activity_ID__c,
                                        SalesforceId = d.Id,
                                        OwnerId = d.Owner.Id
                                    };
                                    sfActivities.Add(sfActivity);
                                }
                                if (qResult.done) queryComplete = true;
                                else qResult = SforceService.queryMore(qResult.queryLocator);
                            }
                            
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error - Activity.cs - GetRecordsFromSalesForce()",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                Success = false;
            }
            return sfActivities;
        }

        /// <summary>
        /// Step 3: Loop through salesforce matching list
        ///         Check each resulting record to see if user has access to
        ///         If has access, add into list that will be ready for update
        ///         Break resulting list in chucks of 200 since that is the limit for each update call
        ///         Create salesforce object and update to salesforce
        /// </summary>
        /// <param name="matchingModifiedSalesforceRecords"></param>
        private void Update(List<SFActivityModel> matchingModifiedSalesforceRecords)
        {
            
            try
            {
                List<Deal__c> listObjects = new List<Deal__c>();
                MarcusMillichap_SF_API api = new MarcusMillichap_SF_API();
                string adUserId = string.Empty;
                //loop through salesforce records, add all valid record into list to be ready for updated
                for (int i = 0; i < matchingModifiedSalesforceRecords.Count; i++)
                {
                    if ((i % 800) == 0)
                    {
                        SforceService = api.LogIn();
                    }
                    MarcusMillichap_User user = MarcusMillichap_Common.GetMatchingUserInformation(matchingModifiedSalesforceRecords[i].OwnerId, UserList);
                    if (user != null && !string.IsNullOrEmpty(user.ADUserId)) {

                        MMActivityModel activity = CheckActivityAccessAndGetData(matchingModifiedSalesforceRecords[i].ActivityId, user.ADUserId);

                        if (activity != null && !string.IsNullOrEmpty(activity.ActivityId))
                        {
                            var sfDeal = SfDealMapper.Map(matchingModifiedSalesforceRecords[i].SalesforceId, activity);
                            listObjects.Add(sfDeal);
                        }
                    }
                }

                //loop through list to update in chucks of 200
                var errors = new List<MarcusMillichap_Errors>();
                var success = new Dictionary<string, sObject>();
                List<List<Deal__c>> list = MarcusMillichap_Common.BreakIntoChunks(listObjects, Convert.ToInt32(ConfigurationManager.AppSettings["MAX_NUMBER_UPDATE"]));
                foreach (var subList in list)
                {
                    int index = 0;
                    Dictionary<string, sObject> temp = new Dictionary<string, sObject>();
                    sObject[] sArray = new sObject[subList.Count];
                    foreach (var item in subList)
                    {
                        sArray[index] = item;
                        index++;
                    }
                    if (sArray != null && sArray.Length > 0)
                    {
                        api.Update(sArray, ref SforceService, out success, out errors);
                        success = success.Concat(temp).ToDictionary(x => x.Key, x => x.Value);
                    }
                }


            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error - Activity.cs - Update()",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                Success = false;
            }
        }

        private void SyncModifiedActivitiesInSalesforce()
        {
            List<MarcusMillichap_Errors> errors = new List<MarcusMillichap_Errors>();
            Dictionary<string, sObject> success = new Dictionary<string, sObject>();
            try
            {
                QueryResult qResult = null;
                bool queryComplete = false;

                List<SFActivityModel> sfActivities = new List<SFActivityModel>();
                string query = "SELECT id, MNet_Activity_ID__c, Owner.Id FROM Deal__c ";
                query += "WHERE LastModifiedById !='" + ConfigurationManager.AppSettings["APIUserSalesforceID"].ToString() + "' ";
                query += "AND MNet_Activity_ID__c != null ";
                query += "AND (Street__c = null OR City__c = null)";

                qResult = SforceService.query(query);
                if (qResult.size > 0)
                {
                    while (!queryComplete)
                    {
                        sObject[] records = qResult.records;
                        for (int i = 0; i < records.Length; i++)
                        {
                            Deal__c d = (Deal__c)records[i];
                            SFActivityModel sfActivity = new SFActivityModel
                            {
                                ActivityId = d.MNet_Activity_ID__c,
                                SalesforceId = d.Id,
                                OwnerId = d.Owner.Id
                            };
                            sfActivities.Add(sfActivity);
                        }
                        if (qResult.done) queryComplete = true;
                        else qResult = SforceService.queryMore(qResult.queryLocator);
                    }
                }

                if (sfActivities.Count > 0)
                {
                    List<Deal__c> listObjects = new List<Deal__c>();
                    foreach (SFActivityModel sfActivity in sfActivities)
                    {
                        string adUserId = string.Empty;
                        MarcusMillichap_User user = MarcusMillichap_Common.GetMatchingUserInformation(sfActivity.OwnerId, UserList);
                        
                        if(user != null && !string.IsNullOrEmpty(user.ADUserId))
                        {
                            MMActivityModel activity = CheckActivityAccessAndGetData(sfActivity.ActivityId, adUserId);
                            if (activity != null && !string.IsNullOrEmpty(activity.ActivityId))
                            {
                                var sfDeal = SfDealMapper.Map(sfActivity.SalesforceId, activity);
                                listObjects.Add(sfDeal);
                            }
                        }
                    }

                    //loop through list to update in chucks of 200
                    List<List<Deal__c>> list = MarcusMillichap_Common.BreakIntoChunks(listObjects, Convert.ToInt32(ConfigurationManager.AppSettings["MAX_NUMBER_UPDATE"]));
                    MarcusMillichap_SF_API api = new MarcusMillichap_SF_API();
                    foreach (var subList in list)
                    {
                        int index = 0;
                        Dictionary<string, sObject> temp = new Dictionary<string, sObject>();
                        sObject[] sArray = new sObject[subList.Count];
                        foreach (var item in subList)
                        {
                            sArray[index] = item;
                            index++;
                        }
                        if (sArray != null && sArray.Length > 0)
                        {
                            api.Update(sArray, ref SforceService, out success, out errors);
                            success = success.Concat(temp).ToDictionary(x => x.Key, x => x.Value);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error - Activity.cs - Update()",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                Success = false;
            }
        }

        /// <summary>
        /// Helper Function: Check if user has access to activity
        /// </summary>
        /// <param name="activityId"></param>
        /// <param name="userId"></param>
        /// <returns></returns>
        private MMActivityModel CheckActivityAccessAndGetData(string activityId, string userId)
        {
            //add code here to call sp_search
            DataProvider dp = new DataProvider(ConfigurationManager.ConnectionStrings["GeoSearch"].ToString());
            NameValueCollection nv = new NameValueCollection();
            DataSet ds = new DataSet();
            MMActivityModel activity = new MMActivityModel();
            try
            {
                if (!string.IsNullOrEmpty(userId) && !string.IsNullOrEmpty(activityId))
                {
                    nv.Add("@UserID", userId);
                    nv.Add("@Activity_ActivityID", activityId);
                    nv.Add("@ActivitySearch", bool.TrueString);
                    nv.Add("@GetResults", bool.TrueString);
                    nv.Add("@IsSalesForceCall", bool.TrueString);
                    ds = dp.GetDataSet("[dbo].[spSearch_FE]", CommandType.StoredProcedure, nv);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {

                        activity.ActivityId = activityId;
                        activity.Name = MarcusMillichap_Common.GetPropertyNameOrAddress(ds.Tables[0].Rows[0]["PropertyName"].ToString(), ds.Tables[0].Rows[0]["Address1"].ToString());
                        activity.City = ds.Tables[0].Rows[0]["City"].ToString();
                        activity.State = ds.Tables[0].Rows[0]["State"].ToString();
                        activity.Status = ds.Tables[0].Rows[0]["StatusSuperGroup"].ToString();
                        activity.Street = MarcusMillichap_Common.GetPropertyAddress(ds.Tables[0].Rows[0]["Address1"].ToString(), ds.Tables[0].Rows[0]["Address2"].ToString());
                        activity.SubStatus = ds.Tables[0].Rows[0]["Status"].ToString();
                        activity.Zip = ds.Tables[0].Rows[0]["Zip5"].ToString();
                        activity.Price = MarcusMillichap_Common.FormatPrice(ds.Tables[0].Rows[0]["Price"]);
                        if (!string.IsNullOrEmpty(ds.Tables[0].Rows[0]["ApprxYearBuilt"].ToString()))
                            activity.YearBuilt = int.Parse(ds.Tables[0].Rows[0]["ApprxYearBuilt"].ToString());
                        if (!string.IsNullOrEmpty(ds.Tables[0].Rows[0]["PropertyID_MPD"].ToString()))
                            activity.MPDId = int.Parse(ds.Tables[0].Rows[0]["PropertyID_MPD"].ToString());
                        activity.StateIntl = ds.Tables[0].Rows[0]["State_INTL"].ToString();
                        activity.Country = ds.Tables[0].Rows[0]["Country"].ToString();
                        activity.Commission = MarcusMillichap_Common.FormatPrice(ds.Tables[0].Rows[0]["AgentIntermediateGrossCommission"]);
                        activity.TotalCommission = MarcusMillichap_Common.FormatPrice(ds.Tables[0].Rows[0]["ActivityGrossCommission"]);
                    }
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                     ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                     "SalesForce API Error - Activity.cs - CheckActivityAccessAndGetData()",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
            }
            finally { dp.Dispose(); nv = null; }
            return activity;
        }
    }
}
