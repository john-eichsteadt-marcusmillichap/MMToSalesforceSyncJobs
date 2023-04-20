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

namespace SalesforceSyncJobs.Services
{
    public class PropertyService
    {
        private SforceService SforceService;

        public PropertyService(SforceService sforceService)
        {
            SforceService = sforceService;
        }

        public bool SyncMPdToSalesforce(DateTime lastRunTime)
        {
            try
            {
                var modifiedRecordsFromMpd = GetRecordsFromMPD(lastRunTime);
                var matchingModifiedSalesforceRecords = GetRecordsFromSalesForce(modifiedRecordsFromMpd);

                Update(matchingModifiedSalesforceRecords);

                return true;
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error - PropertyService.cs - SyncMPdToSalesforce() ",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                return false;
            }
        }

        public bool SyncPropertyRecordsToSQL(DateTime lastRunTime)
        {
           
            try
            {
                List<MMPropertyModel> properties = new List<MMPropertyModel>();
                //build sql query
                string query = "SELECT Id, MNet_ID__c, MNETUSERCode__c  ";
                query += "FROM Property__c ";
                query += "WHERE (CreatedDate >= " + MarcusMillichap_Common.ConvertToSalesforceDateTime(lastRunTime) + " ";
                query += "OR LastModifiedDate >= " + MarcusMillichap_Common.ConvertToSalesforceDateTime(lastRunTime) + ") ";
                query += "AND LastModifiedById != '" + ConfigurationManager.AppSettings["MNetDbaSalesForceID"] + "' ";
                query += "AND LastModifiedById != '" + ConfigurationManager.AppSettings["APIUserSalesforceID"] + "' ";

                bool queryComplete = false;
                //run query
                QueryResult qResult = SforceService.query(query);
                //check results, if there are results
                if (qResult.size > 0)
                {
                    while (!queryComplete)
                    {
                        sObject[] records = qResult.records;
                        for (int i = 0; i < records.Length; i++)
                        {
                            Property__c sfP = (Property__c)records[i];
                            properties.Add(new MMPropertyModel
                            {
                                SalesForceId = sfP.Id,
                                MpdId = sfP.MNet_ID__c,
                                UserCode = !string.IsNullOrEmpty(sfP.MNETUSERCode__c)
                                            ? Convert.ToInt32(sfP.MNETUSERCode__c)
                                            : 0
                            });
                        }
                        if (qResult.done) queryComplete = true;
                        else qResult = SforceService.queryMore(qResult.queryLocator);
                    }
                }

                var failedProperties = new List<MMPropertyModel>();

                if (properties != null && properties.Count > 0)
                {
                    foreach (MMPropertyModel property in properties)
                    {
                        if (!CreateSalesforcePropertyIDIntoMPD(property))
                        {
                            failedProperties.Add(property);
                        }
                    }
                }


                // retry missed ones
                if(failedProperties.Count > 0)
                {
                    foreach(MMPropertyModel property in failedProperties)
                    {
                        CreateSalesforcePropertyIDIntoMPD(property);
                    }
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                   "SalesForce API Error - Property.cs - SyncPropertyRecordsToSQL()",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                return false;
            }

            return true;
        }


        /// <summary>
        /// Step 1: Get records that was last changed from MPD
        /// </summary>
        /// <param name="lastRunTime"></param>
        /// <returns></returns>
        private List<MMPropertyModel> GetRecordsFromMPD(DateTime lastRunTime)
        {
            DataProvider dp = new DataProvider(ConfigurationManager.ConnectionStrings["MPD"].ToString());
            NameValueCollection nv = new NameValueCollection();
            NameValueCollection nvOut = new NameValueCollection();
            NameValueCollection results = new NameValueCollection();
            DataSet ds = new DataSet();
            List<MMPropertyModel> properties = new List<MMPropertyModel>();
            try
            {
                nv.Add("@Time_PST", lastRunTime.ToString());
                nvOut.Add("@ErrCode", "4");
                nvOut.Add("@ErrMessage", "2048");
                ds = dp.GetDataSetMultipleOutputParam("[PropertyDetail].[spProperty_Search_ModifiedAfterTime]", CommandType.StoredProcedure, nv, nvOut, out results);
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0 && results["@ErrMessage"].ToString() == string.Empty)
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        MMPropertyModel property = new MMPropertyModel
                        {
                            MpdId = dr["PropertyID"].ToString()
                        };
                        properties.Add(property);
                    }
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                   "SalesForce API Error - PropertyService.cs - GetRecordsFromMPD()",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                throw;
            }
            finally { dp.Dispose(); nv = null; }
            return properties;
        }

        /// <summary>
        /// Step 2: Based on records from MPD, query SalesForce to get list of matching records that was changed
        /// </summary>
        /// <param name="mpdProperty"></param>
        /// <param name="service"></param>
        /// <returns></returns>
        private List<MMPropertyModel> GetRecordsFromSalesForce(List<MMPropertyModel> mpdProperty)
        {

            ArrayList queryArray = new ArrayList();
            List<MMPropertyModel> returnList = new List<MMPropertyModel>();
            try
            {
                if (mpdProperty.Count > 0)
                {
                    QueryResult qResult = null;
                    string queryParams = string.Empty;
                    int count = 0;
                    bool queryComplete = false;
                    //build sql query
                    string query = "SELECT id, MNet_ID__c, Owner.id ";
                    query += "FROM Property__c ";
                    query += "WHERE MNet_ID__c IN (";
                    foreach (var item in mpdProperty)
                    {
                        if (queryParams.Equals(string.Empty)) queryParams += "'" + item.MpdId + "'";
                        else queryParams += ",'" + item.MpdId + "'";
                        count++;

                        if (queryParams.Length > 19000)
                        {
                            queryArray.Add(query + queryParams + ")");
                            queryParams = string.Empty;
                        }
                    }

                    if (queryParams.Length > 0)
                    {
                        queryArray.Add(query + queryParams + ")");
                        queryParams = string.Empty;
                    }

                    foreach (string queryItem in queryArray)
                    {
                        queryComplete = false;
                        //run query
                        qResult = SforceService.query(queryItem);
                        //check results, if there are results
                        if (qResult.size > 0)
                        {
                            while (!queryComplete)
                            {
                                sObject[] records = qResult.records;
                                for (int i = 0; i < records.Length; i++)
                                {
                                    Property__c sfP = (Property__c)records[i];
                                    MMPropertyModel property = new MMPropertyModel
                                    {
                                        SalesForceId = sfP.Id,
                                        MpdId = sfP.MNet_ID__c,
                                        OwnerId = sfP.Owner.Id
                                    };
                                    returnList.Add(property);
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
                                   "SalesForce API Error - PropertyService.cs - GetRecordsFromSalesForce()",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                throw;
            }
            return returnList;
        }

        /// <summary>
        /// Step 3: Loop through salesforce matching list
        ///         Check each resulting record to see if user has access to
        ///         If has access, add into list that will be ready for update
        ///         Break resulting list in chucks of 200 since that is the limit for each update call
        ///         Create salesforce object and update to salesforce
        /// </summary>
        /// <param name="salesforce"></param>
        /// <param name="service"></param>
        private void Update(List<MMPropertyModel> salesforce)
        {
           
            try
            {

                List<Property__c> listObjects = new List<Property__c>();
                MarcusMillichap_SF_API api = new MarcusMillichap_SF_API();
                //loop through salesforce records, add all valid record into list to be ready for updated
                for (int i = 0; i < salesforce.Count; i++)
                {
                    if ((i % 800) == 0)
                    {
                        SforceService = api.LogIn();
                    }

                    MMPropertyModel property = CheckPropertyAccessAndGetData(salesforce[i].MpdId);
                    if (property != null && !string.IsNullOrEmpty(property.MpdId))
                    {
                        ArrayList nullFields = new ArrayList();

                        Property__c sfProperty = new Property__c
                        {
                            Id = salesforce[i].SalesForceId,
                            MNet_ID__c = property.MpdId
                        };

                        if (property.NumberofUnits > 0)
                        {
                            sfProperty.of_Units__c = property.NumberofUnits;
                            sfProperty.of_Units__cSpecified = true;
                        }
                        else nullFields.Add("of_Units__c");

                        if (!string.IsNullOrEmpty(property.SubType)) sfProperty.Property_Sub_Type__c = property.SubType;
                        else nullFields.Add("Property_Sub_Type__c");

                        if (property.RentableSqFt.HasValue && property.RentableSqFt > 0)
                        {
                            sfProperty.Rentable_SF__c = property.RentableSqFt;
                            sfProperty.Rentable_SF__cSpecified = true;
                        }
                        else nullFields.Add("Rentable_SF__c");

                        if (!string.IsNullOrEmpty(property.StateId)) sfProperty.State__c = property.StateId;
                        else nullFields.Add("State__c");

                        if (!string.IsNullOrEmpty(property.City)) sfProperty.City__c = property.City;
                        else nullFields.Add("City__c");

                        if (!string.IsNullOrEmpty(property.Name)) sfProperty.Name = property.Name;
                        else nullFields.Add("Name");

                        if (!string.IsNullOrEmpty(property.Address)) sfProperty.address__c = property.Address;
                        else nullFields.Add("address__c");

                        if (property.YearBuilt.HasValue && property.YearBuilt > 0) sfProperty.Year_Built__c = property.YearBuilt.ToString();
                        else nullFields.Add("Year_Built__c");

                        if (!string.IsNullOrEmpty(property.Country)) sfProperty.Country__c = property.Country;
                        else nullFields.Add("Country__c");

                        if (!string.IsNullOrEmpty(property.County)) sfProperty.County__c = property.County;
                        else nullFields.Add("County__c");

                        if (!string.IsNullOrEmpty(property.StateIntl)) sfProperty.International_State__c = property.StateIntl;
                        else nullFields.Add("International_State__c");

                        if (!string.IsNullOrEmpty(property.Zip)) sfProperty.Zip__c = property.Zip;
                        else nullFields.Add("Zip__c");

                        if (nullFields.Count > 0)
                        {
                            string[] s = new string[nullFields.Count];
                            for (int j = 0; j < nullFields.Count; j++)
                            {
                                s[j] = nullFields[j].ToString();
                            }
                            sfProperty.fieldsToNull = s;
                        }

                        listObjects.Add(sfProperty);
                    }
                }
                //loop through list to update in chucks of 200
                var errors = new List<MarcusMillichap_Errors>();
                var success = new Dictionary<string, sObject>();
                List<List<Property__c>> list = MarcusMillichap_Common.BreakIntoChunks(listObjects, Convert.ToInt32(ConfigurationManager.AppSettings["MAX_NUMBER_UPDATE"]));
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
                        api.Update(sArray, ref SforceService, out temp, out errors);
                        success = success.Concat(temp).ToDictionary(x => x.Key, x => x.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                   "SalesForce API Error - PropertyService.cs - Update() ",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                throw;
            }
        }

        /// <summary>
        /// Helper Function: Get property data from MPD to sync over
        /// </summary>
        /// <param name="mpdId"></param>
        /// <returns></returns>
        private MMPropertyModel CheckPropertyAccessAndGetData(string mpdId)
        {
            DataProvider dp = new DataProvider(ConfigurationManager.ConnectionStrings["MPD"].ToString());
            NameValueCollection nv = new NameValueCollection();
            DataSet ds = new DataSet();
            MMPropertyModel property = new MMPropertyModel();
            try
            {
                nv.Add("@PropertyID", mpdId);
                ds = dp.GetDataSet("[PropertyDetail].[spGetMainInfo]", CommandType.StoredProcedure, nv);
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    DataRow dr = ds.Tables[0].Rows[0];
                    int temp;
                    property.MpdId = dr["PropertyID"].ToString();
                    property.Name = MarcusMillichap_Common.GetPropertyNameOrAddress(dr["PropertyName"].ToString(), dr["Address"].ToString());
                    property.NumberofUnits = int.TryParse(dr["Units"].ToString(), out temp) ? temp : 0;
                    property.RentableSqFt = int.TryParse(dr["SqFt"].ToString(), out temp) ? temp : 0;
                    property.StateId = dr["StateID"].ToString();
                    property.StateIntl = dr["State_INTL"].ToString();
                    property.Country = dr["Country"].ToString();
                    property.SubType = dr["PropertyType"].ToString();
                    property.City = dr["City"].ToString();
                    property.County = dr["County"].ToString();
                    property.Address = MarcusMillichap_Common.GetPropertyAddress(dr["Address"].ToString(), dr["Address2"].ToString());
                    if (!string.IsNullOrEmpty(dr["ApprxYearBuilt"].ToString()))
                        property.YearBuilt = Convert.ToInt16(dr["ApprxYearBuilt"]);
                    property.Zip = dr["ZipCode"].ToString();
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error - Property.cs - CheckPropertyAccessAndGetData()",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
                throw;
            }
            finally { dp.Dispose(); nv = null; }
            return property;
        }

        private bool CreateSalesforcePropertyIDIntoMPD(MMPropertyModel property)
        {

            DataProvider dp = new DataProvider(ConfigurationManager.ConnectionStrings["MPD"].ToString());
            NameValueCollection nv = new NameValueCollection();
            bool success = true;
            if (!string.IsNullOrEmpty(property.MpdId) 
                && !string.IsNullOrEmpty(property.SalesForceId) 
                && property.UserCode > 0)
            {
                try
                {
                    nv.Add("@PropertyID", property.MpdId);
                    nv.Add("@UserCode", property.UserCode.ToString());
                    nv.Add("@SFPropertyID", property.SalesForceId);
                    success = dp.DoNonQuery("[dbo].[spSFUserProperty_Insert]", CommandType.StoredProcedure, nv);
                }
                catch
                {
                    success = false;
                }
                finally { dp.Dispose(); nv = null; }
            }
            return success;
        }
    }
}
