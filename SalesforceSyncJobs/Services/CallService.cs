using SalesforceSyncJobs.wsSalesforce;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;

namespace SalesforceSyncJobs
{
    public class CallService
    {
        private SforceService SforceService;

        public CallService(SforceService sforceService)
        {
            SforceService = sforceService;
        }

        public bool SyncToSalesforce(DateTime from, DateTime to)
        {
            DataProvider dp = new DataProvider(ConfigurationManager.ConnectionStrings["MMPeople"].ToString());
            NameValueCollection nv = new NameValueCollection();
            DataSet ds = new DataSet();
            List<wsSalesforce.Task> tasks = new List<wsSalesforce.Task>();
            List<MarcusMillichap_Errors> errors = new List<MarcusMillichap_Errors>();
            bool success = true;
            try
            {
                nv.Add("@DateTimeConnect_From", from.ToString());
                nv.Add("@DateTimeConnect_To", to.ToString());
                ds = dp.GetDataSet("[MMPeopleApp].[spRptVOIPCallsSyncToSF]", CommandType.StoredProcedure, nv);
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        wsSalesforce.Task task = new wsSalesforce.Task
                        {
                            WhoId = dr["SFContactID"].ToString(),
                            Subject = "Call",
                            OwnerId = dr["SFOwnerID"].ToString(),
                            Status = "Completed",
                            CallID__c = int.Parse(dr["CallID"].ToString()),
                            CallID__cSpecified = true,
                            CallID_External__c = dr["CallID"].ToString() + dr["SFContactID"].ToString(),
                            Successful_Contact__c = true,
                            Successful_Contact__cSpecified = true
                        };

                        int iTemp;
                        if (int.TryParse(dr["Duration"].ToString(), out iTemp))
                        {
                            task.CallDurationInSeconds = iTemp;
                            task.CallDurationInSecondsSpecified = true;
                        }

                        DateTime dTemp;
                        if (DateTime.TryParse(dr["DateTimeConnect"].ToString(), out dTemp))
                        {
                            task.ActivityDate = dTemp;
                            task.ActivityDateSpecified = true;
                        }

                        if (!string.IsNullOrEmpty(dr["SFAssignedToID"].ToString())) { task.Assigned_to_ID__c = dr["SFAssignedToID"].ToString(); }
                        if (!string.IsNullOrEmpty(dr["AssignedTo_DisplayName"].ToString())) { task.Assigned_To__c = dr["AssignedTo_DisplayName"].ToString(); }
                        if (bool.Parse(dr["Incoming"].ToString())) { task.Description = "Incoming"; }
                        if (bool.Parse(dr["Outgoing"].ToString())) { task.Description = "Outgoing"; }
                        task.Phone__c = MarcusMillichap_Common.FormatPhoneNumberDisplay(dr["PhoneNumber"].ToString());
                        tasks.Add(task);
                    }
                }

                if (tasks != null && tasks.Count > 0)
                {
                    MarcusMillichap_SF_API api = new MarcusMillichap_SF_API();
                    List<List<wsSalesforce.Task>> list = MarcusMillichap_Common.BreakIntoChunks(tasks, Convert.ToInt32(ConfigurationManager.AppSettings["MAX_NUMBER_UPDATE"].ToString()));
                    foreach (var subList in list)
                    {
                        int index = 0;
                        sObject[] sArray = new sObject[subList.Count];
                        Dictionary<string, sObject> temp = new Dictionary<string, sObject>();
                        foreach (var item in subList)
                        {
                            sArray[index] = item;
                            index++;
                        }
                        if (sArray != null && sArray.Length > 0)
                        {
                            api.Upsert("CallID_External__c", sArray, ref SforceService, out errors);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                   ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                   "SalesForce API Error - CallService.cs - SyncToSalesforce() ",
                                   "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                   "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                   "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                   true);
                e.SendEmail();
                return false;
            }
            finally {
                if(ds != null) { ds.Dispose(); }
                dp.Dispose();
                nv = null;
            }

            return success;
        }

    }
}
