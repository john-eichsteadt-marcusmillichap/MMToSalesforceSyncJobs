using SalesforceSyncJobs.Constants;
using SalesforceSyncJobs.Helpers;
using SalesforceSyncJobs.Services;
using SalesforceSyncJobs.wsSalesforce;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;

namespace SalesforceSyncJobs
{
    class Program
    {
        static void Main(string[] args)
        {
            try {

                string argument = args.Length > 0 ? args[0] : string.Empty;

                //get current datetime
                DateTime currentDateTime = DateTime.Now;
                //initialize marcus millichap salesforce helper class
                MarcusMillichap_SF_API helper = new MarcusMillichap_SF_API();
                //log into salesforce
                SforceService service = helper.LogIn();
                //check if login is successful
                if (service != null)
                {
                    if (argument.Equals("-property"))
                    {
                        PropertyService propertyService = new PropertyService(service);
                        // sync added properties from salesforce to database
                        if (propertyService.SyncPropertyRecordsToSQL(LastRunTime.Get(currentDateTime, ConfigurationManager.AppSettings[AppConfigKeys.MPDSyncLastRunTime])))
                        {
                            MarcusMillichap_Common.UpdateLastRunTime(currentDateTime, AppConfigKeys.MPDSyncLastRunTime);
                        }

                        //sync modified Properties
                        if (propertyService.SyncMPdToSalesforce(LastRunTime.Get(currentDateTime, ConfigurationManager.AppSettings[AppConfigKeys.PropertySyncLastRunTime])))
                        {
                            MarcusMillichap_Common.UpdateLastRunTime(currentDateTime, AppConfigKeys.PropertySyncLastRunTime);
                        }
                    }
                    else
                    {

                        #region Sync Marcus Millichap Users Into Salesforce

                        Dictionary<string, MarcusMillichap_User> _salesforceUsersMasterList = new Dictionary<string, MarcusMillichap_User>();
                        Dictionary<int, MarcusMillichap_User> _salesforceUserMasterListWithUserCodeKey = new Dictionary<int, MarcusMillichap_User>();
                        Dictionary<int, MarcusMillichap_User> _mmUsersMasterList = new Dictionary<int, MarcusMillichap_User>();
                        Dictionary<string, MarcusMillichap_Group> _salesforceGroupsMasterList = new Dictionary<string, MarcusMillichap_Group>();
                        Dictionary<string, MarcusMillichap_Group> _mmGroupsMasterList = new Dictionary<string, MarcusMillichap_Group>();
                        Dictionary<string, wsSalesforce.User> _sfTempDict = new Dictionary<string, wsSalesforce.User>();
                        Dictionary<int, MarcusMillichap_User> _missingUsers = new Dictionary<int, MarcusMillichap_User>();
                        Dictionary<int, MarcusMillichap_User> _extraUsers = new Dictionary<int, MarcusMillichap_User>();
                        Dictionary<string, MarcusMillichap_Group> _groupsToAdd = new Dictionary<string, MarcusMillichap_Group>();
                        Dictionary<string, MarcusMillichap_GroupMember> _removeMembersMasterList = new Dictionary<string, MarcusMillichap_GroupMember>();

                        List<string> _groupsToDelete = new List<string>();
                        List<MarcusMillichap_GroupMember> _addMembersMasterList = new List<MarcusMillichap_GroupMember>();
                        List<MarcusMillichap_Errors> _errors = new List<MarcusMillichap_Errors>();

                        //Intialize Marcus Millichap User object
                        MarcusMillichap_User user = new MarcusMillichap_User(service);
                        //Load all users/groups/group members into local objects
                        user.LoadAllMarcusMillichapUsersAndGroups();

                        //check if there are count > 0 since there will always be at least 1 active user
                        //if it's 0, an error occured while connecting to the datbase thus we do not want to run the sync
                        if (user.MMUsers.Count > 0)
                        {
                            //now that there are users in MMPeople, lets load all users from Salesforce to compare
                            user.LoadAllSalesforceUsers();
                            //set results into global variables
                            _salesforceUsersMasterList = user.SaleforceUsers;
                            _salesforceUserMasterListWithUserCodeKey = user.SalesforceUsersFromMnetComparison;
                            _mmUsersMasterList = user.MMUsers;
                            _mmGroupsMasterList = user.MMGroups;

                            //********************************** SYNC Users ********************************************************************************/

                            //now lets sync users into Salesforce
                            user.SyncToSalesforce(_salesforceUserMasterListWithUserCodeKey, _mmUsersMasterList, out _sfTempDict, out _errors);
                            //update master list with new users
                            MarcusMillichap_Common.AppendNewUsers(_sfTempDict, ref _salesforceUsersMasterList, ref _salesforceUserMasterListWithUserCodeKey);
                            //compare users and deactivate users in salesforce
                            user.Compare(_salesforceUserMasterListWithUserCodeKey, _mmUsersMasterList, out _missingUsers, out _extraUsers);
                            //deactivate users
                            user.Deactivate(_extraUsers, _salesforceUsersMasterList, out _errors);

                            //********************************** END SYNC Users *****************************************************************************/


                            //********************************** SYNC Groups ********************************************************************************/

                            //Initialize Marcus Millichap Group object
                            MarcusMillichap_Group g = new MarcusMillichap_Group(service);
                            //Load list of groups in Salesforce
                            g.LoadSalesforceGroups();
                            //set results into master Salesforce list
                            _salesforceGroupsMasterList = g.SalesforceGroups;
                            //compare groups
                            g.Compare(_salesforceGroupsMasterList, _mmGroupsMasterList, out _groupsToAdd, out _groupsToDelete);
                            //add missing groups
                            if (_groupsToAdd != null && _groupsToAdd.Count > 0)
                            {
                                Dictionary<string, MarcusMillichap_Group> groupsAdded = new Dictionary<string, MarcusMillichap_Group>();
                                g.Create(_groupsToAdd, out groupsAdded);
                                //update master list
                                _salesforceGroupsMasterList = MarcusMillichap_Common.AppendNewGroups(groupsAdded, _salesforceGroupsMasterList);
                            }

                            //********************************** END SYNC Groups ****************************************************************************/


                            //********************************** SYNC Group Members *************************************************************************/

                            //Initialize Marcus Millichap Group Member object
                            MarcusMillichap_GroupMember gm = new MarcusMillichap_GroupMember(service);
                            //Load all members in Salesforce
                            gm.GetAllGroupMembers(ref _salesforceGroupsMasterList, _salesforceUsersMasterList);
                            //Loop through all groups to compare it's members
                            foreach (var mmGroups in _mmGroupsMasterList)
                            {
                                if (_salesforceGroupsMasterList.ContainsKey(mmGroups.Key))
                                {
                                    Dictionary<string, MarcusMillichap_GroupMember> add = new Dictionary<string, MarcusMillichap_GroupMember>();
                                    Dictionary<string, MarcusMillichap_GroupMember> remove = new Dictionary<string, MarcusMillichap_GroupMember>();
                                    //Compare Groups
                                    gm.Compare(_salesforceGroupsMasterList[mmGroups.Key].ID, _salesforceGroupsMasterList[mmGroups.Key].GroupMembers, mmGroups.Value.GroupMembers, out add, out remove);
                                    //Add members to master update list
                                    if (add != null && add.Count > 0)
                                    {
                                        foreach (var a in add)
                                        {
                                            MarcusMillichap_GroupMember newGM = a.Value;
                                            if (a.Value.UserCode != null && a.Value.UserCode.HasValue && a.Value.UserCode.Value > 0)
                                            {
                                                newGM.SalesForceUserId = MarcusMillichap_Common.FindUserByUserCode(a.Value.UserCode.Value, _salesforceUsersMasterList);
                                            }
                                            if (!string.IsNullOrEmpty(newGM.SalesForceUserId))
                                            {
                                                _addMembersMasterList.Add(newGM);
                                            }
                                        }
                                    }
                                    //Remove members from master update list
                                    if (remove != null && remove.Count > 0)
                                    {
                                        foreach (var r in remove)
                                        {
                                            string key = remove[r.Key].ID;
                                            if (!_removeMembersMasterList.ContainsKey(key)) _removeMembersMasterList.Add(key, r.Value);
                                        }
                                    }
                                }
                            }

                            //Add Group Members
                            gm.AddMember(_addMembersMasterList);
                            //Remove Group Members
                            gm.RemoveMember(_removeMembersMasterList.Keys.ToList());

                            //********************************** End SYNC Group Members **********************************************************************/

                        }
                        else
                        {
                            //send error email, no users loaded ...
                            MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                            ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                           "SalesForce API Error - No Users Returned ",
                                            "<strong>Error Message:</strong> <br /><br />" + "The store procedure did not return any results for list of users and groups from People",
                                            true);
                            e.SendEmail();
                        }

                        #endregion

                        #region Sync to MM Database

                        //********************************** SYNC Marcus Millichap Database **********************************************************************/

                        //sync newly created/updated users in salesforce to local mm database
                        var usersLastRunTime = LastRunTime.Get(currentDateTime, ConfigurationManager.AppSettings[AppConfigKeys.UsersSyncLastRunTime])
                            .ToUniversalTime();

                        MarcusMillichap_User users = new MarcusMillichap_User(service);
                        if (users.SyncToMarcusMillichap(usersLastRunTime))
                        {
                            MarcusMillichap_Common.UpdateLastRunTime(currentDateTime, AppConfigKeys.UsersSyncLastRunTime);
                        }


                        //sync newly created/updated accounts in salesforce to local mm database
                        var accountsLastRunTime = LastRunTime.Get(currentDateTime, ConfigurationManager.AppSettings[AppConfigKeys.AccountsSyncLastRunTime])
                            .ToUniversalTime();

                        MarcusMillichap_Account accounts = new MarcusMillichap_Account(service);
                        if (accounts.SyncToMarcusMillichap(accountsLastRunTime))
                        {
                            MarcusMillichap_Common.UpdateLastRunTime(currentDateTime, AppConfigKeys.AccountsSyncLastRunTime);
                        }


                        //sync newly created/updated contacts in salesforce to local mm database
                        var contactsLastRunTime = LastRunTime.Get(currentDateTime, ConfigurationManager.AppSettings[AppConfigKeys.ContactsSyncLastRunTime])
                            .ToUniversalTime();

                        MarcusMillichap_Contact contacts = new MarcusMillichap_Contact(service);
                        if (contacts.SyncToMarcusMillichap(contactsLastRunTime))
                        {
                            MarcusMillichap_Common.UpdateLastRunTime(currentDateTime, AppConfigKeys.ContactsSyncLastRunTime);
                        }


                        //********************************** End SYNC Marcus Millichap Database **********************************************************************/

                        #endregion

                        #region Sync modified activities from mnet to salesforce

                        //********************************** SYNC Activity **************************************************************************/

                        var activityLastRunTime = LastRunTime.Get(currentDateTime, ConfigurationManager.AppSettings[AppConfigKeys.ActivitySyncLastRunTime]);

                        ActivityService activity = new ActivityService(service, _salesforceUsersMasterList);
                        if (activity.SyncToSalesforce(activityLastRunTime))
                        {
                            MarcusMillichap_Common.UpdateLastRunTime(currentDateTime, AppConfigKeys.ActivitySyncLastRunTime);
                        }

                        //********************************** End SYNC Activity **************************************************************************/

                        #endregion

                        #region Sync voip calls to salesforce

                        //********************************** SYNC VOIP Calls **************************************************************************/

                        // need UTC since VOIP stores in UTC. This needs to be run after we do the user/accounts/contact sync to our local db since
                        // it depends on that data to figure out what calls belongs to who and what we need to sync
                        var voipLastRunTime = LastRunTime.Get(currentDateTime, ConfigurationManager.AppSettings[AppConfigKeys.VoipSyncLastRunTime])
                            .ToUniversalTime();

                        CallService call = new CallService(service);
                        if (call.SyncToSalesforce(voipLastRunTime, currentDateTime.ToUniversalTime()))
                        {
                            MarcusMillichap_Common.UpdateLastRunTime(currentDateTime, AppConfigKeys.VoipSyncLastRunTime);
                        }

                        //********************************** End SYNC VOIP Calls **************************************************************************/

                        #endregion

                    }
                }
            }
            catch (Exception ex) {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                    "SalesForce API Error <br /><br />",
                                    "Error Message: <br /><br />" + ex.Message +
                                    "Inner Exception: <br /><br />" + ex.InnerException +
                                    "Stack Trace: <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
            }
        }
    }
}
