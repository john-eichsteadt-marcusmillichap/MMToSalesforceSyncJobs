using SalesforceSyncJobs.wsSalesforce;
using SalesforceSyncJobs.wsSalesforceAPEX;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace SalesforceSyncJobs
{
    class MarcusMillichap_GroupMember
    {
        private int? _userCode;
        private string _id, _salesForceUserId, _groupId, _groupName, _developerName, _adUserId;
        private wsSalesforce.SforceService _service;

        public int? UserCode { get { return _userCode; } set { _userCode = value; } }
        public string ID { get { return _id; } set { _id = value; } }
        public string ADUserId { get { return _adUserId; } set { _adUserId = value; } }
        public string SalesForceUserId { get { return _salesForceUserId; } set { _salesForceUserId = value; } }
        public string GroupId { get { return _groupId; } set { _groupId = value; } }
        public string GroupName { get { return _groupName; } set { _groupName = value; } }
        public string DeveloperName { get { return _developerName; } set { _developerName = value; } }

        public MarcusMillichap_GroupMember() {
            _id = string.Empty;
            _adUserId = string.Empty;
            _salesForceUserId = string.Empty;
            _groupId = string.Empty;
            _groupName = string.Empty;
            _developerName = string.Empty;
            _userCode = null;
        }

        public MarcusMillichap_GroupMember(wsSalesforce.SforceService service)
        {
            _service = service;
            _id = string.Empty;
            _adUserId = string.Empty;
            _salesForceUserId = string.Empty;
            _groupId = string.Empty;
            _groupName = string.Empty;
            _developerName = string.Empty;
            _userCode = null;
        }

        public void GetAllGroupMembers(ref Dictionary<string, MarcusMillichap_Group> groups, Dictionary<string, MarcusMillichap_User> sfUsers)
        {
            Dictionary<string, MarcusMillichap_Group> groupsDict = new Dictionary<string, MarcusMillichap_Group>();
            List<MarcusMillichap_GroupMember> GroupMembersList = new List<MarcusMillichap_GroupMember>();
            try
            {
                wsSalesforce.QueryResult qResult = null;
                string query = string.Empty;
                string groupName = string.Empty;
                bool queryComplete = false;
                //query
                query += "SELECT Id, UserOrGroupId, Group.DeveloperName, Group.Type, Group.DoesIncludeBosses, Group.Name, Group.Id ";
                query += "FROM GroupMember ";
                query += "WHERE GroupId IN (SELECT Id FROM Group) ";
                query += "ORDER BY Group.DeveloperName";
                //execute query
                qResult = _service.query(query);
                if (qResult.size > 0)
                {
                    while (!queryComplete)
                    {
                        wsSalesforce.sObject[] records = qResult.records;
                        //Groups g = new Groups();
                        for (int i = 0; i < records.Length; i++)
                        {
                            wsSalesforce.GroupMember gm = (wsSalesforce.GroupMember)records[i];
                            string adUserId = MarcusMillichap_Common.GetMNetADUserIDFromSalesforceUsers(gm.UserOrGroupId, sfUsers);
                            //if (adUserId == "mgjonbalaj")
                            //{
                            //    var a = "b";
                            //}
                            if (groups.ContainsKey(gm.Group.DeveloperName))
                            {
                                MarcusMillichap_Group g = groups[gm.Group.DeveloperName];
                                if (!string.IsNullOrEmpty(adUserId) && !g.GroupMembers.ContainsKey(adUserId.ToLower()))
                                {
                                    g.GroupMembers.Add(adUserId.ToLower(), new MarcusMillichap_GroupMember
                                    {
                                        ID = gm.Id,
                                        GroupId = gm.GroupId,
                                        DeveloperName = gm.Group.DeveloperName,
                                        GroupName = gm.Group.Name,
                                        SalesForceUserId = gm.UserOrGroupId
                                    });
                                }
                            }
                            else
                            {
                                MarcusMillichap_Group g = new MarcusMillichap_Group();
                                g.ID = gm.Group.Id;
                                g.DeveloperName = gm.Group.DeveloperName;
                                g.DoesIncludeBosses = gm.Group.DoesIncludeBosses;
                                g.Type = gm.Group.Type;
                                g.GroupName = gm.Group.Name;
                                g.GroupMembers.Add(adUserId.ToLower(), new MarcusMillichap_GroupMember
                                {
                                    ID = gm.Id,
                                    GroupId = gm.GroupId,
                                    DeveloperName = gm.Group.DeveloperName,
                                    GroupName = gm.Group.Name,
                                    SalesForceUserId = gm.UserOrGroupId
                                });
                                groups.Add(gm.Group.DeveloperName, g);
                            }
                        }
                        if (qResult.done) queryComplete = true;
                        else qResult = _service.queryMore(qResult.queryLocator);
                    }
                }
            }
            catch (Exception ex)
            {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                    ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                   "SalesForce API Error - Users.cs - Load()",
                                    "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                    "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                    "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                    true);
                e.SendEmail();
            }
        }

        /// <summary>
        /// Compare Users in salesforce and mnet
        /// </summary
        /// <param name="salesforceUsers"></param>
        /// <param name="mnetUsers"></param>
        /// <param name="missingUsersInSalesforce"></param>
        /// <param name="extraUsersInSalesforce"></param>
        /// Dictionary<int, Users> salesforceUsers - int -> usercode
        /// Dictionary<int, Users> mnetUsers - int -> usercode
        public void Compare(string groupId, Dictionary<string, MarcusMillichap_GroupMember> salesforceGroupMembers, Dictionary<string, MarcusMillichap_GroupMember> mmGroupMembers, out Dictionary<string, MarcusMillichap_GroupMember> missingUsersGroupMembersSalesforce, out Dictionary<string, MarcusMillichap_GroupMember> extraGroupMembersInSalesforce)
        {
            missingUsersGroupMembersSalesforce = new Dictionary<string, MarcusMillichap_GroupMember>();
            extraGroupMembersInSalesforce = new Dictionary<string, MarcusMillichap_GroupMember>();

            var missing = mmGroupMembers.Keys.Except(salesforceGroupMembers.Keys);
            var extra = salesforceGroupMembers.Keys.Except(mmGroupMembers.Keys);

            foreach (var m in missing)
            {
                if (!missingUsersGroupMembersSalesforce.ContainsKey(m))
                {
                    mmGroupMembers[m].GroupId = groupId;
                    missingUsersGroupMembersSalesforce.Add(m, mmGroupMembers[m]);
                }
            }

            foreach (var e in extra)
            {
                if (!extraGroupMembersInSalesforce.ContainsKey(e))
                {
                    extraGroupMembersInSalesforce.Add(e, salesforceGroupMembers[e]);
                }
            }
        }

        public void AddMember(List<MarcusMillichap_GroupMember> memberList)
        {
            if (memberList != null && memberList.Count > 0)
            {
                int errorCount = 0;
                MarcusMillichap_APEX_API apex = new MarcusMillichap_APEX_API();
                wsSalesforceAPEX.SObjectServiceService s = apex.ConnectToApexAPI(_service);
                wsSalesforceAPEX.sObject[] records = new wsSalesforceAPEX.sObject[memberList.Count];
                //loop through member list
                foreach (var g in memberList)
                {
                    if (!string.IsNullOrEmpty(g.SalesForceUserId))
                    {
                        wsSalesforceAPEX.GroupMember apexGM = new wsSalesforceAPEX.GroupMember();
                        apexGM.GroupId = g.GroupId;
                        apexGM.UserOrGroupId = g.SalesForceUserId;
                        records[memberList.IndexOf(g)] = apexGM;
                    }
                    else { 
                        //records.
                        //records[memberList.IndexOf(g)].remo
                    }
                }

                wsSalesforceAPEX.SaveResult[] results = s.doCreate(records);
                foreach (wsSalesforceAPEX.SaveResult result in results)
                {
                    if (result.success.HasValue && result.success.Value)
                    {
                        //save success
                    }
                    else
                    {
                        errorCount++;
                    }
                }
            }
        }

        public void RemoveMember(List<string> memberIds)
        {
            if (memberIds != null && memberIds.Count > 0)
            {
                int errorCount = 0;
                MarcusMillichap_APEX_API apex = new MarcusMillichap_APEX_API();
                wsSalesforceAPEX.SObjectServiceService s = apex.ConnectToApexAPI(_service);
                string[] records = new string[memberIds.Count];
                //loop through member list
                foreach (var g in memberIds)
                {
                    records[memberIds.IndexOf(g)] = g;
                }

                wsSalesforceAPEX.DeleteResult[] results = s.doDelete(records);
                foreach (wsSalesforceAPEX.DeleteResult result in results)
                {
                    if (result.success.HasValue && result.success.Value)
                    {
                        //save success
                    }
                    else
                    {
                        errorCount++;
                    }
                }
            }
        }

        public static Dictionary<string, MarcusMillichap_GroupMember> GetAllGroupMembersByGroupName(string groupName, Dictionary<string, MarcusMillichap_Group> mmGroupsMasterList) {
            Dictionary<string, MarcusMillichap_GroupMember> gList = new Dictionary<string, MarcusMillichap_GroupMember>();
            if (mmGroupsMasterList != null && mmGroupsMasterList.Count > 0) {
                var m = mmGroupsMasterList[groupName];
                if (m.GroupMembers != null)
                    return m.GroupMembers;
            }
            return gList;
        }
    }
}
