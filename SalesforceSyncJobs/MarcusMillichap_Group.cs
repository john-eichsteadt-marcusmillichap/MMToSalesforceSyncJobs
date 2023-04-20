using SalesforceSyncJobs.wsSalesforce;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace SalesforceSyncJobs
{
    class MarcusMillichap_Group
    {
        private string _groupName, _type, _developerName, _id, _adUserId;
        private bool? _doesIncludeBosses;
        private SforceService _service;
        private Dictionary<string, MarcusMillichap_Group> _salesforceGroups;
        private Dictionary<string, MarcusMillichap_GroupMember> _groupMembers;

        public Dictionary<string, MarcusMillichap_Group> SalesforceGroups { get { return _salesforceGroups; } set { _salesforceGroups = value; } }
        public Dictionary<string, MarcusMillichap_GroupMember> GroupMembers { get { return _groupMembers; } set { _groupMembers = value; } }

        public string ADUserId { get { return _adUserId; } set { _adUserId = value; } }
        public string GroupName { get { return _groupName; } set { _groupName = value; } }
        public string ID { get { return _id; } set { _id = value; } }
        public string Type { get { return _type; } set { _type = value; } }
        public string DeveloperName { get { return _developerName; } set { _developerName = value; } }
        public bool? DoesIncludeBosses { get { return _doesIncludeBosses; } set { _doesIncludeBosses = value; } }

        public MarcusMillichap_Group()
        {
            _developerName = string.Empty;
            _groupName = string.Empty;
            _type = "Regular";
            _doesIncludeBosses = false;
            _service = null;
            _salesforceGroups = new Dictionary<string,MarcusMillichap_Group>();
            _groupMembers = new Dictionary<string, MarcusMillichap_GroupMember>();
        }

        public MarcusMillichap_Group(SforceService service)
        {
            _developerName = string.Empty;
            _groupName = string.Empty;
            _type = "Regular";
            _doesIncludeBosses = false;
            _service = service;
            _salesforceGroups = new Dictionary<string, MarcusMillichap_Group>();
            _groupMembers = new Dictionary<string, MarcusMillichap_GroupMember>();
        }

        public void LoadSalesforceGroups()
        {
            try
            {

                QueryResult qResult = null;
                string query = string.Empty;
                string groupName = string.Empty;
                bool queryComplete = false;
                //query
                query += "SELECT id, DeveloperName, Name, Type, DoesIncludeBosses ";
                query += "FROM Group";
                //execute query
                qResult = _service.query(query);
                if (qResult.size > 0)
                {
                    while (!queryComplete)
                    {
                        sObject[] records = qResult.records;
                        //Groups g = new Groups();
                        for (int i = 0; i < records.Length; i++)
                        {
                            MarcusMillichap_Group g = new MarcusMillichap_Group();
                            wsSalesforce.Group sfG = (wsSalesforce.Group)records[i];
                            g.ID = sfG.Id;
                            g.DeveloperName = sfG.DeveloperName;
                            g.GroupName = sfG.Name;
                            g.Type = sfG.Type;
                            g.DoesIncludeBosses = sfG.DoesIncludeBosses;
                            if (!_salesforceGroups.ContainsKey(g.DeveloperName))
                            {
                                _salesforceGroups.Add(g.DeveloperName, g);
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

        public bool Create(Dictionary<string, MarcusMillichap_Group> groupList, out Dictionary<string, MarcusMillichap_Group> addedGroups)
        {
            MarcusMillichap_SF_API api = new MarcusMillichap_SF_API();
            bool success = false;
            addedGroups = new Dictionary<string, MarcusMillichap_Group>();
            List<wsSalesforce.Group> listObjects = new List<wsSalesforce.Group>();
            foreach (var g in groupList.Values)
            {
                wsSalesforce.Group newGroup = new wsSalesforce.Group();
                newGroup.DoesIncludeBosses = false;
                newGroup.Name = newGroup.DeveloperName = g.DeveloperName;
                newGroup.Type = "Regular";
                listObjects.Add(newGroup);
            }

            List<List<wsSalesforce.Group>> list = MarcusMillichap_Common.BreakIntoChunks(listObjects, 200);
            Dictionary<string, sObject> successObjects = new Dictionary<string, sObject>();
            List<MarcusMillichap_Errors> errors = new List<MarcusMillichap_Errors>();
            foreach (var subList in list)
            {
                Dictionary<string, sObject> temp = new Dictionary<string,sObject>();
                int index = 0;
                sObject[] sArray = new sObject[subList.Count];
                foreach (var item in subList)
                {
                    sArray[index] = item;
                    index++;
                }
                if (sArray != null && sArray.Length > 0)
                {
                    api.CreateGroups(sArray, ref _service, out temp, out errors);
                    addedGroups = addedGroups.Concat(ConvertSFObjectToGroupDictionary(temp)).ToDictionary(x => x.Key, x => x.Value);
                }

            }
            return success;
        }

        private Dictionary<string, MarcusMillichap_Group> ConvertSFObjectToGroupDictionary(Dictionary<string, sObject> d) {
            Dictionary<string, MarcusMillichap_Group> dict = new Dictionary<string, MarcusMillichap_Group>();
            if (d != null && d.Count > 0)
            {
                foreach (var a in d) {
                    wsSalesforce.Group sfG = (wsSalesforce.Group)a.Value;
                    if (!dict.ContainsKey(sfG.DeveloperName)) {
                        MarcusMillichap_Group g = new MarcusMillichap_Group();
                        g.ID = a.Key;
                        g.DeveloperName = sfG.DeveloperName;
                        g.GroupName = sfG.Name;
                        dict.Add(sfG.DeveloperName, g);
                    }
                }
            }
            return dict;
        }

        public bool Delete(List<string> groupIds)
        {
            MarcusMillichap_SF_API api = new MarcusMillichap_SF_API();
            bool results = true;
            List<MarcusMillichap_Errors> error = new List<MarcusMillichap_Errors>();
            List<string> success = new List<string>();
            List<List<string>> list = MarcusMillichap_Common.BreakIntoChunks(groupIds, int.Parse(ConfigurationManager.AppSettings["MAX_NUMBER_UPDATE"].ToString()));
            foreach (var subList in list)
            {
                int index = 0;
                string[] sArray = new string[subList.Count];
                foreach (var item in subList)
                {
                    sArray[index] = item;
                    index++;
                }
                if (sArray != null && sArray.Length > 0)
                {
                    api.Delete(sArray, ref _service, out success, out error);
                }
                if (error.Count > 0) results = false;
            }
            return results;
        }

        public void Compare(Dictionary<string, MarcusMillichap_Group> salesforceGroups, Dictionary<string, MarcusMillichap_Group> mnetGroups, out Dictionary<string, MarcusMillichap_Group> missingGroupsInSalesforce, out List<string> extraGroupsInSalesforce)
        {
            missingGroupsInSalesforce = new Dictionary<string, MarcusMillichap_Group>();
            extraGroupsInSalesforce = new List<string>();

            var missing = mnetGroups.Keys.Except(salesforceGroups.Keys);
            var extra = salesforceGroups.Keys.Except(mnetGroups.Keys);

            foreach (var m in missing)
            {
                if (!missingGroupsInSalesforce.ContainsKey(m))
                {
                    //Groups g = mnetGroups[m];
                    //g.GroupMembers = new Dictionary<string,SalesforceSync.GroupMembers>();
                    MarcusMillichap_Group g = new MarcusMillichap_Group();
                    g.DeveloperName = mnetGroups[m].DeveloperName;
                    g.GroupName = mnetGroups[m].DeveloperName;
                    missingGroupsInSalesforce.Add(m, g);
                }
            }

            foreach (var e in extra)
            {
                if(!e.ToString().ToLower().StartsWith("mnet_psg_") && !e.ToLower().Equals("allinternalusers"))
                    extraGroupsInSalesforce.Add(salesforceGroups[e].ID);
            }
        }


        public static List<string> FindGroupNameForUser(int userCode, Dictionary<string, MarcusMillichap_Group> mmGroupsMasterList) {
            List<string> teams = new List<string>();
            if (userCode > 0 && mmGroupsMasterList != null && mmGroupsMasterList.Count > 0)
            {
                foreach (var m in mmGroupsMasterList)
                {
                    if (m.Value.GroupMembers != null && m.Value.GroupMembers.Count > 0) {
                        foreach (var gm in m.Value.GroupMembers) {
                            if (gm.Value.UserCode == userCode)
                            {
                                teams.Add(m.Key);
                            }
                        }
                    }
                }
            }
            return teams;
        }
    }
}
