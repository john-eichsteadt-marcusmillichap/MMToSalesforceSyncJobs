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
    class MarcusMillichap_User
    {
        private List<MarcusMillichap_User> _failedUsers;
        private SforceService _service;
        private Dictionary<string, MarcusMillichap_User> _saleforceUsers;
        private Dictionary<int, MarcusMillichap_User> _salesforceUsersFromMnetComparison;
        private Dictionary<int, MarcusMillichap_User> _mmUsers;
        private Dictionary<string, MarcusMillichap_Group> _mmGroups;
        private int _id, _userCode;
        int? _userCodeManager;
        private string _salesforceId, _firstName, _lastName, _title, _department, _company, _alias, _userName, _emailAddress, _communityNickName, _groupName;
        private string _profileId, _emailEncoding, _federationId, _timeZone, _locale, _language, _adUserId, _phone, _fax, _mobile, _flag, _officeName, _officeCity;
        private bool? _active, _isMMCC;
        private DateTime? _lastModified;

        public int ID { get { return _id; } set { _id = value; } }
        public int UserCode { get { return _userCode; } set { _userCode = value; } }
        public string SalesforceId { get { return _salesforceId; } set { _salesforceId = value; } }
        public DateTime? LastModified { get { return _lastModified; } set { _lastModified = value; } }

        public int? UserCodeManager { get { return _userCodeManager; } set { _userCodeManager = value; } }
        public string ADUserId { get { return _adUserId; } set { _adUserId = value; } }
        public string FirstName { get { return _firstName; } set { _firstName = value; } }
        public string LastName { get { return _lastName; } set { _lastName = value; } }
        public string Title { get { return _title; } set { _title = value; } }
        public string Department { get { return _department; } set { _department = value; } }
        public string Company { get { return _company; } set { _company = value; } }
        public string Alias { get { return _alias; } set { _alias = value; } }
        public string UserName { get { return _userName; } set { _userName = value; } }
        public string EmailAddress { get { return _emailAddress; } set { _emailAddress = value; } }
        public string CommunityNickName { get { return _communityNickName; } set { _communityNickName = value; } }
        public string ProfileId { get { return _profileId; } set { _profileId = value; } }
        public string EmailEncoding { get { return _emailEncoding; } set { _emailEncoding = value; } }
        public string FederationId { get { return _federationId; } set { _federationId = value; } }
        public string TimeZone { get { return _timeZone; } set { _timeZone = value; } }
        public string Locale { get { return _locale; } set { _locale = value; } }
        public string Language { get { return _language; } set { _language = value; } }
        public string Phone { get { return _phone; } set { _phone = value; } }
        public string Fax { get { return _fax; } set { _fax = value; } }
        public string Mobile { get { return _mobile; } set { _mobile = value; } }
        public string GroupName { get { return _groupName; } set { _groupName = value; } }
        public string Flag { get { return _flag; } set { _flag = value; } }
        public string OfficeName { get { return _officeName; } set { _officeName = value; } }
        public string OfficeCity { get { return _officeCity; } set { _officeCity = value; } }
        public bool? Active { get { return _active; } set { _active = value; } }
        public bool? IsMMCC { get { return _isMMCC; } set { _isMMCC = value; } }

        public Dictionary<int, MarcusMillichap_User> MMUsers { get { return _mmUsers; } set { _mmUsers = value; } }
        public Dictionary<int, MarcusMillichap_User> SalesforceUsersFromMnetComparison { get { return _salesforceUsersFromMnetComparison; } set { _salesforceUsersFromMnetComparison = value; } }
        public Dictionary<string, MarcusMillichap_User> SaleforceUsers { get { return _saleforceUsers; } set { _saleforceUsers = value; } }
        public Dictionary<string, MarcusMillichap_Group> MMGroups { get { return _mmGroups; } set { _mmGroups = value; } }

        public MarcusMillichap_User() {
            _id = 0;
            _salesforceId = string.Empty;
            _userCode = 0;
            _lastModified = null;
            _failedUsers = new List<MarcusMillichap_User>();
            _saleforceUsers = new Dictionary<string, MarcusMillichap_User>();
            _salesforceUsersFromMnetComparison = new Dictionary<int, MarcusMillichap_User>();
            _mmGroups = new Dictionary<string, MarcusMillichap_Group>();
            _mmUsers = new Dictionary<int, MarcusMillichap_User>();
            _firstName = string.Empty;
            _lastName = string.Empty;
            _title = string.Empty;
            _department = string.Empty;
            _company = string.Empty;
            _alias = string.Empty;
            _userName = string.Empty;
            _emailAddress = string.Empty;
            _communityNickName = string.Empty;
            _profileId = ConfigurationManager.AppSettings["MMStandardUserProfileID"].ToString();
            _emailEncoding = "ISO-8859-1";
            _federationId = string.Empty;
            _timeZone = "America/Los_Angeles";
            _locale = "en_US";
            _language = "en_US";
            _userCode = 0;
            _flag = string.Empty;
            _officeName = string.Empty;
            _officeCity = string.Empty;
            _isMMCC = false;
        }

        public MarcusMillichap_User(SforceService service)
        {
            _service = service;
            _id = 0;
            _salesforceId = string.Empty;
            _userCode = 0;
            _lastModified =null;
            _failedUsers = new List<MarcusMillichap_User>();
            _saleforceUsers = new Dictionary<string, MarcusMillichap_User>();
            _salesforceUsersFromMnetComparison = new Dictionary<int, MarcusMillichap_User>();
            _mmGroups = new Dictionary<string, MarcusMillichap_Group>();
            _mmUsers = new Dictionary<int, MarcusMillichap_User>();
            _firstName = string.Empty;
            _lastName = string.Empty;
            _title = string.Empty;
            _department = string.Empty;
            _company = string.Empty;
            _alias = string.Empty;
            _userName = string.Empty;
            _emailAddress = string.Empty;
            _communityNickName = string.Empty;
            _profileId = ConfigurationManager.AppSettings["MMStandardUserProfileID"].ToString();
            _emailEncoding = "ISO-8859-1";
            _federationId = string.Empty;
            _timeZone = "America/Los_Angeles";
            _locale = "en_US";
            _language = "en_US";
            _userCode = 0;
            _flag = string.Empty;
            _officeName = string.Empty;
            _officeCity = string.Empty;
            _isMMCC = false;
        }

        public Results AddUpdateMarcusMillichap()
        {
            Results r = new Results();
            if (!string.IsNullOrEmpty(_salesforceId) && _userCode > 0)
            {
                DataProvider dp = new DataProvider(ConfigurationManager.ConnectionStrings["MMPeople"].ToString());
                NameValueCollection iParam = new NameValueCollection();
                NameValueCollection oParam = new NameValueCollection();
                NameValueCollection output = new NameValueCollection();
                DataSet ds = new DataSet();
                try
                {
                    //input parameters
                    iParam.Add("@SFOwnerID", _salesforceId);
                    iParam.Add("@UserCode", _userCode.ToString());
                    iParam.Add("@SFLastModifiedDate", _lastModified.ToString());
                    //output parameters
                    oParam.Add("@OwnerID", "4");
                    oParam.Add("@ErrCode", "4");
                    oParam.Add("@ErrMessage", "2048");
                    //run query 
                    output = dp.DoNonQueryGetMultipleOutputParam("[MMPeopleApp].[spSFOwner_InsertUpdate]", CommandType.StoredProcedure, iParam, oParam);
                    if (output != null)
                    {
                        int iTemp = 0;
                        int.TryParse(output["@OwnerID"].ToString(), out iTemp);
                        r.ID = iTemp;
                        int.TryParse(output["@ErrCode"].ToString(), out iTemp);
                        r.ErrorCode = iTemp;
                        r.ErrorMessage = output["@ErrMessage"].ToString();
                        if (r.ID > 0 && string.IsNullOrEmpty(r.ErrorMessage)) r.Success = true;
                    }

                }
                catch (Exception ex) { throw ex; }
                finally { dp.Dispose(); iParam = null; oParam = null; output = null; }
            }
            return r;
        }

        public bool SyncToMarcusMillichap(DateTime lastRun)
        {
            bool success = true;
            try
            {
                //initialize sf queryresult object
                QueryResult qResult = null;
                string query = string.Empty;
                string sfDate = string.Empty;
                bool queryComplete = false;

                //convert to sf date
                sfDate = MarcusMillichap_Common.ConvertToSalesforceDateTime(lastRun);
                if (!string.IsNullOrEmpty(sfDate))
                {
                    query += "SELECT Id, MNetUserCode__c, LastModifiedDate";
                    query += " FROM User";
                    query += " WHERE CreatedDate >= " + sfDate;
                    query += " OR LastModifiedDate >= " + sfDate;
                    qResult = _service.query(query);

                    if (qResult.size > 0)
                    {
                        while (!queryComplete)
                        {
                            sObject[] records = qResult.records;
                            for (int i = 0; i < records.Length; i++)
                            {
                                wsSalesforce.User user = (wsSalesforce.User)records[i];
                                //sync data here to mm database
                                if (!string.IsNullOrEmpty(user.MNetUserCode__c))
                                {
                                    MarcusMillichap_User u = new MarcusMillichap_User
                                    {
                                        SalesforceId = user.Id,
                                        UserCode = int.Parse(user.MNetUserCode__c),
                                        LastModified = user.LastModifiedDate
                                    };
                                    Results r = u.AddUpdateMarcusMillichap();
                                    if (!r.Success)
                                    {
                                        _failedUsers.Add(u);
                                    }
                                }
                            }
                            if (qResult.done) queryComplete = true;
                            else qResult = _service.queryMore(qResult.queryLocator);
                        }
                    }
                }

                //retry the ones that failed
                if (_failedUsers != null && _failedUsers.Count > 0)
                {
                    Results rFailed;
                    foreach (MarcusMillichap_User failed in _failedUsers)
                    {
                        rFailed = failed.AddUpdateMarcusMillichap();
                        if (!rFailed.Success)
                        {
                            success = false;
                            //do something here
                        }
                    }
                }

            }
            catch (Exception ex) {
                MarcusMillichap_Email e = new MarcusMillichap_Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                                   ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                                  "SalesForce API Error - Users.cs - SyncToMarcusMillichap()",
                                   "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                                   "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                                   "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                                   true);
                e.SendEmail();
                success = false;
            }
            return success;
        }

        /// <summary>
        /// Get all users in Salesforce
        /// </summary>
        /// <returns></returns>
        public void LoadAllSalesforceUsers()
        {
            try
            {
                Console.WriteLine("\nGetting users from salesforce ...\n");
                QueryResult qResult = null;
                string query = string.Empty;
                bool queryComplete = false;
                //query
                query += "SELECT id, FirstName, LastName, Title, Department, CompanyName, UserName, Alias, MNetUserCode__c, Email, CommunityNickName, ProfileId, EmailEncodingKey, FederationIdentifier, TimeZoneSidKey, LocaleSidKey, LanguageLocaleKey, IsActive, MNet__c, Type_of_License__c, Phone, MobilePhone, Fax, Office__c, Office_City__c, Manager.MNetUserCode__c ";
                query += "FROM User ";
                query += "WHERE IsActive = true";
                qResult = _service.query(query);
                if (qResult.size > 0)
                {
                    while (!queryComplete)
                    {
                        sObject[] records = qResult.records;
                        int temp = 0;
                        for (int i = 0; i < records.Length; i++)
                        {
                            wsSalesforce.User u = (wsSalesforce.User)records[i];
                            MarcusMillichap_User sfUsers = new MarcusMillichap_User();
                            sfUsers.SalesforceId = u.Id;
                            sfUsers.FirstName = u.FirstName;
                            sfUsers.LastName = u.LastName;
                            sfUsers.Title = u.Title;
                            sfUsers.Department = u.Department;
                            sfUsers.Company = u.CompanyName;
                            sfUsers.Alias = u.Alias;
                            sfUsers.EmailAddress = u.Email;
                            sfUsers.CommunityNickName = u.CommunityNickname;
                            sfUsers.ProfileId = u.ProfileId;
                            sfUsers.FederationId = u.FederationIdentifier;
                            sfUsers.TimeZone = u.TimeZoneSidKey;
                            sfUsers.Locale = u.LocaleSidKey;
                            sfUsers.Language = u.LanguageLocaleKey;
                            sfUsers.UserName = u.Username;
                            sfUsers.EmailEncoding = u.EmailEncodingKey;
                            sfUsers.Active = u.IsActive;
                            sfUsers.ADUserId = u.MNet__c;
                            sfUsers.Phone = u.Phone.CleanPhoneNumber();
                            sfUsers.Mobile = u.MobilePhone.CleanPhoneNumber();
                            sfUsers.Fax = u.Fax.CleanPhoneNumber();
                            sfUsers.Flag = u.Type_of_License__c;
                            sfUsers.OfficeName = u.Office__c;
                            sfUsers.OfficeCity = u.Office_City__c;
                            sfUsers.UserCodeManager = u.Manager != null ? u.Manager.MNetUserCode__c.ToNullableInt() : null;

                            if (!String.IsNullOrEmpty(u.MNetUserCode__c)) sfUsers.UserCode = int.TryParse(u.MNetUserCode__c, out temp) ? temp : 0;

                            if (!_saleforceUsers.ContainsKey(u.Id))
                                _saleforceUsers.Add(u.Id, sfUsers);

                            if (!_salesforceUsersFromMnetComparison.ContainsKey(sfUsers.UserCode))
                                _salesforceUsersFromMnetComparison.Add(sfUsers.UserCode, sfUsers);
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

        public void LoadAllMarcusMillichapUsersAndGroups()
        {
            Console.WriteLine("\nGetting users from marcus millichap ...\n");
            DataProvider dp = new DataProvider(ConfigurationManager.ConnectionStrings["MMPeople"].ToString());
            DataSet ds = new DataSet();
            bool bTemp;
            try
            {
                ds = dp.GetDataSet("[MMPeopleApp].[spSFGroupsAndUsers]", CommandType.StoredProcedure, null);
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        //add group to list
                        if (!_mmGroups.ContainsKey(dr["GroupName"].ToString()))
                        {
                            _mmGroups.Add(dr["GroupName"].ToString(), new MarcusMillichap_Group { DeveloperName = dr["GroupName"].ToString(), GroupName = dr["GroupName"].ToString() });
                        }
                        //add user to list
                        int userCode = 0, temp = 0;
                        int? userCodeManager = null;
                        if (!String.IsNullOrEmpty(dr["UserCode"].ToString()) && int.TryParse(dr["UserCode"].ToString(), out userCode) && userCode > 0)
                        {

                            if (!String.IsNullOrEmpty(dr["UserCode_Manager_Primary"].ToString()))
                            {
                                int.TryParse(dr["UserCode_Manager_Primary"].ToString(), out temp);
                                if (temp > 0) userCodeManager = temp;
                            }
                            string userName = dr["ADUserID"].ToString() + "@marcusmillichap.com";
                            if (dr["ADUserID"].ToString().ToLower().Equals("jodonnell "))
                            {
                                userName = "jo'donnell@marcusmillichap.com";
                            }

                            var firstName = dr["DisplayName"].ToString().Trim().Replace(dr["LastName"].ToString().Trim(), "").Trim();
                            if (string.IsNullOrEmpty(firstName)){
                                firstName = dr["FirstName"].ToString();
                            }

                            MarcusMillichap_User u = new MarcusMillichap_User
                            {
                                UserCode = userCode,
                                ADUserId = dr["ADUserID"].ToString(),
                                FirstName = firstName,
                                LastName = dr["LastName"].ToString(),
                                Phone = dr["PhoneNumber_OfficeDirect"].ToString(),
                                Mobile = dr["PhoneNumber_Mobile"].ToString(),
                                Fax = dr["PhoneNumber_Fax"].ToString(),
                                Company = dr["Company_ShortName"].ToString(),
                                EmailAddress = dr["EmailAddress"].ToString(),
                                Title = dr["Title"].ToString(),
                                UserCodeManager = userCodeManager,
                                GroupName = dr["GroupName"].ToString(),
                                Alias = dr["ADUserID"].ToString(),
                                UserName = userName,
                                FederationId = dr["UserPrincipalName"].ToString(), 
                                CommunityNickName = dr["ADUserID"].ToString() + "_" + userCode.ToSafeString(),
                                Flag = dr["Flag"].ToString(),
                                OfficeName = dr["OfficeName"].ToString(),
                                OfficeCity = dr["Office_City"].ToString(),
                                TimeZone = GetTimeZone(dr["GMTOffSet"].ToNullableInt(), bool.TryParse(dr["DaylightSavings"].ToString(), out bTemp) ? bTemp : false),
                                IsMMCC = (dr["Company_ShortName"].ToString().ToLower().Equals("mmcc")) ? true : false,
                                ProfileId = (!string.IsNullOrEmpty(dr["SalesForceProfile_SalesForceID"].ToString())) ? dr["SalesForceProfile_SalesForceID"].ToString() : ConfigurationManager.AppSettings["MMStandardUserProfileID"].ToString()
                            };
                            if (!_mmUsers.ContainsKey(userCode))
                            {
                                _mmUsers.Add(userCode, u);
                            }

                            MarcusMillichap_GroupMember gm = new MarcusMillichap_GroupMember
                            {
                                ADUserId = dr["ADUserID"].ToString().ToLower(),
                                DeveloperName = dr["GroupName"].ToString(),
                                GroupName = dr["GroupName"].ToString(),
                                UserCode = userCode
                            };

                            //add members to groups
                            if (_mmGroups.ContainsKey(dr["GroupName"].ToString()))
                            {
                                if (!_mmGroups[dr["GroupName"].ToString()].GroupMembers.ContainsKey(dr["ADUserID"].ToString().ToLower()))
                                {
                                    _mmGroups[dr["GroupName"].ToString()].GroupMembers.Add(dr["ADUserID"].ToString().ToLower(), gm);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { throw ex; }
            finally { dp.Dispose(); }
        }

        public void SyncToSalesforce(Dictionary<int, MarcusMillichap_User> salesforceUsers, Dictionary<int, MarcusMillichap_User> mmUsers, out Dictionary<string, wsSalesforce.User> newlyAddedUsers, out List<MarcusMillichap_Errors> errors)
        {
            Console.WriteLine("\nSyncing users ...\n");
            List<wsSalesforce.User> sfUserObjects = new List<wsSalesforce.User>();
            errors = new List<MarcusMillichap_Errors>();
            newlyAddedUsers = new Dictionary<string, wsSalesforce.User>();
            try
            {
                foreach (int key in mmUsers.Keys)
                {
                    wsSalesforce.User sfUserInsert = new wsSalesforce.User();
                    bool addToList = true;
                    //update user info if changes
                    if (salesforceUsers.ContainsKey(key))
                    {
                        if (MarcusMillichap_Common.IsDirty(salesforceUsers, mmUsers, key))
                        {
                            sfUserInsert.ProfileId = salesforceUsers[key].ProfileId;
                            addToList = true;
                        }
                        else addToList = false;
                    }

                    if (addToList)
                    {
                        ArrayList nullFields = new ArrayList();
                        MarcusMillichap_User insertUser = mmUsers[key];
                        sfUserInsert.FirstName = insertUser.FirstName;
                        sfUserInsert.LastName = insertUser.LastName;
                        // set to 80 since there is a CAP in salesforce
                        sfUserInsert.Title = insertUser.Title.ToMaxLength(80);
                        sfUserInsert.CompanyName = insertUser.Company;
                        // set to 8 since there is a CAP in salesforce
                        sfUserInsert.Alias = insertUser.Alias.ToMaxLength(8);
                        sfUserInsert.Username = insertUser.UserName;
                        sfUserInsert.Email = insertUser.EmailAddress;
                        sfUserInsert.CommunityNickname = insertUser.CommunityNickName;
                        sfUserInsert.FederationIdentifier = insertUser.FederationId;
                        sfUserInsert.TimeZoneSidKey = insertUser.TimeZone;
                        sfUserInsert.LocaleSidKey = insertUser.Locale;
                        sfUserInsert.LanguageLocaleKey = insertUser.Language;
                        sfUserInsert.MNetUserCode__c = insertUser.UserCode.ToString();
                        sfUserInsert.EmailEncodingKey = insertUser.EmailEncoding;
                        sfUserInsert.ProfileId = insertUser.ProfileId;
                        if (!string.IsNullOrEmpty(insertUser.Phone))
                        {
                            string extension = string.Empty;
                            sfUserInsert.Phone = MarcusMillichap_Common.FormatPhoneNumber(insertUser.Phone, out extension);
                            sfUserInsert.Extension = extension;
                        }
                        else
                            nullFields.Add("Phone");
                        if (!string.IsNullOrEmpty(insertUser.Fax))
                            sfUserInsert.Fax = MarcusMillichap_Common.FormatPhoneNumber(insertUser.Fax);
                        else
                            nullFields.Add("Fax");
                        if (!string.IsNullOrEmpty(insertUser.Mobile))
                            sfUserInsert.MobilePhone = MarcusMillichap_Common.FormatPhoneNumber(insertUser.Mobile);
                        else
                            nullFields.Add("MobilePhone");

                        sfUserInsert.MNet__c = insertUser.ADUserId;

                        sfUserInsert.Office__c = insertUser.OfficeName;
                        sfUserInsert.Office_City__c = insertUser.OfficeCity;

                        if (insertUser.UserCodeManager.HasValue && insertUser.UserCodeManager > 0)
                        {
                            if (salesforceUsers.ContainsKey(insertUser.UserCodeManager.Value))
                            {
                                string getSalesForceUserId = salesforceUsers[insertUser.UserCodeManager.Value].SalesforceId;
                                if (!String.IsNullOrEmpty(getSalesForceUserId))
                                {
                                    sfUserInsert.Manager = new wsSalesforce.User { MNetUserCode__c = insertUser.UserCodeManager.Value.ToSafeString() };
                                }
                            }
                        }

                        sfUserInsert.IsActive = true;
                        sfUserInsert.IsActiveSpecified = true;

                        if (nullFields.Count > 0)
                        {
                            string[] s = new string[nullFields.Count];
                            for (int j = 0; j < nullFields.Count; j++)
                            {
                                s[j] = nullFields[j].ToString();
                            }
                            sfUserInsert.fieldsToNull = s;
                        }

                        sfUserInsert.Type_of_License__c = insertUser.Flag;

                        sfUserObjects.Add(sfUserInsert);
                    }
                }

                 MarcusMillichap_SF_API api = new MarcusMillichap_SF_API();
                //loop through list to update in chucks of 200
                List<List<wsSalesforce.User>> list = MarcusMillichap_Common.BreakIntoChunks(sfUserObjects, int.Parse(ConfigurationManager.AppSettings["MAX_NUMBER_UPDATE"].ToString()));
                foreach (var subList in list)
                {
                    int index = 0;
                    sObject[] sArray = new sObject[subList.Count];
                    Dictionary<string, wsSalesforce.User> temp = new Dictionary<string, wsSalesforce.User>();
                    foreach (var item in subList)
                    {
                        sArray[index] = item;
                        index++;
                    }
                    if (sArray != null && sArray.Length > 0)
                    {
                        api.UpsertUser("MNetUserCode__c", sArray, ref _service, out errors);
                        newlyAddedUsers = newlyAddedUsers.Concat(temp).ToDictionary(x => x.Key, x => x.Value);
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Upsert User
        /// </summary>
        /// <param name="usersToUpsert"></param>
        /// <param name="listOfAllCurrentUsersInSalesforce"></param>
        /// <param name="newlyAddedUsers"></param>
        /// <param name="errors"></param>
        public void Deactivate(Dictionary<int, MarcusMillichap_User> usersToDelete, Dictionary<string, MarcusMillichap_User> listOfAllCurrentUsersInSalesforce, out List<MarcusMillichap_Errors> errors)
        {
            List<wsSalesforce.User> sfUserObjects = new List<wsSalesforce.User>();
            MarcusMillichap_SF_API api = new MarcusMillichap_SF_API();
            errors = new List<MarcusMillichap_Errors>();
            Dictionary<string, sObject> success = new Dictionary<string, sObject>();
            try
            {
                foreach (int i in usersToDelete.Keys)
                {
                    if (i > 0 && usersToDelete[i].UserCode > 0)
                    {
                        wsSalesforce.User sfUser = new wsSalesforce.User();
                        string salesforceId = MarcusMillichap_Common.FindUserByUserCode(i, listOfAllCurrentUsersInSalesforce);
                        if (listOfAllCurrentUsersInSalesforce.ContainsKey(salesforceId))
                        {
                            MarcusMillichap_User u = listOfAllCurrentUsersInSalesforce[salesforceId];
                            sfUser.Id = salesforceId;
                            sfUser.IsActive = false;
                            sfUser.IsActiveSpecified = true;
                            if (u.UserCode.ToSafeString() != "")
                                sfUserObjects.Add(sfUser);
                        }
                    }
                }
                //loop through list to update in chucks of 200
                List<List<wsSalesforce.User>> list = MarcusMillichap_Common.BreakIntoChunks(sfUserObjects, int.Parse(ConfigurationManager.AppSettings["MAX_NUMBER_UPDATE"].ToString()));
                foreach (var subList in list)
                {
                    int index = 0;
                    sObject[] sArray = new sObject[subList.Count];
                    foreach (var item in subList)
                    {
                        sArray[index] = item;
                        index++;
                    }
                    if (sArray != null && sArray.Length > 0)
                    {
                        api.Update(sArray, ref _service, out success, out errors);
                    }
                }
            }
            catch
            {

            }
        }


        /// <summary>
        /// Compare Users in salesforce and mnet
        /// </summary>
        /// <param name="salesforceUsers"></param>
        /// <param name="mnetUsers"></param>
        /// <param name="missingUsersInSalesforce"></param>
        /// <param name="extraUsersInSalesforce"></param>
        /// Dictionary<int, Users> salesforceUsers - int -> usercode
        /// Dictionary<int, Users> mnetUsers - int -> usercode
        public void Compare(Dictionary<int, MarcusMillichap_User> salesforceUsers, Dictionary<int, MarcusMillichap_User> mnetUsers, out Dictionary<int, MarcusMillichap_User> missingUsersInSalesforce, out Dictionary<int, MarcusMillichap_User> extraUsersInSalesforce)
        {
            missingUsersInSalesforce = new Dictionary<int, MarcusMillichap_User>();
            extraUsersInSalesforce = new Dictionary<int, MarcusMillichap_User>();

            var missing = mnetUsers.Keys.Except(salesforceUsers.Keys);
            var extra = salesforceUsers.Keys.Except(mnetUsers.Keys);

            foreach (var m in missing)
            {
                if (!missingUsersInSalesforce.ContainsKey(m))
                {
                    missingUsersInSalesforce.Add(m, mnetUsers[m]);
                }
            }

            foreach (var e in extra)
            {
                if (!extraUsersInSalesforce.ContainsKey(e))
                {
                    extraUsersInSalesforce.Add(e, salesforceUsers[e]);
                }
            }
        }


        private string GetTimeZone(int? gmtOffSet, bool isDaylightSavings)
        {
            if (gmtOffSet.HasValue)
            {
                if (isDaylightSavings) gmtOffSet = gmtOffSet + 1;
                switch (gmtOffSet.Value)
                {
                    case -3:
                        if (isDaylightSavings) return "Atlantic/Bermuda";
                        else return "Atlantic/Bermuda";
                    case -4:
                        if (isDaylightSavings) return "America/New_York";
                        else return "America/Indianapolis";
                    case -5:
                        if (isDaylightSavings) return "America/Chicago";
                        else return "America/Panama";
                    case -6:
                        if (isDaylightSavings) return "America/Denver";
                        else return "America/Mexico_City";
                    case -7:
                        if (isDaylightSavings) return "America/Los_Angeles";
                        else return "America/Phoenix";
                    case -8:
                        if (isDaylightSavings) return "America/Anchorage";
                        else return "Pacific/Pitcairn";
                    case -9:
                        if (isDaylightSavings) return "America/Atka";
                        else return "America/Atka";
                    case -10:
                        if (isDaylightSavings) return "Pacific/Honolulu";
                        else return "Pacific/Honolulu";
                    case -11:
                        if (isDaylightSavings) return "Pacific/Pago_Pago";
                        else return "Pacific/Pago_Pago";
                    default:
                        return "America/Los_Angeles";
                }
            }
            else
            {
                return "America/Los_Angeles";
            }
        }
    }
}
