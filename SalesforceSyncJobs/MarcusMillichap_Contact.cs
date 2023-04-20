using SalesforceSyncJobs.wsSalesforce;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace SalesforceSyncJobs
{
    class MarcusMillichap_Contact
    {
        private SforceService _service;
        private List<MarcusMillichap_Contact> _failedContacts;
        private int _id;
        private string _sfId, _sfAccountId, _sfOwnerId, _firstName, _lastName, _businessPhone, _directPhone, _fax, _email, _loginEmail, _otherEmail, _productType, _statePreference, _metros, _title;
        private string _sfAssignedToId, _mobilePhone, _homePhone, _assistantDirectPhone, _otherPhone;
        private bool _hasOptedOutOfEmail, _emailContactDNC, _faxDnc, _businessPhoneDnc, _directPhoneDnc, _mobilePhoneDnc, _assistantDirectPhoneDnc, _otherPhoneDNC, _homePhoneDnc, _status, _ipa, _researchReports, _doNotEmail, _isCompanyRecord;
        private DateTime? _lastModified;

        public int ID { get { return _id; } set { _id = value; } }

        public string SFAssignedToID { get { return _sfAssignedToId; } set { _sfAssignedToId = value; } }
        public string SFId { get { return _sfId; } set { _sfId = value; } }
        public string SFAccountId { get { return _sfAccountId; } set { _sfAccountId = value; } }
        public string SFOwnerId { get { return _sfOwnerId; } set { _sfOwnerId = value; } }
        public string FirstName { get { return _firstName; } set { _firstName = value; } }
        public string LastName { get { return _lastName; } set { _lastName = value; } }
        public string BusinessPhone { get { return _businessPhone; } set { _businessPhone = value; } }
        public string DirectPhone { get { return _directPhone; } set { _directPhone = value; } }
        public string Fax { get { return _fax; } set { _fax = value; } }
        public string Email { get { return _email; } set { _email = value; } }
        public string LoginEmail { get { return _loginEmail; } set { _loginEmail = value; } }
        public string OtherEmail { get { return _otherEmail; } set { _otherEmail = value; } }
        public string ProductType { get { return _productType; } set { _productType = value; } }
        public string StatePreference { get { return _statePreference; } set { _statePreference = value; } }
        public string Metros { get { return _metros; } set { _metros = value; } }
        public string MobilePhone { get { return _mobilePhone; } set { _mobilePhone = value; } }
        public string HomePhone { get { return _homePhone; } set { _homePhone = value; } }
        public string AssistantDirectPhone { get { return _assistantDirectPhone; } set { _assistantDirectPhone = value; } }
        public string OtherPhone { get { return _otherPhone; } set { _otherPhone = value; } }
        public string Title { get { return _title; } set { _title = value; } }

        public bool IsCompanyRecord { get { return _isCompanyRecord; } set { _isCompanyRecord = value; } }
        public bool BusinessPhoneDnc { get { return _businessPhoneDnc; } set { _businessPhoneDnc = value; } }
        public bool DirectPhoneDnc { get { return _directPhoneDnc; } set { _directPhoneDnc = value; } }
        public bool MobilePhoneDnc { get { return _mobilePhoneDnc; } set { _mobilePhoneDnc = value; } }
        public bool AssistantDirectPhoneDnc { get { return _assistantDirectPhoneDnc; } set { _assistantDirectPhoneDnc = value; } }
        public bool OtherPhoneDNC { get { return _otherPhoneDNC; } set { _otherPhoneDNC = value; } }
        public bool HomePhoneDnc { get { return _homePhoneDnc; } set { _homePhoneDnc = value; } }
        public bool FaxDnc { get { return _faxDnc; } set { _faxDnc = value; } }
        public bool Status { get { return _status; } set { _status = value; } }
        public bool IPA { get { return _ipa; } set { _ipa = value; } }
        public bool ResearchReports { get { return _researchReports; } set { _researchReports = value; } }
        public bool DoNotEmail { get { return _doNotEmail; } set { _doNotEmail = value; } }
        public bool HasOptedOutOfEmail { get { return _hasOptedOutOfEmail; } set { _hasOptedOutOfEmail = value; } }
        public bool EmailContactDNC { get { return _emailContactDNC; } set { _emailContactDNC = value; } }

        public DateTime? LastModified { get { return _lastModified; } set { _lastModified = value; } }

        public MarcusMillichap_Contact()
        {
            _sfId = string.Empty;
            _sfAccountId = string.Empty;
            _sfOwnerId = string.Empty;
            _firstName = string.Empty;
            _lastName = string.Empty;
            _businessPhone = string.Empty;
            _directPhone = string.Empty;
            _fax = string.Empty;
            _email = string.Empty;
            _loginEmail = string.Empty;
            _otherEmail = string.Empty;
            _productType = string.Empty;
            _statePreference = string.Empty;
            _metros = string.Empty;
            _mobilePhone = string.Empty;
            _homePhone = string.Empty;
            _assistantDirectPhone = string.Empty;
            _otherPhone = string.Empty;
            _title = string.Empty;
            _sfAssignedToId = string.Empty;
            _businessPhoneDnc = false;
            _directPhoneDnc = false;
            _mobilePhoneDnc = false;
            _assistantDirectPhoneDnc = false;
            _otherPhoneDNC = false;
            _homePhoneDnc = false;
            _faxDnc = false;
            _ipa = false;
            _status = true;
            _isCompanyRecord = false;
            _researchReports = true;
            _doNotEmail = true;
            _hasOptedOutOfEmail = false;
            _emailContactDNC = false;
            _lastModified = null;
            _failedContacts = new List<MarcusMillichap_Contact>();
        }

        public MarcusMillichap_Contact(SforceService service)
        {
            _service = service;
            _sfId = string.Empty;
            _sfAccountId = string.Empty;
            _sfOwnerId = string.Empty;
            _firstName = string.Empty;
            _lastName = string.Empty;
            _businessPhone = string.Empty;
            _directPhone = string.Empty;
            _fax = string.Empty;
            _email = string.Empty;
            _loginEmail = string.Empty;
            _otherEmail = string.Empty;
            _productType = string.Empty;
            _statePreference = string.Empty;
            _metros = string.Empty;
            _mobilePhone = string.Empty;
            _homePhone = string.Empty;
            _assistantDirectPhone = string.Empty;
            _otherPhone = string.Empty;
            _title = string.Empty;
            _sfAssignedToId = string.Empty;
            _isCompanyRecord = false;
            _businessPhoneDnc = false;
            _directPhoneDnc = false;
            _mobilePhoneDnc = false;
            _assistantDirectPhoneDnc = false;
            _otherPhoneDNC = false;
            _homePhoneDnc = false;
            _faxDnc = false;
            _ipa = false;
            _status = true;
            _researchReports = true;
            _doNotEmail = true;
            _hasOptedOutOfEmail = false;
            _emailContactDNC = false;
            _lastModified = null;
            _failedContacts = new List<MarcusMillichap_Contact>();
        }

        public Results AddUpdateMarcusMillichap()
        {
            Results r = new Results();
            if (!string.IsNullOrEmpty(_sfId) && !string.IsNullOrEmpty(_sfAccountId) && !string.IsNullOrEmpty(_sfOwnerId))
            {
                DataProvider dp = new DataProvider(ConfigurationManager.ConnectionStrings["MMPeople"].ToString());
                NameValueCollection iParam = new NameValueCollection();
                NameValueCollection oParam = new NameValueCollection();
                NameValueCollection output = new NameValueCollection();
                DataSet ds = new DataSet();
                try
                {
                    //input parameters
                    iParam.Add("@SFContactID", _sfId);
                    iParam.Add("@SFAccountID", _sfAccountId);
                    iParam.Add("@SFOwnerID", _sfOwnerId);
                    if (!string.IsNullOrEmpty(_firstName)) iParam.Add("@FirstName", _firstName);
                    if (!string.IsNullOrEmpty(_lastName)) iParam.Add("@LastName", _lastName);
                    if (!string.IsNullOrEmpty(_businessPhone)) iParam.Add("@BusinessPhone", _businessPhone);
                    iParam.Add("@BusinessPhoneDNC", _businessPhoneDnc.ToString());
                    if (!string.IsNullOrEmpty(_directPhone)) iParam.Add("@DirectPhone", _directPhone);
                    iParam.Add("@DirectPhoneDNC", _directPhoneDnc.ToString());
                    if (!string.IsNullOrEmpty(_mobilePhone)) iParam.Add("@MobilePhone", _mobilePhone);
                    iParam.Add("@MobilePhoneDNC", _mobilePhoneDnc.ToString());
                    if (!string.IsNullOrEmpty(_homePhone)) iParam.Add("@HomePhone", _homePhone);
                    iParam.Add("@HomePhoneDNC", _homePhoneDnc.ToString());
                    if (!string.IsNullOrEmpty(_assistantDirectPhone)) iParam.Add("@AssistantDirectPhone", _assistantDirectPhone);
                    iParam.Add("@AssistantDirectPhoneDNC", _assistantDirectPhoneDnc.ToString());
                    if (!string.IsNullOrEmpty(_otherPhone)) iParam.Add("@OtherPhone", _otherPhone);
                    iParam.Add("@OtherPhoneDNC", _otherPhoneDNC.ToString());
                    if (!string.IsNullOrEmpty(_fax)) iParam.Add("@Fax", _fax);
                    iParam.Add("@FaxDNC", _faxDnc.ToString());
                    if (!string.IsNullOrEmpty(_email)) iParam.Add("@Email", _email);
                    if (!string.IsNullOrEmpty(_loginEmail)) iParam.Add("@LoginEmail", _loginEmail);
                    if (!string.IsNullOrEmpty(_otherEmail)) iParam.Add("@OtherEmail", _otherEmail);
                    iParam.Add("@Status", _status.ToString());
                    if (!string.IsNullOrEmpty(_lastModified.ToString())) iParam.Add("@SFLastModifiedDate", _lastModified.ToString());
                    if (!string.IsNullOrEmpty(_productType)) iParam.Add("@ProductType", _productType);
                    if (!string.IsNullOrEmpty(_statePreference)) iParam.Add("@StatePreference", _statePreference);
                    if (!string.IsNullOrEmpty(_metros)) iParam.Add("@Metros", _metros);
                    if (!string.IsNullOrEmpty(_title)) iParam.Add("@Title", _title);
                    iParam.Add("@IPA", _ipa.ToString());
                    iParam.Add("@ResearchReports", _researchReports.ToString());
                    iParam.Add("@DoNotEmail", _doNotEmail.ToString());
                    iParam.Add("@CompanyRecord", _isCompanyRecord.ToString());
                    iParam.Add("@SFAssignedToID", _sfAssignedToId);
                    iParam.Add("@HasOptedOutOfEmail", _hasOptedOutOfEmail.ToString());
                    iParam.Add("@EmailContactDNC", _emailContactDNC.ToString());
                    //output parameters
                    oParam.Add("@ContactID", "4");
                    oParam.Add("@ErrCode", "4");
                    oParam.Add("@ErrMessage", "2048");
                    //run query 
                    output = dp.DoNonQueryGetMultipleOutputParam("[MMPeopleApp].[spSFContact_InsertUpdate]", CommandType.StoredProcedure, iParam, oParam);
                    if (output != null)
                    {
                        int iTemp = 0;
                        int.TryParse(output["@ContactID"].ToString(), out iTemp);
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
                    query += "SELECT Id, AccountId, OwnerId, FirstName, LastName, Business__c, DNC_Business__c, Phone, DNC_Direct_Phone__c, MobilePhone, DNC_Mobile__c, HomePhone, DNC_Home_Phone__c,";
                    query += " AssistantPhone, OtherPhone, DNC_Other_Phone__c, Fax, DNC_Fax__c, Email, Login_Email__c, Other_Email__c, Active__c, LastModifiedDate, IPA__c, Product_Type__c,";
                    query += " Acquisition_State_Preference__c, Research_Reports__c, ReportMarkets__c, Mail_Opt_Out__c, Title, Company_Record__c, Assigned_to_ID__c, HasOptedOutOfEmail, DNC_Email_Contact__c";
                    query += " FROM Contact";
                    query += " WHERE CreatedDate >= " + sfDate;
                    query += " OR LastModifiedDate >= " + sfDate;
                    qResult = _service.query(query);

                    if (qResult.size > 0)
                    {
                        int count = 0;
                        while (!queryComplete)
                        {
                            sObject[] records = qResult.records;
                            for (int i = 0; i < records.Length; i++)
                            {

                                wsSalesforce.Contact contact = (wsSalesforce.Contact)records[i];
                                //sync data here to mm databaseAccoun
                                MarcusMillichap_Contact c = new MarcusMillichap_Contact();
                                c.SFId = contact.Id;
                                c.SFAccountId = contact.AccountId;
                                c.SFOwnerId = contact.OwnerId;
                                c.FirstName = contact.FirstName;
                                c.LastName = contact.LastName;
                                c.BusinessPhone = contact.Business__c;
                                c.IsCompanyRecord = contact.Company_Record__c.HasValue ? contact.Company_Record__c.Value : false;
                                c.BusinessPhoneDnc = contact.DNC_Business__c.HasValue ? contact.DNC_Business__c.Value : false;
                                c.DirectPhone = contact.Phone;
                                c.DirectPhoneDnc = contact.DNC_Direct_Phone__c.HasValue ? contact.DNC_Direct_Phone__c.Value : false;
                                c.MobilePhone = contact.MobilePhone;
                                c.MobilePhoneDnc = contact.DNC_Mobile__c.HasValue ? contact.DNC_Mobile__c.Value : false;
                                c.HomePhone = contact.HomePhone;
                                c.HomePhoneDnc = contact.DNC_Home_Phone__c.HasValue ? contact.DNC_Home_Phone__c.Value : false;
                                c.AssistantDirectPhone = contact.AssistantPhone;
                                c.OtherPhone = contact.OtherPhone;
                                c.OtherPhoneDNC = contact.DNC_Other_Phone__c.HasValue ? contact.DNC_Other_Phone__c.Value : false;
                                c.Fax = contact.Fax;
                                c.FaxDnc = contact.DNC_Fax__c.HasValue ? contact.DNC_Fax__c.Value : false;
                                c.Email = contact.Email;
                                c.LoginEmail = contact.Login_Email__c;
                                c.OtherEmail = contact.Other_Email__c;
                                c.Status = contact.Active__c.HasValue ? contact.Active__c.Value : true;
                                c.IPA = contact.IPA__c.HasValue ? contact.IPA__c.Value : false;
                                c.DoNotEmail = contact.Mail_Opt_Out__c.HasValue ? contact.Mail_Opt_Out__c.Value : false;
                                c.ResearchReports = contact.Research_Reports__c.HasValue ? contact.Research_Reports__c.Value : false;
                                c.ProductType = BuildXml(contact.Product_Type__c);
                                c.StatePreference = BuildXml(contact.Acquisition_State_Preference__c);
                                c.Metros = BuildXml(contact.ReportMarkets__c);
                                c.LastModified = contact.LastModifiedDate;
                                c.Title = contact.Title;
                                c.SFAssignedToID = contact.Assigned_to_ID__c;
                                c.HasOptedOutOfEmail = contact.HasOptedOutOfEmail.HasValue ? contact.HasOptedOutOfEmail.Value : false;
                                c.EmailContactDNC = contact.DNC_Email_Contact__c.HasValue ? contact.DNC_Email_Contact__c.Value : false;
                                //update results
                                Results r = c.AddUpdateMarcusMillichap();
                                if (!r.Success)
                                {
                                    _failedContacts.Add(c);
                                }
                                count++;
                            }
                            if (qResult.done) queryComplete = true;
                            else qResult = _service.queryMore(qResult.queryLocator);
                        }
                    }
                }
                //retry the ones that failed
                if (_failedContacts != null && _failedContacts.Count > 0)
                {
                    Results rFailed;
                    foreach (MarcusMillichap_Contact failed in _failedContacts)
                    {
                        rFailed = failed.AddUpdateMarcusMillichap();
                        if (!rFailed.Success)
                        {
                            //do something here
                            //success = false;
                        }
                    }
                }

            }
            catch { success = false; }

            return success;
        }

        private string BuildXml(string value)
        {
            string returnVal = string.Empty;
            if (!string.IsNullOrEmpty(value))
            {
                string[] split = value.Split(';');
                if (split.Length > 0)
                {
                    XElement root = new XElement("data");
                    foreach (string s in split)
                    {
                        root.Add(new XElement("value", new XAttribute("text", s)));
                    }
                    returnVal = root.ToString();
                }
            }
            return returnVal;
        }
    }
}
