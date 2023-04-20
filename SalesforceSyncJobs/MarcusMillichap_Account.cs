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
    class MarcusMillichap_Account
    {
        private List<MarcusMillichap_Account> _failedAccounts;
        private SforceService _service;
        private int _id;
        private string _sfId, _ownerId, _name, _phone, _fax;
        private bool _dnc, _ipa;
        private DateTime _lastModified;

        public int ID { get { return _id; } set { _id = value; } }
        public string SFId { get { return _sfId; } set { _sfId = value; } }
        public string OwnerId { get { return _ownerId; } set { _ownerId = value; } }
        public string Name { get { return _name; } set { _name = value; } }
        public string Phone { get { return _phone; } set { _phone = value; } }
        public string Fax { get { return _fax; } set { _fax = value; } }

        public bool DNC { get { return _dnc; } set { _dnc = value; } }
        public bool IPA { get { return _ipa; } set { _ipa = value; } }

        public DateTime LastModified { get { return _lastModified; } set { _lastModified = value; } }


        public MarcusMillichap_Account() {
            _id = 0;
            _sfId = string.Empty;
            _ownerId = string.Empty;
            _name = string.Empty;
            _phone = string.Empty;
            _fax = string.Empty;
            _dnc = false;
            _ipa = false;
            _lastModified = DateTime.Now;
            _failedAccounts = new List<MarcusMillichap_Account>();
        }

        public MarcusMillichap_Account(SforceService service)
        {
            _service = service;
            _id = 0;
            _sfId = string.Empty;
            _ownerId = string.Empty;
            _name = string.Empty;
            _phone = string.Empty;
            _fax = string.Empty;
            _dnc = false;
            _ipa = false;
            _lastModified = DateTime.Now;
            _failedAccounts = new List<MarcusMillichap_Account>();
        }

        public Results AddUpdateMarcusMillichap() {
            Results r = new Results();
            if (!string.IsNullOrEmpty(_sfId) && !string.IsNullOrEmpty(_ownerId)) {
                DataProvider dp = new DataProvider(ConfigurationManager.ConnectionStrings["MMPeople"].ToString());
                NameValueCollection iParam = new NameValueCollection();
                NameValueCollection oParam = new NameValueCollection();
                NameValueCollection output = new NameValueCollection();
                DataSet ds = new DataSet();
                try {
                    //input parameters
                    iParam.Add("@SFAccountID", _sfId);
                    iParam.Add("@SFOwnerID", _ownerId);
                    if (!string.IsNullOrEmpty(_name)) iParam.Add("@AccountName", _name);
                    if (!string.IsNullOrEmpty(_phone)) iParam.Add("@Phone", _phone);
                    iParam.Add("@PhoneDNC", _dnc.ToString());
                    if (!string.IsNullOrEmpty(_fax)) iParam.Add("@Fax", _fax);
                    iParam.Add("@IPA", _ipa.ToString());
                    iParam.Add("@SFLastModifiedDate", _lastModified.ToString());
                    //output parameters
                    oParam.Add("@AccountID", "4");
                    oParam.Add("@ErrCode", "4");
                    oParam.Add("@ErrMessage", "2048");
                    //run query 
                    output = dp.DoNonQueryGetMultipleOutputParam("[MMPeopleApp].[spSFAccount_InsertUpdate]", CommandType.StoredProcedure, iParam, oParam);
                    if (output != null) {
                        int iTemp = 0;
                        int.TryParse(output["@AccountID"].ToString(), out iTemp);
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
                    query += "SELECT Id, OwnerId, Name, Phone, Fax, IPA__c, LastModifiedDate";
                    query += " FROM Account";
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
                                wsSalesforce.Account account = (wsSalesforce.Account)records[i];
                                //sync data here to mm database
                                MarcusMillichap_Account a = new MarcusMillichap_Account();
                                a.SFId = account.Id;
                                a.OwnerId = account.OwnerId;
                                a.Name = account.Name;
                                a.Phone = account.Phone;
                                a.Fax = account.Fax;
                                a.IPA = account.IPA__c.HasValue ? account.IPA__c.Value : false;
                                a.LastModified = account.LastModifiedDate.HasValue ? account.LastModifiedDate.Value : DateTime.Now;
                                Results r = a.AddUpdateMarcusMillichap();
                                if (!r.Success) {
                                    _failedAccounts.Add(a);
                                }
                            }
                            if (qResult.done) queryComplete = true;
                            else qResult = _service.queryMore(qResult.queryLocator);
                        }
                    }
                }
                //retry the ones that failed
                if (_failedAccounts != null && _failedAccounts.Count > 0)
                {
                    Results rFailed;
                    foreach (MarcusMillichap_Account failed in _failedAccounts)
                    {
                        rFailed = failed.AddUpdateMarcusMillichap();
                        if (!rFailed.Success)
                        {
                            //success = false;
                            //do something here
                        }
                    }
                }

            }
            catch { success = false; }
            return success;
        }
    }
}
