using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace SalesforceSyncJobs
{
    class MarcusMillichap_Common
    {
        /// <summary>
        /// Salesforce DateTime conversion
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string ConvertToSalesforceDateTime(DateTime dt)
        {
            try
            {
                dt = DateTime.SpecifyKind(dt, DateTimeKind.Utc);
                return dt.ToString("yyyy-MM-ddTHH:mm:ssZ");
            }
            catch { return string.Empty; }
        }

        /// <summary>
        /// Update last runtime in app.config file
        /// </summary>
        /// <param name="lastRunTime"></param>
        /// <param name="keyName"></param>
        public static void UpdateLastRunTime(DateTime lastRunTime, string keyName)
        {
            try
            {
                Configuration appConfiguration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                appConfiguration.AppSettings.Settings[keyName].Value = lastRunTime.ToString();
                appConfiguration.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }
            catch (Exception ex)
            {
                //    Email e = new Email(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                //                        ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                //                       "SalesForce Lender API Error - Common.cs - UpdateLastRunTime() ",
                //                        "<strong>Error Message:</strong> <br /><br />" + ex.Message +
                //                        "<br /><br /><strong>Inner Exception:</strong> <br /><br />" + ex.InnerException +
                //                        "<br /><br /><strong>Stack Trace:</strong> <br /><br />" + ex.StackTrace,
                //                        true);
                //    e.SendEmail();
            }
        }

        public static Dictionary<string, MarcusMillichap_Group> AppendNewGroups(Dictionary<string, MarcusMillichap_Group> newGroups, Dictionary<string, MarcusMillichap_Group> existingGroups)
        {
            if (newGroups != null && newGroups.Count > 0)
            {
                foreach (var g in newGroups)
                {
                    if (!existingGroups.ContainsKey(g.Key))
                    {
                        existingGroups.Add(g.Key, g.Value);
                    }
                }
            }
            return existingGroups;
        }

        public static string FormatPhoneNumberDisplay(string p)
        {
            if (p.Length == 7) return FormatVOIPNumber(p);
            else if (p.Length == 10) return FormatPhoneNumber(p);
            else return p;
        }

        /// <summary>
        /// Format Phone Number
        /// </summary>
        /// <param name="phoneNumber"></param>
        /// <returns></returns>
        public static string FormatVOIPNumber(string phoneNumber)
        {
            phoneNumber = phoneNumber.Replace("(", "").Replace(")", "").Replace("-", "");
            string results = string.Empty;
            string formatPattern = @"(\d{3})(\d{4})";
            results = System.Text.RegularExpressions.Regex.Replace(phoneNumber, formatPattern, "$1-$2");
            //now return the formatted phone number
            return results;
        }

        /// <summary>
        /// Format Phone Number
        /// </summary>
        /// <param name="phoneNumber"></param>
        /// <returns></returns>
        public static string FormatPhoneNumber(string phoneNumber, out string extension)
        {
            extension = string.Empty;
            phoneNumber = phoneNumber.Replace("(", "").Replace(")", "").Replace("-", "");
            if (phoneNumber != null && phoneNumber.Length == 10)
                return "(" + phoneNumber.Substring(0, 3) + ") " + phoneNumber.Substring(3, 3).ToString() + "-" + phoneNumber.Substring(6, 4).ToString();
            else if (phoneNumber.Length > 10)
            {
                extension = phoneNumber.Substring(7, phoneNumber.Length - 10).ToString();
                return "(" + phoneNumber.Substring(0, 3) + ") " + phoneNumber.Substring(3, 3).ToString() + "-" + phoneNumber.Substring(6, 4).ToString();
            }
            else
                return phoneNumber;
        }

        /// <summary>
        /// Format Phone Number
        /// </summary>
        /// <param name="phoneNumber"></param>
        /// <returns></returns>
        public static string FormatPhoneNumber(string phoneNumber)
        {
            phoneNumber = phoneNumber.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").Replace(".", "").Trim();
            if (phoneNumber != null && phoneNumber.Length == 10)
                return "(" + phoneNumber.Substring(0, 3) + ") " + phoneNumber.Substring(3, 3).ToString() + "-" + phoneNumber.Substring(6, 4).ToString();
            else if (phoneNumber.Length > 10)
            {
                return "(" + phoneNumber.Substring(0, 3) + ") " + phoneNumber.Substring(3, 3).ToString() + "-" + phoneNumber.Substring(6, 4).ToString() + " x " + phoneNumber.Substring(7, phoneNumber.Length - 10).ToString();
            }
            else
                return phoneNumber;
        }

        public static void AppendNewUsers(Dictionary<string, wsSalesforce.User> newlyAddedUsers, ref Dictionary<string, MarcusMillichap_User> _existingSalesForceUsers, ref Dictionary<int, MarcusMillichap_User> _existingSalesForceUsersUserCodeKey)
        {

            if (newlyAddedUsers != null && newlyAddedUsers.Count > 0)
            {
                foreach (var u in newlyAddedUsers)
                {
                    MarcusMillichap_User users = ConvertUserToUsersObject(u.Value);
                    if (!_existingSalesForceUsers.ContainsKey(u.Key))
                    {
                        users.SalesforceId = u.Key;
                        _existingSalesForceUsers.Add(u.Key, users);
                    }
                    if (users.UserCode > 0 && !_existingSalesForceUsersUserCodeKey.ContainsKey(users.UserCode))
                    {
                        _existingSalesForceUsersUserCodeKey.Add(users.UserCode, users);
                    }
                }
            }
        }

        private static MarcusMillichap_User ConvertUserToUsersObject(wsSalesforce.User u)
        {
            MarcusMillichap_User users = new MarcusMillichap_User();
            if (u != null)
            {
                int temp = 0;
                users.FirstName = u.FirstName;
                users.LastName = u.LastName;
                users.Title = u.Title;
                users.Department = u.Department;
                users.Company = u.CompanyName;
                users.Alias = u.Alias;
                users.UserName = u.Username;
                users.EmailAddress = u.Email;
                users.CommunityNickName = string.Empty;
                users.ProfileId = u.ProfileId;
                users.EmailEncoding = u.EmailEncodingKey;
                users.FederationId = u.FederationIdentifier;
                users.TimeZone = u.TimeZoneSidKey;
                users.Locale = u.LocaleSidKey;
                users.Language = u.LanguageLocaleKey;
                if (u.MNetUserCode__c.ToSafeString() != "")
                    users.UserCode = int.TryParse(u.MNetUserCode__c, out temp) ? temp : 0;
                users.ADUserId = u.CommunityNickname;
            }
            return users;
        }

        public static bool IsDirty(Dictionary<int, MarcusMillichap_User> salesforceUsers,
            Dictionary<int, MarcusMillichap_User> mmUsers,
            int key)
        {

            var salesforceUser = salesforceUsers[key];
            var mmUser = mmUsers[key];

            bool isDirty = false;

            if (!salesforceUser.FirstName.ToSafeString().ToLower().Equals(mmUser.FirstName.ToSafeString().ToLower())) isDirty = true;
            if (!salesforceUser.LastName.ToSafeString().ToLower().Equals(mmUser.LastName.ToSafeString().ToLower())) isDirty = true;
            // set to 80 since there is a CAP in salesforce
            if (!salesforceUser.Title.ToSafeString().ToLower().Equals(mmUser.Title.ToSafeString().ToMaxLength(80).ToLower())) isDirty = true;
            if (!salesforceUser.Company.ToSafeString().ToLower().Equals(mmUser.Company.ToSafeString().ToLower())) isDirty = true;
            // set to 8 since there is a CAP in salesforce
            if (!salesforceUser.Alias.ToSafeString().ToLower().Equals(mmUser.Alias.ToSafeString().ToMaxLength(8).ToLower())) isDirty = true;
            if (!salesforceUser.EmailAddress.ToSafeString().ToLower().Equals(mmUser.EmailAddress.ToSafeString().ToLower())) isDirty = true;
            if (!salesforceUser.CommunityNickName.ToSafeString().ToLower().Equals(mmUser.CommunityNickName.ToSafeString().ToLower())) isDirty = true;
            if (!salesforceUser.UserCode.ToSafeString().Equals(mmUser.UserCode.ToSafeString())) isDirty = true;
            if (!salesforceUser.UserCodeManager.ToSafeString().Equals(mmUser.UserCodeManager.ToSafeString())
                && mmUser.UserCodeManager.HasValue
                && salesforceUsers.ContainsKey(mmUser.UserCodeManager.Value))
            {
                isDirty = true;
            }
            if (!salesforceUser.Phone.ToSafeString().Equals(mmUser.Phone.ToSafeString().CleanPhoneNumber())) isDirty = true;
            if (!salesforceUser.Fax.ToSafeString().Equals(mmUser.Fax.ToSafeString().CleanPhoneNumber())) isDirty = true;
            if (!salesforceUser.Mobile.ToSafeString().Equals(mmUser.Mobile.ToSafeString().CleanPhoneNumber())) isDirty = true;
            if (!salesforceUser.Flag.ToSafeString().Equals(mmUser.Flag.ToSafeString())) isDirty = true;
            if (!salesforceUser.OfficeName.ToSafeString().ToLower().Equals(mmUser.OfficeName.ToSafeString().ToLower())) isDirty = true;
            if (!salesforceUser.TimeZone.ToSafeString().ToLower().Equals(mmUser.TimeZone.ToSafeString().ToLower())) isDirty = true;
            if (!salesforceUser.ProfileId.ToLower().Equals(mmUser.ProfileId.ToLower())) isDirty = true;
            if (!salesforceUser.FederationId.ToLower().Equals(mmUser.FederationId.ToLower())) isDirty = true;
            return isDirty;

        }


        public static MarcusMillichap_User GetMatchingUserInformation(string salesforceUserId, Dictionary<string, MarcusMillichap_User> allUsers)
        {
            if (allUsers.ContainsKey(salesforceUserId)) return allUsers[salesforceUserId];
            else return null;
        }

        public static string GetMNetADUserIDFromSalesforceUsers(string salesforceUserId, Dictionary<string, MarcusMillichap_User> allUsers)
        {
            if (allUsers.ContainsKey(salesforceUserId))
            {
                MarcusMillichap_User u = allUsers[salesforceUserId];
                return u.ADUserId;
            }
            else return string.Empty;
        }

        public static string FindUserByUserCode(int userCode, Dictionary<string, MarcusMillichap_User> allUsers)
        {
            string salesforceId = string.Empty;
            var values = from value in allUsers.Values
                         where value.UserCode.Equals(userCode)
                         select value;
            List<MarcusMillichap_User> u = values.ToList();
            if (u != null && u.Count > 0)
            {
                salesforceId = u[0].SalesforceId;
            }

            return salesforceId;
        }

        public static string FindUserByUserCode(string userCode, Dictionary<string, wsSalesforce.User> allUsers)
        {
            string salesforceId = string.Empty;
            var values = from value in allUsers.Values
                         where value.MNetUserCode__c.Equals(userCode)
                         select value;
            List<wsSalesforce.User> u = values.ToList();
            if (u != null && u.Count > 0)
            {
                salesforceId = u[0].Id;
            }

            return salesforceId;
        }

        /// <summary>
        /// Get Property Name
        /// </summary>
        /// <param name="propertyName"></param>
        /// <param name="propertyAddress"></param>
        /// <returns></returns>
        public static string GetPropertyNameOrAddress(string propertyName, string propertyAddress)
        {
            if (!String.IsNullOrEmpty(propertyName)) return propertyName;
            else return propertyAddress;
        }

        /// <summary>
        /// Get full address
        /// </summary>
        /// <param name="address1"></param>
        /// <param name="address2"></param>
        /// <returns></returns>
        public static string GetPropertyAddress(string address1, string address2)
        {
            if (!String.IsNullOrEmpty(address2)) return address1 + " " + address2;
            else return address1;
        }

        /// <summary>
        /// Get full zip
        /// </summary>
        /// <param name="zip5"></param>
        /// <param name="zip4"></param>
        /// <returns></returns>
        public static string GetPropertyZip(string zip5, string zip4)
        {
            if (!String.IsNullOrEmpty(zip4)) return zip5 + "-" + zip4;
            else return zip5;
        }

        /// <summary>
        /// Format price
        /// </summary>
        /// <param name="price"></param>
        /// <returns></returns>
        public static double? FormatPrice(object price)
        {
            double tempParse;
            double? value = null;
            if (price != null)
            {
                if (double.TryParse(price.ToString(), out tempParse)) value = tempParse;
                else value = null;
            }
            return value;
        }


        /// <summary>
        /// Helper Function: Breaks List in Lists of max size (currently 200)
        /// </summary>
        /// <param name="list"></param>
        /// <param name="chunkSize"></param>
        /// <returns></returns>
        public static List<List<T>> BreakIntoChunks<T>(List<T> list, int chunkSize)
        {
            List<List<T>> retVal = new List<List<T>>();
            if (chunkSize > 0)
            {
                int index = 0;
                while (index < list.Count)
                {
                    int count = list.Count - index > chunkSize ? chunkSize : list.Count - index;
                    retVal.Add(list.GetRange(index, count));
                    index += chunkSize;
                }
            }
            return retVal;
        }


    }

    static class CommonExtension
    {
        public static string ToSafeString(this object obj)
        {
            return obj != null ? obj.ToString() : String.Empty;
        }

        public static string ToMaxLength(this object obj, int parameter)
        {
            if (obj != null && obj.ToString().Length > parameter) return obj.ToString().Substring(0, parameter);
            else return obj.ToString();
        }

        public static string CleanPhoneNumber(this object obj)
        {
            if (obj != null && obj.ToString().Length > 0)
            {
                return obj.ToString().Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", "").Replace(".", "").Trim();
            }
            else return string.Empty;
        }

        public static int? ToNullableInt(this object obj)
        {
            if (obj != null)
            {
                int temp = 0;
                if (int.TryParse(obj.ToString(), out temp)) return temp;
                else return null;
            }
            else return null;
        }

        public static string CleanMMVoipNumber(this object obj)
        {
            if (obj != null)
            {
                string temp = obj.ToString();
                if (temp.Length > 10)
                {
                    temp = temp.Substring((temp.Length - 10), 10);
                    return temp;
                }
                else return temp;
            }
            else return null;
        }

        public static string SetWildCardSearch(this object obj)
        {
            if (obj != null)
            {
                string phone = obj.ToString();
                if (phone.Length == 10)
                {
                    return "%" + phone.Substring(0, 3) + "%" + phone.Substring(3, 3) + "%" + phone.Substring(6, 4) + "%";
                }
                else if (phone.Length > 10)
                {
                    string newPhone = phone.Substring(phone.Length - 10, 10);
                    return "%" + newPhone.Substring(0, 3) + "%" + newPhone.Substring(3, 3) + "%" + newPhone.Substring(6, 4) + "%";
                }
                else
                {
                    return string.Empty;
                }

            }
            else return string.Empty;
        }

    }
}
