using SalesforceSyncJobs.Models;
using SalesforceSyncJobs.wsSalesforce;
using System.Collections;

namespace SalesforceSyncJobs.Mappers
{
    public class SfDealMapper
    {
        public static Deal__c Map(string salesforceId, MMActivityModel activity)
        {
            Deal__c deal = new Deal__c();

            ArrayList nullFields = new ArrayList();

            deal.Id = salesforceId;
            deal.MNet_Activity_ID__c = activity.ActivityId;

            if (!string.IsNullOrEmpty(activity.City)) { deal.City__c = activity.City; }
            else { nullFields.Add("City__c"); }

            if (!string.IsNullOrEmpty(activity.Street)) { deal.Street__c = activity.Street; }
            else { nullFields.Add("Street__c"); }

            if (!string.IsNullOrEmpty(activity.State)) { deal.State__c = activity.State; }
            else { nullFields.Add("State__c"); }

            if (!string.IsNullOrEmpty(activity.Zip)) { deal.Zip_Code__c = activity.Zip; }
            else { nullFields.Add("Zip_Code__c"); }

            double dTemp;
            double.TryParse(activity.Price.ToString(), out dTemp);
            if (dTemp > 0)
            {
                deal.Price__c = dTemp;
                deal.Price__cSpecified = true;
            }
            else { nullFields.Add("Price__c"); }

            if (!string.IsNullOrEmpty(activity.Status)) { deal.Status__c = activity.Status; }
            else { nullFields.Add("Status__c"); }

            if (!string.IsNullOrEmpty(activity.SubStatus)) { deal.Sub_Status__c = activity.SubStatus; }
            else { nullFields.Add("Sub_Status__c"); }

            double.TryParse(activity.Commission.ToString(), out dTemp);
            if (dTemp > 0)
            {
                deal.Commission__c = dTemp;
                deal.Commission__cSpecified = true;
            }
            else { nullFields.Add("Commission__c"); }

            double.TryParse(activity.TotalCommission.ToString(), out dTemp);
            if (dTemp > 0)
            {
                deal.Total_Commission__c = dTemp;
                deal.Total_Commission__cSpecified = true;
            }
            else { nullFields.Add("Total_Commission__c"); }

            if (activity.YearBuilt.HasValue && activity.YearBuilt > 0) { deal.Year_Built__c = activity.YearBuilt.ToString(); }
            else { nullFields.Add("Year_Built__c"); }

            if (!string.IsNullOrEmpty(activity.Country)) { deal.Country__c = activity.Country; }
            else { nullFields.Add("Country__c"); }

            if (!string.IsNullOrEmpty(activity.StateIntl)) { deal.International_State__c = activity.StateIntl; }
            else nullFields.Add("International_State__c");

            if (nullFields.Count > 0)
            {
                string[] s = new string[nullFields.Count];
                for (int j = 0; j < nullFields.Count; j++)
                {
                    s[j] = nullFields[j].ToString();
                }
                deal.fieldsToNull = s;
            }

            return deal;
        }
    }
}
