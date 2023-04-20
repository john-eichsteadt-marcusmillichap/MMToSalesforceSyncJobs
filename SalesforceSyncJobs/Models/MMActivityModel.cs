using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SalesforceSyncJobs.Models
{
    public class MMActivityModel
    {
        public int? YearBuilt { get; set; }
        public int MPDId { get; set; }
        public string OwnerId { get; set; }
        public string SalesforceId { get; set; }
        public string ActivityId { get; set; }
        public string Name { get; set; }
        public string Street { get; set; }
        public string Zip { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Status { get; set; }
        public string SubStatus { get; set; }
        public string UserId { get; set; }
        public string Country { get; set; }
        public string StateIntl { get; set; }
        public double? Price { get; set; }
        public double? Commission { get; set; }
        public double? TotalCommission { get; set; }
    }
}
