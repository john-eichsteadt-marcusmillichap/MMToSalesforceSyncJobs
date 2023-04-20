using System;

namespace SalesforceSyncJobs.Models
{
    public class MMPropertyModel
    {
        public string SalesForceId { get; set; }
        public int? RentableSqFt { get; set; }
        public int? NumberofUnits { get; set; }
        public int? YearBuilt { get; set; }
        public int UserCode { get; set; }
        public string Address { get; set; }
        public string MpdId { get; set; }
        public string Name { get; set; }
        public string City { get; set; }
        public string StateId { get; set; }
        public string SubType { get; set; }
        public string UserId { get; set; }
        public string OwnerId { get; set; }
        public string Country { get; set; }
        public string StateIntl { get; set; }
        public string Zip { get; set; }
        public string County { get; set; }
        public string[] OwnerIdList { get; set; }
        public DateTime? LastSaleDate { get; set; }
    }
}
