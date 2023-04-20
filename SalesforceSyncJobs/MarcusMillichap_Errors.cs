using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SalesforceSyncJobs
{
    public class MarcusMillichap_Errors
    {
        private string _salesforceId, _errorMessage, _emailAddress, _firstName, _lastName, _externalId;

        public string SalesforceId { get { return _salesforceId; } set { _salesforceId = value; } }
        public string ErrorMessage { get { return _errorMessage; } set { _errorMessage = value; } }
        public string EmailAddress { get { return _emailAddress; } set { _emailAddress = value; } }
        public string FirstName { get { return _firstName; } set { _firstName = value; } }
        public string LastName { get { return _lastName; } set { _lastName = value; } }
        public string ExternalId { get { return _externalId; } set { _externalId = value; } }
    }
}
