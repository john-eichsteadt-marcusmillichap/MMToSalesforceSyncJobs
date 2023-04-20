using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SalesforceSyncJobs
{
    class Results
    {
        private int _id, _errorCode;
        private string _errorMessage;
        private bool _success;

        public int ID { get { return _id; } set { _id = value; } }
        public int ErrorCode { get { return _errorCode; } set { _errorCode = value; } }
        public string ErrorMessage { get { return _errorMessage; } set { _errorMessage = value; } }
        public bool Success { get { return _success; } set { _success = value; } }

        public Results()
        {
            _id = 0;
            _errorCode = 0;
            _errorMessage = string.Empty;
            _success = false;
        }
    }
}
