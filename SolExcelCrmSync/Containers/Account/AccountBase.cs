using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolExcelCrmSync.Containers.Account
{
    public class AccountBase
    {
        public string sCustomerNumber { get; set; }
        public string sAccountName { get; set; }
        public string sAddressLine1 { get; set; }
        public string sAddressLine2 { get; set; }
        public string sAddressLine3 { get; set; }
        public string sPostCode { get; set; }
        public string sTelephone { get; set; }        
        public string sEmail { get; set; }
    }
}
