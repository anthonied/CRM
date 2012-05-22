using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolExcelCrmSync.Containers.Account
{
    public class AccountExcel:AccountBase
    {
        public string sCountryCode { get; set; }
        public string sWeb { get; set; }
        public string sKAM { get; set; }
        public string sVatNumber { get; set; }
    }
}
