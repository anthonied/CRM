using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace SolExcelCrmSync.Classes
{
    public static class Connect
    {
        public static string sConStr = ConfigurationManager.AppSettings["logdb"];
    }
}
