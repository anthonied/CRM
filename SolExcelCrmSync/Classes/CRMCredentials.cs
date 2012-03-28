using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel.Description;
using System.Configuration;

namespace SolExcelCrmSync.Classes
{
    public class CRMCredentials
    {
        public ClientCredentials Credentials { get; set; }
        public Uri OrganizationUri { get; set; }

        public CRMCredentials()
        {
            Credentials = new ClientCredentials();
            Credentials.UserName.UserName = ConfigurationManager.AppSettings["CRMServiceUserName"];
            Credentials.UserName.Password = ConfigurationManager.AppSettings["CRMServicePassword"];
            OrganizationUri = new Uri(ConfigurationManager.AppSettings["OrganizationEndPointURI"]);
        }

        
    }
}
