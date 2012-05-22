using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel.Description;
using System.Configuration;

namespace SolExcelCrmSync.Classes
{
    public static class CRMCredentials
    {


        public static void SetCredentials(out ClientCredentials Credentials, out Uri OrganizationUri)
        {
            Credentials = new ClientCredentials();
            Credentials.UserName.UserName = @"pecs\\crmadmin";//ConfigurationManager.AppSettings["CRMServiceUserName"];
            Credentials.UserName.Password = "Password%%";//ConfigurationManager.AppSettings["CRMPassword"];
            OrganizationUri = new Uri(ConfigurationManager.AppSettings["OrganizationEndPointURI"]);
            /*
            DeviceCredentials = new ClientCredentials();
            DeviceCredentials.UserName.UserName = ConfigurationManager.AppSettings["CRMServiceUserName"];
            DeviceCredentials.UserName.Password = ConfigurationManager.AppSettings["CRMServicePassword"];*/
        }        
    }
}
