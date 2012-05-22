using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Configuration;

namespace SolExcelCrmSync.Services
{
    class CRMOrganizationService
    {
        private OrganizationServiceProxy _serviceProxy;
        public OrganizationServiceProxy ServiceProxy
        {
            get
            {
                return _serviceProxy;
            }
            set
            {
                _serviceProxy = value;
            }
        }

        public  ClientCredentials Credentials { get; set; }
        public  ClientCredentials DeviceCredentials { get; set; }
        public  Uri OrganizationUri { get; set; }

        public void SetCredentials()
        {
            Credentials = new ClientCredentials();
            Credentials.UserName.UserName = ConfigurationManager.AppSettings["CRMServiceUserName"];
            Credentials.UserName.Password = ConfigurationManager.AppSettings["CRMPassword"];
            OrganizationUri = new Uri(ConfigurationManager.AppSettings["OrganizationEndPointURI"]);

            DeviceCredentials = new ClientCredentials();
            DeviceCredentials.UserName.UserName = ConfigurationManager.AppSettings["CRMServiceUserName"];
            DeviceCredentials.UserName.Password = ConfigurationManager.AppSettings["CRMServicePassword"];
        }        

        public CRMOrganizationService()
        {
            SetCredentials();
            ServiceProxy = new OrganizationServiceProxy(OrganizationUri, null, Credentials, null);
        }

        
            

    }
}
