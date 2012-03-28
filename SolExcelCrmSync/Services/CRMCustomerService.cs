using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Client;
using SolExcelCrmSync.Classes;
using Microsoft.Xrm.Sdk;
using SolExcelCrmSync.Containers.Account;

namespace SolExcelCrmSync.Services
{
    public class CRMCustomerService
    {
        public void importFromSqlSolCRM(List<AccountBase> myCustomers)
        {
            var locCRMCredentials = new CRMCredentials();
            using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(locCRMCredentials.OrganizationUri, null, locCRMCredentials.Credentials, null))
            {
                IOrganizationService webservice = (IOrganizationService)serviceProxy;
                foreach (var Customer in myCustomers)
                {
                    Customer.addAccountToCRM(webservice);
                }
            }
        }

        
    }
}
