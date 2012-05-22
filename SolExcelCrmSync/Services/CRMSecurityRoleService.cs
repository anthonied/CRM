using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using SolExcelCrmSync.Classes;
using Microsoft.Xrm.Sdk.Client;
using SolExcelCrmSync.Containers.SecurityRole;

namespace SolExcelCrmSync.Services
{
   public class CRMSecurityRoleService
    {
       public void getSecurityRoleFromCRM(SecurityRoleBase mySecurityRole)
       {         
           /*
           using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(CRMCredentials.OrganizationUri, null, CRMCredentials.Credentials, null))
           {
               IOrganizationService webservice = (IOrganizationService)serviceProxy;
               mySecurityRole.getSecurityRoleGuidBySecurityRoleName(webservice);
               
           }*/
       }

       public IOrganizationService serviceProxy { get; set; }
    }
}
