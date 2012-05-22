using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using SolExcelCrmSync.Classes;
using SolExcelCrmSync.Containers.User;
using Microsoft.Xrm.Sdk.Client;

namespace SolExcelCrmSync.Services
{
   public class CRMUserService
    {
       public void getUserFromCRM(UserBase myUser)
       {    
           /*
           using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(CRMCredentials.OrganizationUri, null, CRMCredentials.Credentials, null))
           {
               IOrganizationService webservice = (IOrganizationService)serviceProxy;
               myUser.FindGuidForObject(webservice);
           }*/
       }


       public IOrganizationService serviceProxy { get; set; }
    }
}
