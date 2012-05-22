using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using SolExcelCrmSync.Containers.SecurityRole;

namespace SolExcelCrmSync.Containers.User
{
   public class UserBase:BaseContainer
    {
        public string sUserNumber { get; set; }
        public string sName { get; set; }
        public string sSurname { get; set; }
        public string sUsername { get; set; }
        public string Email { get; set; }
        public string sPhone { get; set; }
        public Guid guidUserId { get; set; }
        public string sCrmUsername { get; set; }

        public UserBase()
        {
            this.sCrmEntityName = "user";
            this.sCrmEntityGuidFieldName = "systemuserid";
            this.sCrmFilterAttributeName = "domainname";
        }


        public void AssignSecurityRole(SecurityRoleBase mySecurityRoleBase, IOrganizationService prmCrmWebService)
            //public void AssignSecurityRole(Guid prmUserId, Guid prmSecurityRoleId, IOrganizationService prmCrmWebService)
        {
            // Create new Associate Request object for creating a N:N link between User and Security
            AssociateRequest wod_AssosiateRequest = new AssociateRequest();

            // Create related entity reference object for associating relationship
            // In our case we will pass (SystemUser) record reference 

            wod_AssosiateRequest.RelatedEntities = new EntityReferenceCollection();
            wod_AssosiateRequest.RelatedEntities.Add(new EntityReference("systemuser", this.guidUserId));

            // Create new Relationship object for System User & Security Role entity schema and assigning it
            // to request relationship property

            wod_AssosiateRequest.Relationship = new Relationship("systemuserroles_association");

            // Create target entity reference object for associating relationship
            mySecurityRoleBase.getSecurityRoleGuidBySecurityRoleName(prmCrmWebService);
            wod_AssosiateRequest.Target = new EntityReference("role", mySecurityRoleBase.guidSecurityRoleId);

            // Passing AssosiateRequest object to Crm Service Execute method for assigning Security Role to User
            prmCrmWebService.Execute(wod_AssosiateRequest);
        }

    }
    
}
