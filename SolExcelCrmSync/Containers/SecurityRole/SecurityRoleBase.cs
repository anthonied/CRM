using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace SolExcelCrmSync.Containers.SecurityRole
{
    public class SecurityRoleBase
    {
        public string sSecurityRoleName { get; set; }
        public Guid guidSecurityRoleId { get; set; }

        public void getSecurityRoleGuidBySecurityRoleName(IOrganizationService securityRoleCrmWebservice)
        {
            var myQuery = new QueryExpression
            {
                EntityName = "role",
                ColumnSet = new ColumnSet("roleid"),
                Criteria = new FilterExpression
                {
                    Conditions = 
                    {
                        new ConditionExpression
                        {
                          AttributeName  =  "name",
                          Operator = ConditionOperator.Equal,
                          Values = {"roleName"}
                        }
                    }
                }
            };

            //DataCollection<Entity> securityRoleResult = CrmWebservice.RetrieveMultiple(myQuery).Entities;
            DataCollection<Entity> securityRoleResult = securityRoleCrmWebservice.RetrieveMultiple(myQuery).Entities;
            if (securityRoleResult.Count > 0)
                this.guidSecurityRoleId = (Guid)securityRoleResult[0]["roleid"];
        }
    }

    
}
