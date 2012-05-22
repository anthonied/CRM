using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace SolExcelCrmSync.Containers
{
    public class BaseContainer
    {
        public Guid crmGuidId { get; set; }
        public string sCrmEntityName { get; set; }
        public string sCrmEntityGuidFieldName { get; set; }
        public string sCrmFilterAttributeName { get; set; }
        public string sCrmFilterAttributeValue { get; set; }

        public void FindGuidForObject(IOrganizationService myCRMWebservice)
        {
            var myQuery = new QueryExpression
            {
                EntityName = this.sCrmEntityName,
                ColumnSet = new ColumnSet(this.sCrmEntityGuidFieldName),
                Criteria = new FilterExpression
                {
                    Conditions = 
                    {
                        new ConditionExpression
                        {
                          AttributeName  =  this.sCrmFilterAttributeName,
                          Operator = ConditionOperator.Equal,
                          Values = {this.sCrmFilterAttributeValue}
                        }
                    }
                }
                
            };
        
            DataCollection<Entity> accountResult = myCRMWebservice.RetrieveMultiple(myQuery).Entities;
            if (accountResult.Count > 0)
                this.crmGuidId = (Guid)accountResult[0][this.sCrmEntityGuidFieldName];
        }
    }
}
