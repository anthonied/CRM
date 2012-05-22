using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace SolExcelCrmSync.Containers.Account
{
    public class AccountBase:BaseContainer
    {
        public string sCustomerNumber { get; set; }
        public string sAccountName { get; set; }
        public string sAddressLine1 { get; set; }
        public string sAddressLine2 { get; set; }
        public string sAddressLine3 { get; set; }
        public string sPostCode { get; set; }
        public string sTelephone { get; set; }        
        public string sEmail { get; set; }
        public string sRegion { get; set; }
        public Guid guidAccountId { get; set; }

        public AccountBase()
        {
            this.sCrmEntityName = "account";
            this.sCrmEntityGuidFieldName = "accountid";
            this.sCrmFilterAttributeName = "accountnumber";
        }



        public void addAccountToCRM(IOrganizationService myCRMServiceProxy)
        {
            var newAccount = new Entity("account");
            newAccount["accountnumber"] = this.sCustomerNumber;
            newAccount["name"] = this.sAccountName;
            newAccount["address1_line1"] = this.sAddressLine1;
            newAccount["address1_line2"] = this.sAddressLine2;
            newAccount["address1_line3"] = this.sAddressLine3;
            newAccount["emailaddress1"] = this.sEmail;
            newAccount["address1_postalcode"] = this.sPostCode;
            newAccount["telephone1"] = this.sTelephone;
            myCRMServiceProxy.Create(newAccount);
        }

        public void getAccountGuidByCustomerNumber(IOrganizationService myCRMWebservice)
        {
            var myQuery = new QueryExpression
            {
                EntityName = "account",
                ColumnSet = new ColumnSet("accountid"),
                Criteria = new FilterExpression
                {
                    Conditions = 
                    {
                        new ConditionExpression
                        {
                          AttributeName  =  "accountnumber",
                          Operator = ConditionOperator.Equal,
                          Values = {"accountNumber"}
                        }
                    }
                }
            };
            DataCollection<Entity> accountResult = myCRMWebservice.RetrieveMultiple(myQuery).Entities;
            if (accountResult.Count > 0)
                this.guidAccountId = (Guid)accountResult[0]["accountid"];
        }

        public string EntityName { get; set; }
    }
}
