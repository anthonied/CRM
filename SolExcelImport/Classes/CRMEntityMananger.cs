using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel.Description;

using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Net;
using Microsoft.Xrm.Sdk;
using System.Runtime.Serialization;


namespace SolExcelImport.Classes
{
    public class CRMEntityMananger
    {
        public ClientCredentials Credentials { get; set; }
        public Uri OrganizationUri { get; set; }


        public CRMEntityMananger()
        {
            Credentials = new ClientCredentials();
            Credentials.UserName.UserName = "pecs\\crmadmin";
            Credentials.UserName.Password = "Password%%";
            OrganizationUri = new Uri("http://crm.pecs.co.za:5555/PECorporateServices/XRMServices/2011/Organization.svc");
        }

        public string updateAccountDetails(accountDetails excelAccountDetails)
        {

            using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, null, Credentials, null))
            {
                IOrganizationService service = (IOrganizationService)serviceProxy;
                QueryExpression myAccountQuery = new QueryExpression
                {
                    EntityName = "account",
                    ColumnSet = new ColumnSet("accountid"),
                    Criteria = new FilterExpression
                    {
                        Conditions = 
                    {
                        new ConditionExpression
                        {
                          AttributeName  =  "new_customernumber",
                          Operator = ConditionOperator.Equal,
                          Values = {excelAccountDetails.sCustomerNumber}
                        }
                    }
                    }
                };

                DataCollection<Entity> entities = service.RetrieveMultiple(myAccountQuery).Entities;
                foreach (var myEntity in entities)
                {
                    myEntity["new_customernumber"] = excelAccountDetails.sCustomerNumber;
                    myEntity["name"] = excelAccountDetails.sAccountName;
                    myEntity["address1_line1"] = excelAccountDetails.sAddressLine1;
                    myEntity["address1_line2"] = excelAccountDetails.sAddressLine2;
                    myEntity["address1_line3"] = excelAccountDetails.sAddressLine3;
                    myEntity["address1_postalcode"] = excelAccountDetails.sPostCode;
                    myEntity["address1_telephone1"] = excelAccountDetails.sTelephone;
                    myEntity["telephone1"] = excelAccountDetails.sTelephone;
                    myEntity["new_vatnumber"] = excelAccountDetails.sVatNumber;
                    myEntity["address1_country"] = excelAccountDetails.sCountryCode;
                    myEntity["emailaddress1"] = excelAccountDetails.sEmail;
                    myEntity["Websiteurl"] = excelAccountDetails.sWeb;
                    service.Update(myEntity);
                }
            }
            return "0";
        }
    }

    public class accountDetails
    {
        public string sCustomerNumber { get; set; }
        public string sAccountName { get; set; }
        public string sAddressLine1 { get; set; }
        public string sAddressLine2 { get; set; }
        public string sAddressLine3 { get; set; }
        public string sPostCode { get; set; }
        public string sTelephone { get; set; }
        public string sVatNumber { get; set; }
        public string sCountryCode { get; set; }
        public string sEmail { get; set; }
        public string sWeb { get; set; }
    }

}