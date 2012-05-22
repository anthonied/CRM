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
using System.Configuration;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.ComponentModel;
using SolExcelCrmSync.Containers.Account;
using SolExcelCrmSync.Domain;

namespace SolExcelCrmSync.Classes
{
    class CRMEntityMananger
    {
        public ClientCredentials Credentials { get; set; }
        public Uri OrganizationUri { get; set; }
        //public CRMCredentials locCRMCredentials;
                
        public CRMEntityMananger()
        {
          ////  Credentials = new ClientCredentials();
          ////  Credentials.UserName.UserName = "pecs\\crmadmin";
          ////  Credentials.UserName.Password = "Password%%";
          ////  OrganizationUri = new Uri(ConfigurationManager.AppSettings["OrganizationEndPointURI"]);
          //////  locCRMCredentials = new CRMCredentials();
        }

        public string updateAccountDetails(List<AccountExcel> excelAccountDetails, ref int iAdded, ref int iUpdated, ref List<string> lNewCustomers, BackgroundWorker CustomerBackGroundWorkerBackgroundWorker)
        {
            var iCounter = 0;
            lNewCustomers = new List<string>();
            CustomerBackGroundWorkerBackgroundWorker.ReportProgress(0, "Connecting to server...");
            using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, null, Credentials, null))
            //using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(CRMCredentials.OrganizationUri, null, CRMCredentials.Credentials, null))
            {
                CustomerBackGroundWorkerBackgroundWorker.ReportProgress(0, "Successfully connected to the server");
                IOrganizationService service = (IOrganizationService)serviceProxy;
                CustomerBackGroundWorkerBackgroundWorker.ReportProgress(0, excelAccountDetails.Count.ToString() + " lines to sync.");
                CustomerBackGroundWorkerBackgroundWorker.ReportProgress(excelAccountDetails.Count, "Starting to sync");
                iCounter = 0;
                foreach (var theAccount in excelAccountDetails)
                {
                    iCounter++;
                    CustomerBackGroundWorkerBackgroundWorker.ReportProgress(iCounter);
                    if (theAccount.sCustomerNumber != "")
                    {
                        
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
                          Values = {theAccount.sCustomerNumber}
                        }
                    }
                            }
                        };

                        DataCollection<Entity> entities = service.RetrieveMultiple(myAccountQuery).Entities;

                        foreach (var myEntity in entities)
                        {

                            myEntity["new_customernumber"] = theAccount.sCustomerNumber;
                            myEntity["name"] = theAccount.sAccountName;
                            myEntity["address1_line1"] = theAccount.sAddressLine1;
                            myEntity["address1_line2"] = theAccount.sAddressLine2;
                            myEntity["address1_line3"] = theAccount.sAddressLine3;
                            if (theAccount.sPostCode.Length > 20)
                                theAccount.sPostCode = theAccount.sPostCode.Substring(0, 20);
                            myEntity["address1_postalcode"] = theAccount.sPostCode;

                            if (theAccount.sAddressLine1.Length > 20)
                                theAccount.sAddressLine1 = theAccount.sAddressLine1.Substring(0, 20);
                            myEntity["address2_postofficebox"] = theAccount.sAddressLine1;
                            myEntity["address2_city"] = theAccount.sAddressLine2;
                            myEntity["address2_stateorprovince"] = theAccount.sAddressLine3;
                            myEntity["address2_postalcode"] = theAccount.sPostCode;

                            myEntity["address1_telephone1"] = theAccount.sTelephone;
                            myEntity["telephone1"] = theAccount.sTelephone;
                            myEntity["new_vatnumber"] = theAccount.sVatNumber;
                            myEntity["address1_country"] = theAccount.sCountryCode;
                            myEntity["emailaddress1"] = theAccount.sEmail;
                            myEntity["websiteurl"] = theAccount.sWeb;
                            myEntity["address1_stateorprovince"] = theAccount.sRegion;
                            myEntity["address2_stateorprovince"] = theAccount.sRegion;
                            myEntity["new_kam"] = theAccount.sKAM;

                            service.Update(myEntity);
                            iUpdated++;
                        }
                        if (iCounter == 0)
                        {


                            var myNewEntity = new Entity("account");
                            myNewEntity["new_customernumber"] = theAccount.sCustomerNumber;
                            myNewEntity["name"] = theAccount.sAccountName;
                            myNewEntity["address1_line1"] = theAccount.sAddressLine1;
                            myNewEntity["address1_line2"] = theAccount.sAddressLine2;
                            myNewEntity["address1_line3"] = theAccount.sAddressLine3;
                            myNewEntity["address1_postalcode"] = theAccount.sPostCode;
                            myNewEntity["address1_telephone1"] = theAccount.sTelephone;
                            if (theAccount.sAddressLine1.Length > 20)
                                theAccount.sAddressLine1 = theAccount.sAddressLine1.Substring(0, 20);

                            myNewEntity["address2_postofficebox"] = theAccount.sAddressLine1;
                            myNewEntity["address2_city"] = theAccount.sAddressLine2;
                            myNewEntity["address2_stateorprovince"] = theAccount.sAddressLine3;
                            myNewEntity["address2_postalcode"] = theAccount.sPostCode;

                            myNewEntity["telephone1"] = theAccount.sTelephone;
                            myNewEntity["new_vatnumber"] = theAccount.sVatNumber;
                            myNewEntity["address1_country"] = theAccount.sCountryCode;
                            myNewEntity["emailaddress1"] = theAccount.sEmail;
                            myNewEntity["websiteurl"] = theAccount.sWeb;
                            myNewEntity["Address1_StateProvince"] = theAccount.sRegion;
                            myNewEntity["Address2_StateProvince"] = theAccount.sRegion;
                            myNewEntity["new_kam"] = theAccount.sKAM;
                            service.Create(myNewEntity);
                            iAdded++;
                            lNewCustomers.Add(theAccount.sCustomerNumber);

                        }
                    }
                }
            }
            return iUpdated.ToString() + " - " + iAdded.ToString();
        }

        private string RetrieveAttributeMetadataPicklistValue(IOrganizationService CrmService, int AddTypeValue, string Entity, string LogicalName)
        {

            RetrieveAttributeRequest objRetrieveAttributeRequest;

            objRetrieveAttributeRequest = new RetrieveAttributeRequest();

            objRetrieveAttributeRequest.EntityLogicalName = Entity;

            objRetrieveAttributeRequest.LogicalName = LogicalName;

            int indexAddtypeCode = AddTypeValue - 1;

            // Execute the request

            RetrieveAttributeResponse attributeResponse = (RetrieveAttributeResponse)CrmService.Execute(objRetrieveAttributeRequest);

            PicklistAttributeMetadata objPckLstAttMetadata = new PicklistAttributeMetadata();

            ICollection<object> objCollection = attributeResponse.Results.Values;

            objPckLstAttMetadata.OptionSet = ((EnumAttributeMetadata)(objCollection.ElementAt(0))).OptionSet;

            Microsoft.Xrm.Sdk.Label objLabel = objPckLstAttMetadata.OptionSet.Options[indexAddtypeCode].Label;

            string lblAddTypeLabel = objLabel.LocalizedLabels.ElementAt(0).Label;

            return lblAddTypeLabel;

        }



       
    }
}
