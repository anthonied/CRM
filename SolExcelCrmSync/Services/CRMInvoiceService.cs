using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SolExcelCrmSync.Domain;
using System.ComponentModel;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SolExcelCrmSync.Classes;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.ServiceModel.Description;
using System.Configuration;

namespace SolExcelCrmSync.Services
{
    class CRMInvoiceService
    {
        private ClientCredentials Credentials;
        private Uri OrganizationUri;

        public CRMInvoiceService()
        {
            Credentials = new ClientCredentials();
            Credentials.UserName.UserName = "pecs\\crmadmin";
            Credentials.UserName.Password = "Password%%";
            OrganizationUri = new Uri(ConfigurationManager.AppSettings["OrganizationEndPointURI"]);
        }
        public List<Invoice> ListInvoices(ref BackgroundWorker ActiveBackGroundWorker, DateTime LastExport)
        {
            ActiveBackGroundWorker.ReportProgress(0, "Connecting to server...");
            
            using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, null, Credentials, null))
            {
                ActiveBackGroundWorker.ReportProgress(0, "Connection establish to server");
                ActiveBackGroundWorker.ReportProgress(0, "Retrieving invoices...");
                IOrganizationService service = (IOrganizationService)serviceProxy;
                QueryExpression myInvoiceQuery = new QueryExpression
                {
                    EntityName = "invoice",
                    ColumnSet = new ColumnSet("invoiceid", "new_customernumber", "new_customerrefinvoice", "new_pecsinvoiceref", "new_invoicedate", "name", "new_salesperson", "shippingmethodcode", "new_processedby", "new_specialinstructionsnote1", "new_titlespecialinstructionsnote2"),
                    Criteria = new FilterExpression
                    {
                        Conditions = 
                    {
                        new ConditionExpression
                        {
                          AttributeName  =  "new_pecsinvoiceref",
                          Operator = ConditionOperator.NotNull                          
                        },
                        new ConditionExpression
                        {
                            AttributeName = "statuscode",
                            Operator = ConditionOperator.Equal,
                            Values = {"100001"}
                        }
                    }
                    }
                };

                if (LastExport.Year != 1970) //First Export
                {
                    myInvoiceQuery.Criteria.AddCondition(new ConditionExpression
                    {
                        AttributeName = "createdon",
                        Operator = ConditionOperator.GreaterThan,
                        Values = { LastExport }
                    });
                }
                DataCollection<Entity> entities = service.RetrieveMultiple(myInvoiceQuery).Entities;
                ActiveBackGroundWorker.ReportProgress(0, entities.Count.ToString() + " invoice entries found.");
                ActiveBackGroundWorker.ReportProgress(entities.Count, "Starting export process...");

                Invoice myInvoice;
                var lInvoices = new List<Invoice>();
                int iCounter = 0;

                foreach (var myEntity in entities)
                {
                    iCounter++;
                    ActiveBackGroundWorker.ReportProgress(iCounter);
                    if (myEntity.Contains("new_pecsinvoiceref")) //List 
                    {
                        //****
                        //Invoice Header Info
                        //****

                        myInvoice = new Invoice();
                        myInvoice.InvoiceNumber = myEntity["new_pecsinvoiceref"].ToString(); // will have a value enforced in if above

                        if (myEntity.Contains("new_customernumber"))
                            myInvoice.DebtorAccountNumber = myEntity["new_customernumber"].ToString();

                        if (myEntity.Contains("new_invoicedate"))
                            myInvoice.InvoiceDate = DateTime.Parse(myEntity["new_invoicedate"].ToString());

                        if (myEntity.Contains("new_customerrefinvoice"))
                            myInvoice.CustomerReference = myEntity["new_customerrefinvoice"].ToString();

                        // if (myEntity.Contains("name"))
                        //   myInvoice.JobNumber = myEntity["name"].ToString();

                        if (myEntity.Contains("new_salesperson"))
                            myInvoice.Consultant = GetOptionSetValueLabel(service, myEntity, "new_salesperson", (OptionSetValue)myEntity["new_salesperson"]);
                        else
                            myInvoice.Consultant = "";

                        if (myEntity.Contains("new_processedby"))
                            myInvoice.ShippingMethod = GetOptionSetValueLabel(service, myEntity, "new_processedby", (OptionSetValue)myEntity["new_processedby"]);

                        // if (myEntity.Contains("name"))
                        //  myInvoice.DeliveryLine2 = myEntity["name"].ToString();

                        //     if (myEntity.Contains("new_specialinstructionsnote1"))
                        //       myInvoice.SpecialInstructions1 = myEntity["new_specialinstructionsnote1"].ToString();

                        // if (myEntity.Contains("new_titlespecialinstructionsnote2"))
                        //   myInvoice.SpecialInstructions2 = myEntity["new_titlespecialinstructionsnote2"].ToString();

                        //*****
                        //Invoice Product Info
                        //*****

                        QueryExpression myInvoiceProductQuery = new QueryExpression
                        {
                            EntityName = "invoicedetail",
                            ColumnSet = new ColumnSet("productid", "new_note1", "new_note2", "new_note3", "new_note4", "new_note5", "new_note6", "new_note7", "new_earnedby_invoice", "new_jobnumberinvoiceproduct", "new_sigaccnumber_invoice", "quantity", "priceperunit"),
                            Criteria = new FilterExpression
                            {
                                Conditions = 
                            {
                                new ConditionExpression
                                {
                                  AttributeName  =  "invoiceid",
                                  Operator = ConditionOperator.Equal,
                                  Values = {myEntity["invoiceid"]}
                                }
                            }
                            }
                        };

                        DataCollection<Entity> productEntities = service.RetrieveMultiple(myInvoiceProductQuery).Entities;
                        if (productEntities.Count > 0)
                        {
                            var lInvoiceDetail = new List<Invoice_detail>();
                            Invoice_detail myInvoiceDetail;

                            foreach (var myProductEntity in productEntities)
                            {
                                myInvoiceDetail = new Invoice_detail();
                                string productnumber = "";
                                string sigaref = "";



                                if (myProductEntity.Contains("productid"))
                                {
                                    GetProductInfoNumberFromId(((Microsoft.Xrm.Sdk.EntityReference)myProductEntity["productid"]).Id, ref service, ref productnumber, ref sigaref);
                                    myInvoiceDetail.ProductCode = productnumber;
                                }
                                else
                                    myInvoiceDetail.ProductCode = "\"";     //handle empty values as LINQtoCSV see this as one field                            

                                if (myProductEntity.Contains("new_note1"))
                                    myInvoiceDetail.TextNarrativeLine1 = QuoteField(myProductEntity["new_note1"].ToString());
                                else
                                    myInvoiceDetail.TextNarrativeLine1 = "\"";

                                if (myProductEntity.Contains("new_note2"))
                                    myInvoiceDetail.TextNarrativeLine2 = QuoteField(myProductEntity["new_note2"].ToString());
                                else
                                    myInvoiceDetail.TextNarrativeLine2 = "\"";

                                if (myProductEntity.Contains("new_note3"))
                                    myInvoiceDetail.TextNarrativeLine3 = QuoteField(myProductEntity["new_note3"].ToString());
                                else
                                    myInvoiceDetail.TextNarrativeLine3 = "\"";

                                if (myProductEntity.Contains("new_note4"))
                                    myInvoiceDetail.TextNarrativeLine4 = QuoteField(myProductEntity["new_note4"].ToString());
                                else
                                    myInvoiceDetail.TextNarrativeLine4 = "\"";

                                if (myProductEntity.Contains("new_note5"))
                                    myInvoiceDetail.TextNarrativeLine5 = QuoteField(myProductEntity["new_note5"].ToString());
                                else
                                    myInvoiceDetail.TextNarrativeLine5 = "\"";

                                if (myProductEntity.Contains("new_note6"))
                                    myInvoiceDetail.TextNarrativeLine6 = QuoteField(myProductEntity["new_note6"].ToString());
                                else
                                    myInvoiceDetail.TextNarrativeLine6 = "\"";

                                if (myProductEntity.Contains("new_note7"))
                                    myInvoiceDetail.TextNarrativeLine7 = QuoteField(myProductEntity["new_note7"].ToString());
                                else
                                    myInvoiceDetail.TextNarrativeLine7 = "\"";

                                if (myProductEntity.Contains("new_earnedby_invoice") && myProductEntity.Contains("new_jobnumberinvoiceproduct"))
                                {
                                    string _earnby = GetOptionSetValueLabel(service, myProductEntity, "new_earnedby_invoice", (OptionSetValue)myProductEntity["new_earnedby_invoice"]).PadRight(4,' ');
                                    _earnby += myProductEntity["new_jobnumberinvoiceproduct"].ToString().PadRight(4, ' ');
                                    myInvoiceDetail.JobNumber_EarnBy = _earnby.Substring(0,8);
                                }
                                else
                                    myInvoiceDetail.JobNumber_EarnBy = "\"";

                                if (myProductEntity.Contains("new_soldby_orderproductdetail"))
                                    myInvoiceDetail.SerialNumber_SoldBy = GetOptionSetValueLabel(service, myProductEntity, "new_soldby_orderproductdetail", (OptionSetValue)myProductEntity["new_soldby_orderproductdetail"]);
                                else
                                    myInvoiceDetail.SerialNumber_SoldBy = "\"";

                                if (myProductEntity.Contains("new_sigaccnumber_invoice"))
                                    myInvoiceDetail.NomCostCenter_GLCC = QuoteField(myProductEntity["new_sigaccnumber_invoice"].ToString());
                                else
                                    myInvoiceDetail.NomCostCenter_GLCC = "\"";

                                if (myProductEntity.Contains("new_sigaccref"))
                                    myInvoiceDetail.GLAC = QuoteField(sigaref);
                                else
                                    myInvoiceDetail.GLAC = "\"";
                                if (myProductEntity.Contains("quantity"))
                                    myInvoiceDetail.Quantity = decimal.Parse(myProductEntity["quantity"].ToString());
                                else
                                    myInvoiceDetail.Quantity = 0.00m;

                                if (myProductEntity.Contains("priceperunit"))
                                    myInvoiceDetail.UnitPrice = decimal.Parse(((Microsoft.Xrm.Sdk.Money)myProductEntity["priceperunit"]).Value.ToString());
                                else
                                    myInvoiceDetail.UnitPrice = 0.00m;

                                lInvoiceDetail.Add(myInvoiceDetail);
                            }
                            myInvoice.Invoice_Product = lInvoiceDetail;
                            myInvoice.Product_String = GetProductString(lInvoiceDetail);
                        }

                        lInvoices.Add(myInvoice);
                    }
                }
                if (iCounter == 0) //nothing to export;
                    ActiveBackGroundWorker.ReportProgress(0, "No invoices to export.");
                return (lInvoices);

            }
        }
        
        private void GetProductInfoNumberFromId(Guid productid, ref IOrganizationService service, ref string productnumber, ref string new_sigaccref)
        {
            QueryExpression myProductIDQuery = new QueryExpression
            {
                EntityName = "product",
                ColumnSet = new ColumnSet("productnumber"),
                Criteria = new FilterExpression
                {
                    Conditions = 
                            {
                                new ConditionExpression
                                {
                                  AttributeName  =  "productid",
                                  Operator = ConditionOperator.Equal,
                                  Values = {productid}
                                }
                            }
                }
            };


            var entity = service.Retrieve("product", productid, new ColumnSet("productnumber"));
            productnumber = entity.GetAttributeValue<string>("productnumber");
            new_sigaccref = entity.GetAttributeValue<string>("new_sigaccref");
        }

        private string QuoteField(string Attribute)
        {
            return "*quote*" + Attribute + "*quote*";
        }

        private string GetProductString(List<Invoice_detail> lDetailToBeConverted)
        {
            bool bHasRecords = false;
            string sReturn = "*remove*";
            foreach (var DetailToBeConverted in lDetailToBeConverted)
            {
                bHasRecords = true;
                if (DetailToBeConverted.ProductCode != null)
                    sReturn += DetailToBeConverted.ProductCode;
                else
                    sReturn += "\"";
                sReturn += ",";
              
                if (DetailToBeConverted.TextNarrativeLine1 != null)
                    sReturn += DetailToBeConverted.TextNarrativeLine1;
                sReturn += ",";

                if (DetailToBeConverted.TextNarrativeLine2 != null)
                    sReturn += DetailToBeConverted.TextNarrativeLine2;
                else
                    sReturn += "\"";
                sReturn += ",";

                if (DetailToBeConverted.TextNarrativeLine3 != null)
                    sReturn += DetailToBeConverted.TextNarrativeLine3;
                else
                    sReturn += "\"";
                sReturn += ",";

                if (DetailToBeConverted.TextNarrativeLine4 != null)
                    sReturn += DetailToBeConverted.TextNarrativeLine4;
                else
                    sReturn += "\"";
                sReturn += ",";

                if (DetailToBeConverted.TextNarrativeLine5 != null)
                    sReturn += DetailToBeConverted.TextNarrativeLine5;
                else
                    sReturn += "\"";
                sReturn += ",";

                if (DetailToBeConverted.TextNarrativeLine6 != null)
                    sReturn += DetailToBeConverted.TextNarrativeLine6;
                else
                    sReturn += "\"";
                sReturn += ",";

                if (DetailToBeConverted.TextNarrativeLine7 != null)
                    sReturn += DetailToBeConverted.TextNarrativeLine7;
                else
                    sReturn += "\"";
                sReturn += ",";

                if (DetailToBeConverted.JobNumber_EarnBy != null)
                    sReturn += DetailToBeConverted.JobNumber_EarnBy;
                else
                    sReturn += "\"";
                sReturn += ",";

                if (DetailToBeConverted.SerialNumber_SoldBy != null)
                    sReturn += DetailToBeConverted.SerialNumber_SoldBy;
                else
                    sReturn += "\"";
                sReturn += ",";

                if (DetailToBeConverted.NomCostCenter_GLCC != null)
                    sReturn += DetailToBeConverted.NomCostCenter_GLCC;
                else
                    sReturn += "\"";
                sReturn += ",";
                if (DetailToBeConverted.Quantity != null)
                    sReturn += DetailToBeConverted.Quantity;
                else
                    sReturn += "\"";
                sReturn += ",";

                if (DetailToBeConverted.UnitPrice != null)
                    sReturn += DetailToBeConverted.UnitPrice;
                else
                    sReturn += "\"";
                sReturn += ",";

            }
            if (bHasRecords)
                sReturn = sReturn.Remove(sReturn.Length - 1, 1);
            sReturn += "*remove*";

            return sReturn;
        }

        private string GetOptionSetValueLabel(IOrganizationService service, Entity entity, string attribute, OptionSetValue option)
        {
            string optionLabel = String.Empty;
            var attributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entity.LogicalName,
                LogicalName = attribute,
                RetrieveAsIfPublished = true
            };

            RetrieveAttributeResponse attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);
            AttributeMetadata attrMetadata = (AttributeMetadata)attributeResponse.AttributeMetadata;
            PicklistAttributeMetadata picklistMetadata = (PicklistAttributeMetadata)attrMetadata;

            // For every status code value within all of our status codes values
            //  (all of the values in the drop down list)
            foreach (OptionMetadata optionMeta in
                picklistMetadata.OptionSet.Options)
            {
                // Check to see if our current value matches
                if (optionMeta.Value == option.Value)
                {
                    // If our numeric value matches, set the string to our status code
                    //  label
                    optionLabel = optionMeta.Label.UserLocalizedLabel.Label;
                }
            }

            return optionLabel;
        }
    }
}

