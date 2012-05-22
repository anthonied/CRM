using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SolExcelCrmSync.Containers.Account;
using SolExcelCrmSync.Containers.User;
using Microsoft.Xrm.Sdk;


namespace SolExcelCrmSync.Containers.Tasks
{
    public class TaskBase:BaseContainer
    {
        public string sUserId { get; set; }
        public string sClientId { get; set; }
        public string sClientNumber { get; set; }
        public string sDuedate { get; set; }
        public string sMessage { get; set; }

        
        public void addTaskToUser(Marketer myMarketer, AccountBase myAccount, IOrganizationService myCRMService)
        {

            //myMarketer.sCrmEntityName = "user";
            //myMarketer.sCrmEntityGuidFieldName = "systemuserid";
            //myMarketer.sCrmAttributeName = "domainname";
            myMarketer.sCrmFilterAttributeValue = myMarketer.sCrmUsername;
            myMarketer.FindGuidForObject(myCRMService);

            //myAccount.sCrmEntityName = "account";
            //myAccount.sCrmEntityName = "accountid";
            //myAccount.sCrmAttributeName = "accountnumber";
            myAccount.sCrmFilterAttributeValue = myAccount.sCustomerNumber;
            myAccount.FindGuidForObject(myCRMService);

            var newTask = new Entity("task");
            newTask["ownerid"] = myMarketer.crmGuidId;
            newTask["regardingobjectid"] = myAccount.crmGuidId;
            newTask["description"] = sMessage;
            newTask["scheduledend"] = sDuedate;
            newTask["subject"] = myAccount.sAccountName;
            myCRMService.Create(newTask);

        }
    }
}
