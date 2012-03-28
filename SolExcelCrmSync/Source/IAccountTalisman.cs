using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SolExcelCrmSync.Containers.Account;
using SolExcelCrmSync.Classes;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace SolExcelCrmSync.Source
{
    public class IAccountTalisman
    {
        public List< AccountBase> FillAccountBaseFromSql()
        {
            var Accounts= new List<AccountBase>();

            using(var myConnection = new SqlConnection(Connect.sTalismanConStr))
            {
                myConnection.Open();
                string sSql = "Select top 1 * from existing_clients";
                var Reader = new SqlCommand(sSql, myConnection).ExecuteReader();
                while (Reader.Read())
                {
                    var AccountBase = new AccountBase();
                    AccountBase.sAccountName = Reader["sName"].ToString();
                    AccountBase.sAddressLine1 = Reader["sAddressLine1"].ToString();
                    AccountBase.sAddressLine2 = Reader["sAddressLine2"].ToString();
                    AccountBase.sAddressLine3 = Reader["sAddressLine3"].ToString();
                    AccountBase.sCustomerNumber = Reader["sClientNumber"].ToString();
                    AccountBase.sEmail = Reader["sEmail"].ToString();
                    AccountBase.sPostCode = Reader["sPostalCode"].ToString();
                    AccountBase.sTelephone = Reader["sTelephone"].ToString();
                    //MessageBox.Show(AccountBase.sAccountName);
                    //MessageBox.Show(AccountBase.sTelephone);
                    Accounts.Add(AccountBase);
                }
                Reader.Close();
                myConnection.Close();
            }


           
            return Accounts;
        }

    }
}
