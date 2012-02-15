using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
namespace SolExcelCrmSync.Classes
{
    public static class Loging
    {
        public static string addTransaction(string sDateStamp,string sSourceDir, string sSourceFile, int iUpdated, int iAdded, string sDescription)
        {
            string sReturn = "-1";
            using (SqlConnection oConn = new SqlConnection(Connect.sConStr))
            {
                oConn.Open();
                
                string sSql = "INSERT into sync_transactions (datetimestamp, sourcedir, sourcefile, updated, added, description) ";
                sSql += " VALUES ";
                sSql += "(";
                sSql += "'" + sDateStamp + "'";
                sSql += ",'" + sSourceDir + "'";
                sSql += ",'" + sSourceFile + "'";
                sSql += "," + iUpdated;
                sSql += "," + iAdded;
                sSql += ",'" + sDescription + "'";
                sSql += ")";
                new SqlCommand(sSql, oConn).ExecuteNonQuery();

                sSql = "SELECT id from sync_transactions where datetimestamp = '" + sDateStamp + "'";
                sReturn = new SqlCommand(sSql, oConn).ExecuteScalar().ToString();
                oConn.Close();
            }
            return sReturn;
        }

        public static void addNewCustomers(string sTransactionId, List<string> lNewCustomers)
        {
            if (lNewCustomers.Count > 0)
            {
                string sSql = "";
                using (SqlConnection oConn = new SqlConnection(Connect.sConStr))
                {
                    oConn.Open();
                    foreach (string sCustomerNumber in lNewCustomers)
                    {
                        sSql = "INSERT into sync_newcustomers (fki_sync_transaction, customernumber) ";
                        sSql += " VALUES ";
                        sSql += "(";
                        sSql += sTransactionId;
                        sSql += ",'" + sCustomerNumber + "'";
                        sSql += ")";
                        new SqlCommand(sSql, oConn).ExecuteNonQuery();
                    }
                    oConn.Close();
                }
            }
        }
    }
}
