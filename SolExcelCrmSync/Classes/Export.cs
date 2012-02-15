using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

namespace SolExcelCrmSync.Classes
{
    public class Export
    {
        ExportImplement eiExport;

        private DateTime _LastExport;
        public DateTime LastExport            
        {
            get { return _LastExport; }
            set { _LastExport = value; }
        }

        private int _InvoicesExportCount;
        public int InvoicesExportCount
        {
            get { return _InvoicesExportCount; }
            set { _InvoicesExportCount = value; }
        }
        public Export()
        {
            eiExport = new ExportImplement();
            this.LastExport = eiExport.GetLastExport();
        }

        public void SetLastExport(DateTime MyLastExport)
        {
           eiExport.SetLastExport(MyLastExport);
           this.LastExport = MyLastExport;           
        }
    }

    public class ExportImplement
    {
        public DateTime GetLastExport()
        {
            DateTime LastExport;
            using (var oConn = new SqlConnection(Connect.sConStr))
            {
                oConn.Open();
                string sSql = "Select max(ExportDate) from invoice_export";
                var objLastDate = new SqlCommand(sSql, oConn).ExecuteScalar();
                if (objLastDate.GetType().Name == "DBNull")
                    LastExport = new DateTime(1970, 01, 01);
                else
                    LastExport = Convert.ToDateTime(objLastDate);
                oConn.Close();
            }
            return LastExport;
        }

        public void SetLastExport(DateTime LastExport)
        {
            using (var oConn = new SqlConnection(Connect.sConStr))
            {
                oConn.Open();
                string sSql = "Insert into invoice_export ";
                sSql += "(ExportDate) ";
                sSql += " VALUES ";
                sSql += " ('" + LastExport.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                new SqlCommand(sSql, oConn).ExecuteNonQuery();                
                oConn.Close();
            }
        }

    }
}
