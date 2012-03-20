using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Configuration;
using SolExcelCrmSync.Containers.Account;


namespace SolExcelCrmSync.Classes
{
    class ExcelRead
    {
        public Excel.Application objExcelApp;
        public Excel._Workbook objBook;
        Excel.Workbooks objBooks;
        Excel.Sheets objSheets;
        Excel._Worksheet objSheet;
        Excel.Range rngLast;

        public List<AccountExcel> readExcel(string sExcelPath)
        {
            var lReturn = new List<AccountExcel>();
            string valueString = string.Empty;
            objExcelApp = new Microsoft.Office.Interop.Excel.Application();
            objBooks = (Excel.Workbooks)objExcelApp.Workbooks;
            //Open the workbook containing the address data.
            objBook = objBooks.Open(sExcelPath, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
            //Get a reference to the first sheet of the workbook.
            objSheets = objBook.Worksheets;
            objSheet = (Excel._Worksheet)objSheets.get_Item(1);
            
            rngLast = objSheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
            long lLastRow = rngLast.Row;
            long lLastCol = rngLast.Column;
            
            for (long rowCounter = 2; rowCounter <= lLastRow; rowCounter++) //FirstRow Has Headers - start at row 2
            {
                if (ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 1]).Value) != "")
                {
                    var adAccount = new AccountExcel();

                    adAccount.sCustomerNumber = ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 1]).Value);
                    adAccount.sAccountName = ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 40]).Value);
                    adAccount.sAddressLine1 = ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 2]).Value);
                    adAccount.sAddressLine2 = ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 5]).Value);
                    adAccount.sAddressLine3 = ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 9]).Value);
                    adAccount.sPostCode = ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 15]).Value);
                    adAccount.sTelephone = ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 17]).Value);
                    adAccount.sVatNumber = ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 18]).Value);
                    adAccount.sCountryCode = ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 21]).Value);
                    adAccount.sEmail = ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 37]).Value);
                    adAccount.sWeb = "";// ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 38]).Value);
                    adAccount.sKAM = ToStringHandlesNulls(((Excel.Range)objSheet.Cells[rowCounter, 31]).Value);
                    lReturn.Add(adAccount);
                }
            }
          //Close the Excel Object
            objBook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

            objBooks.Close();
            objExcelApp.Quit();

            Marshal.ReleaseComObject(objSheet);
            Marshal.ReleaseComObject(objSheets);
            Marshal.ReleaseComObject(objBooks);
            Marshal.ReleaseComObject(objBook);
            Marshal.ReleaseComObject(objExcelApp);

            objSheet = null;
            objSheets = null;
            objBooks = null;
            objBook = null;
            objExcelApp = null;

            GC.GetTotalMemory(false);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.GetTotalMemory(true);              
            return (lReturn);
        }

        private string ToStringHandlesNulls(object oValue)
        {
            if (oValue == null)
                return "";
            else
                return oValue.ToString();
        }

        
    }
}


