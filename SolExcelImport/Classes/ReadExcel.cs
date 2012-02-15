using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;


namespace SolExcelImport.Classes
{
    class ReadExcel
    {

        public Excel.Application objExcelApp;

        public Excel._Workbook objBook;
        Excel.Workbooks objBooks;
        Excel.Sheets objSheets;
        Excel._Worksheet objSheet;
        Excel.Range rngLast;

        public void readExcel()
        {
            string valueString = string.Empty;
            objExcelApp = new Microsoft.Office.Interop.Excel.Application();
            objBooks = (Excel.Workbooks)objExcelApp.Workbooks;
            //Open the workbook containing the address data.
            objBook = objBooks.Open(@"C:\Temp\data\Test.xlsx", Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
            //Get a reference to the first sheet of the workbook.
            objSheets = objBook.Worksheets;
            objSheet = (Excel._Worksheet)objSheets.get_Item(1);

            //Select the range of data containing the addresses and get the outer boundaries.
            rngLast = objSheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
            long lLastRow = rngLast.Row;
            long lLastCol = rngLast.Column;

            // Iterate through the data and concatenate the values into a comma-delimited string.
            for (long rowCounter = 1; rowCounter <= lLastRow; rowCounter++)
            {
                for (long colCounter = 1; colCounter <= lLastCol; colCounter++)
                {
                    //Write the next value into the string.
                    Excel.Range cell = (Excel.Range)objSheet.Cells[rowCounter, colCounter];
                    string cellvalue = cell.Value.ToString();
                    //TODO: add your business logic for retrieve cell value
                }
            }
        }
    }
}
   

