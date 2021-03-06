using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Contrôleur_RIB
{
    class ExcelApp
    {
        private Excel.Application myApplication;
        private Excel.Workbook myWorkbook;
        private Excel.Worksheet myWorksheet;
        private String path;
        private bool isOpen = false;
        public ExcelApp(String path)
        {
            myApplication = new Excel.Application();// Starting the app
            myApplication.Visible = false;
            myWorkbook = myApplication.Workbooks.Open(path);//Opening the file
            myWorksheet = /*(Excel.Worksheet)*/myWorkbook.Sheets[1];// Selecting the first worksheet, may need a user input to determine which sheet to open later on
            MyApplication = myApplication;// Automated setters don't work for some reason, manual attribution
            MyWorkbook = myWorkbook;
            MyWorksheet = myWorksheet;
            IsOpen = true;
        }

        public List<String> GetAllRIBs()
        {
            List<String> listOfRIBs = new List<String>();
            Excel.Range excelRange = myWorksheet.UsedRange;// Range property returns how many rows and columns are used inside the Excel sheet, we use it to get the total amount of rows and columns
            int rowCount = excelRange.Rows.Count;
            const int RIBLocationColumn = 1;// We define the column to read for RIBs
            for (int i = 2; i <= rowCount; i++)// Starting i at 2 because we want to skip row 1 which contains definitions and not data (Excel starts at 1 and not 0 like lists)
            {
                Excel.Range range = (myWorksheet.Cells[i, RIBLocationColumn] as Excel.Range);// Iterating over the specific column containing RIBs on each line
                listOfRIBs.Add(range.Value.ToString());// We add the value to the list
            }
            return listOfRIBs;
        }

        public Boolean ColumnIsEmpty(int columnToAnalyse)
        {
            Boolean result = true;
            Excel.Range excelRange = myWorksheet.UsedRange;// Range property returns how many rows and columns are used inside the Excel sheet, we use it to get the total amount of rows and columns
            int rowCount = excelRange.Rows.Count;
            int colCount = excelRange.Columns.Count;
            if (colCount >= 4)
            {
                for (int i = 2; i <= rowCount; i++)
                {
                    Excel.Range range = (myWorksheet.Cells[i, columnToAnalyse] as Excel.Range);// Getting the position of the cell we are looking at
                    string cellValue = range.Value;// Storing that value into a variable
                    if (cellValue != null)
                    {
                        result = false;
                        return result;
                    }
                } 
                return result;
            }
            else
            {
            return result;
            }
        }

        public void WriteResults(List<String> results, int columnToOverwrite)
        {
            Excel.Range excelRange = myWorksheet.UsedRange;// Range property returns how many rows and columns are used inside the Excel sheet, we use it to get the total amount of rows and columns
            int rowCount = excelRange.Rows.Count;
            for (int i = 2; i <= rowCount; i++)
            {
                this.MyWorksheet.Cells[i, columnToOverwrite] = results[i-2];
                this.MyWorkbook.Save();
            }
        }

        public void Terminate()// Release the file from use when the application is shutting down
        {
            myWorkbook.Close();
            myApplication.Quit();
            IsOpen = false;
        }

        public Excel.Application MyApplication { get; set; }
        public Excel.Workbook MyWorkbook { get; set; }
        public Excel.Worksheet MyWorksheet { get; set; }
        public String Path { get; set; }

        public bool IsOpen { get; set; }
    }
}
