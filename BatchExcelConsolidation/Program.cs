using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Reflection;

namespace BatchExcelConsolidation
{
    class Program
    {
        private static Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        private static Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        private static Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        private static Microsoft.Office.Interop.Excel.Worksheet mWSheet2;
        private static Microsoft.Office.Interop.Excel.Application oXL;
        StreamWriter w = File.AppendText("log.txt");
        public const int ERRCOUNT = 0;

        public static string[] prepGeneralData(string pathName, string filename, string startingQuarter, StreamWriter w)
        {
            try
            {
                //Gathering General information about company.
                //Typical file name "Tesla Q1 Report", "Tesla Q2 Report.xlsx". A Q4 report would have all previous financial info.
                //We look for the Quarter 4 value first and work backwards if it's missing. CEO may change in between questers or location moves.
                //Destination file hard coded to "Destination./xlsx" for testing. 
                //Place source files and destination in same directory as exe and run.
                //Open the source workbook and extract the values into array and return it.
                w.WriteLine("Starting import: " + startingQuarter + " Report", w);
                string file = pathName + filename;
                Application excel = new Application();
                Workbook workbook = excel.Workbooks.Open(file);
                Worksheet sheet = workbook.Sheets[startingQuarter + " Report"];
                string[] valueArray = new string[] {Convert.ToString(sheet.Cells[2,2].Value2), //[row , col]  //ceo
                                                Convert.ToString(sheet.Cells[4,2].Value2)}; //Location
                workbook.Close(false, Missing.Value, Missing.Value);


                //For Testing: Iterate through each cell and display the contents.
                /*foreach (string i in valueArray)
                {
                    System.Console.WriteLine(i);
                }
                Console.WriteLine("Formula array built for " + pathName + filename + " completed.");
                Console.ReadLine();
                */

                return valueArray;
            }
            catch (Exception ex)
            {
                //If specific quarter info is missing, return empty array and continue import procedure
                w.WriteLine("Error encountered: " + startingQuarter + ": " + ex);
                w.WriteLine("Returning empty array for write. Check log file.");
                Console.WriteLine("error has occur for " + startingQuarter + ". The application will move on to the next quarter. Please review export file/log for missing data");
                string[] emptyValueArray = new string[] { "", "" };
                return emptyValueArray;
            }
        }

        public static string[] prepQuarterSummary(string pathName, string filename, string quarter, StreamWriter w)
        {
            try
            {
                w.WriteLine("Starting import of Quarter Data from report " + quarter, w);
                string file = pathName + filename;
                Console.WriteLine("Accessing tab : Report " + quarter);
                Application excel = new Application();
                Workbook workbook = excel.Workbooks.Open(file);
                Worksheet sheet = workbook.Sheets[quarter + " Report"];
                string[] valueArray = new string[] {
                                                //Summary
                                                Convert.ToString(sheet.Cells[2,4].Value2),
                                                Convert.ToString(sheet.Cells[3,4].Value2),
                                                Convert.ToString(sheet.Cells[4,4].Value2),
                                                Convert.ToString(sheet.Cells[5,4].Value2),

                                                //Average Financial Data
                                                Convert.ToString(sheet.Cells[1,6].Value2),
                                                Convert.ToString(sheet.Cells[2,6].Value2),
                                                Convert.ToString(sheet.Cells[3,6].Value2),
                                                Convert.ToString(sheet.Cells[4,6].Value2),
                                                Convert.ToString(sheet.Cells[5,6].Value2),
                                                };
                workbook.Close(false, Missing.Value, Missing.Value);
                Console.WriteLine("Closing file.");
                /*//For Testing: Iterate through each cell and display the contents.
                foreach (string i in valueArray)
                {
                    System.Console.WriteLine(i);
                }
                Console.WriteLine("Formula array built for " + pathName + filename + " completed.");
                */
                return valueArray;
            }
            catch (Exception ex)
            {
                w.WriteLine("Error encountered: " + quarter + ": " + ex);
                w.WriteLine("Returning empty array for write.");
                Console.WriteLine("error has occur for " + quarter + ". The application will move on to the next quarter. Please review export file/log for missing data");
                string[] emptyValueArray = new string[] { "", "", "", "", "", "", "", "", "" };
                return emptyValueArray;
            }
        }

        public static string[] prepQuarterFinancial(string pathName, string filename, string quarter, StreamWriter w)
        {
            try
            {
                w.WriteLine("Importing Quarter financial numbers from Report " + quarter , w);
                string file = pathName + filename;
                Console.WriteLine("Accessing tab : Report " + quarter);
                Application excel = new Application();
                Workbook workbook = excel.Workbooks.Open(file);
                Worksheet sheet = workbook.Sheets[quarter + " Report"];
                string[] valueArray = new string[] {
                                                //USA Month 1 Data (if you refer to Tesla 2016 Q4.xlsx)
                                                Convert.ToString(sheet.Cells[8,2].Value2),
                                                Convert.ToString(sheet.Cells[9,2].Value2),
                                                Convert.ToString(sheet.Cells[10,2].Value2),
                                                Convert.ToString(sheet.Cells[11,2].Value2),
                                                Convert.ToString(sheet.Cells[12,2].Value2),
                                                //USA Month 2 Data
                                                Convert.ToString(sheet.Cells[8,3].Value2),
                                                Convert.ToString(sheet.Cells[9,3].Value2),
                                                Convert.ToString(sheet.Cells[10,3].Value2),
                                                Convert.ToString(sheet.Cells[11,3].Value2),
                                                Convert.ToString(sheet.Cells[12,3].Value2),
                                                //USA Month 3 Data
                                                Convert.ToString(sheet.Cells[8,4].Value2),
                                                Convert.ToString(sheet.Cells[9,4].Value2),
                                                Convert.ToString(sheet.Cells[10,4].Value2),
                                                Convert.ToString(sheet.Cells[11,4].Value2),
                                                Convert.ToString(sheet.Cells[12,4].Value2),
                                                //Canada Month 1 Data
                                                Convert.ToString(sheet.Cells[14,2].Value2),
                                                Convert.ToString(sheet.Cells[15,2].Value2),
                                                Convert.ToString(sheet.Cells[16,2].Value2),
                                                Convert.ToString(sheet.Cells[17,2].Value2),
                                                Convert.ToString(sheet.Cells[18,2].Value2),
                                                //Canada Month 2 Data
                                                Convert.ToString(sheet.Cells[14,3].Value2),
                                                Convert.ToString(sheet.Cells[15,3].Value2),
                                                Convert.ToString(sheet.Cells[16,3].Value2),
                                                Convert.ToString(sheet.Cells[17,3].Value2),
                                                Convert.ToString(sheet.Cells[18,3].Value2),
                                                //Canada Month 3 Data
                                                Convert.ToString(sheet.Cells[14,4].Value2),
                                                Convert.ToString(sheet.Cells[15,4].Value2),
                                                Convert.ToString(sheet.Cells[16,4].Value2),
                                                Convert.ToString(sheet.Cells[17,4].Value2),
                                                Convert.ToString(sheet.Cells[18,4].Value2),
                                                //Europe Month 1 Data
                                                Convert.ToString(sheet.Cells[20,2].Value2),
                                                Convert.ToString(sheet.Cells[21,2].Value2),
                                                Convert.ToString(sheet.Cells[22,2].Value2),
                                                Convert.ToString(sheet.Cells[23,2].Value2),
                                                Convert.ToString(sheet.Cells[24,2].Value2),
                                                //Europe Month 2 Data
                                                Convert.ToString(sheet.Cells[20,3].Value2),
                                                Convert.ToString(sheet.Cells[21,3].Value2),
                                                Convert.ToString(sheet.Cells[22,3].Value2),
                                                Convert.ToString(sheet.Cells[23,3].Value2),
                                                Convert.ToString(sheet.Cells[24,3].Value2),
                                                //Europe Month 3 Data
                                                Convert.ToString(sheet.Cells[20,4].Value2),
                                                Convert.ToString(sheet.Cells[21,4].Value2),
                                                Convert.ToString(sheet.Cells[22,4].Value2),
                                                Convert.ToString(sheet.Cells[23,4].Value2),
                                                Convert.ToString(sheet.Cells[24,4].Value2)};
                workbook.Close(false, Missing.Value, Missing.Value);
                Console.WriteLine("Closing file.");

                /*//For Testing: Iterate through each cell and display the contents.
                foreach (string i in valueArray)
                {
                    System.Console.WriteLine(i);
                }
                Console.WriteLine("Formula array built for " + pathName + filename + " completed.");
                */
                return valueArray;
            }
            catch (Exception ex)
            {
                w.WriteLine("Error encountered: " + quarter + ": " + ex);
                w.WriteLine("Returning empty array for write.");
                Console.WriteLine("Error encountered: " + quarter + ". The application will move on to the next quarter. Please review export file/log for missing data");
                string[] emptyValueArray = new string[] { "", "", "", "", "", "", "", "", "" };
                return emptyValueArray;
            }
        }

        public static void writeGeneralDataToExcel(string[] dataSetGeneralData, string destinationPath, int Row, string startQuarter, StreamWriter w)
        {
            w.WriteLine("Write " + startQuarter + "'s General Data to Destination Excel", w);

            //This is used to write the formula array to the destination excel
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;
            oXL.DisplayAlerts = false;
            mWorkBook = oXL.Workbooks.Open(destinationPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Get all the sheets in the workbook
            mWorkSheets = mWorkBook.Worksheets;
            //Get existing sheet
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Detail");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
            mWSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Financial " + startQuarter);
            Microsoft.Office.Interop.Excel.Range range2 = mWSheet2.UsedRange;

            int startGenDataCol = 3;
            for (int index = 0; index < dataSetGeneralData.Length; index++)
            {
                mWSheet1.Cells[Row, startGenDataCol] = dataSetGeneralData[index];
                startGenDataCol++;
            }
            Console.WriteLine("Done write");
            mWorkBook.SaveAs(destinationPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public static void writeQ4ToExcel(string[] dataDetail, string[] dataFinancial, string destinationPath, int Row, int Row2, string startQuarter, StreamWriter w)
        {
            w.WriteLine("Writing Q4's Detail and Financial to Excel.", w);
            //This is used to write the source array to the destination excel
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;
            oXL.DisplayAlerts = false;
            mWorkBook = oXL.Workbooks.Open(destinationPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Get all the sheets in the workbook
            mWorkSheets = mWorkBook.Worksheets;
            //Get existing sheet
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Detail");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
            //int colCount = range.Columns.Count;
            //int rowCount = range.Rows.Count;
            mWSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Financial " + startQuarter);
            Microsoft.Office.Interop.Excel.Range range2 = mWSheet2.UsedRange;

            int startDetailColumn = 32;
            int startFinancialColumn = 3;
            int counterFinancialRow = Row2;

            for (int index = 0; index < dataDetail.Length; index++)
            {
                mWSheet1.Cells[Row, startDetailColumn] = dataDetail[index];
                startDetailColumn++;
            }

            for (int index = 0; index < dataFinancial.Length; index++)
            {
                mWSheet2.Cells[Row2, startFinancialColumn] = dataFinancial[index];
                startFinancialColumn++;
                if (startFinancialColumn == 18) //When 18 is reach move to next line. //Console.WriteLine("Writing " + dataDetail[index] + " to row " + Row + " Column " + startDetailColumn);
                {
                    Row2++;
                    startFinancialColumn = 3;
                }
            }
            mWorkBook.SaveAs(destinationPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public static void writeQ3ToExcel(string[] dataDetail, string[] dataFinancial, string destinationPath, int Row, int Row2, string startQuarter, StreamWriter w)
        {
            w.WriteLine("Writing Q3's Detail and Financial to Excel.", w);
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;
            oXL.DisplayAlerts = false;
            mWorkBook = oXL.Workbooks.Open(destinationPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            mWorkSheets = mWorkBook.Worksheets;
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Detail");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
            mWSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Financial " + startQuarter);
            Microsoft.Office.Interop.Excel.Range range2 = mWSheet2.UsedRange;

            int startDetailColumn = 23;
            int startFinancialColumn = 3;
            int counterFinancialRow = Row2;

            for (int index = 0; index < dataDetail.Length; index++)
            {
                mWSheet1.Cells[Row, startDetailColumn] = dataDetail[index];
                startDetailColumn++;
            }

            for (int index = 0; index < dataFinancial.Length; index++)
            {
                mWSheet2.Cells[Row2, startFinancialColumn] = dataFinancial[index];
                startFinancialColumn++;
                if (startFinancialColumn == 18)
                {
                    Row2++;
                    startFinancialColumn = 3;
                }
            }
            Console.WriteLine("Done Q3 write");
            mWorkBook.SaveAs(destinationPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public static void writeQ2ToExcel(string[] dataDetail, string[] dataFinancial, string destinationPath, int Row, int Row2, string startQuarter, StreamWriter w)
        {
            w.WriteLine("Writing Q2's Detail and Financial to Excel.", w);
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;
            oXL.DisplayAlerts = false;
            mWorkBook = oXL.Workbooks.Open(destinationPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            mWorkSheets = mWorkBook.Worksheets;
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Detail");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
            mWSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Financial " + startQuarter);
            Microsoft.Office.Interop.Excel.Range range2 = mWSheet2.UsedRange;
            int startDetailColumn = 14;
            int startFinancialColumn = 3;
            int counterFinancialRow = Row2;

            for (int index = 0; index < dataDetail.Length; index++)
            {
                mWSheet1.Cells[Row, startDetailColumn] = dataDetail[index];
                startDetailColumn++;
            }
            for (int index = 0; index < dataFinancial.Length; index++)
            {
                mWSheet2.Cells[Row2, startFinancialColumn] = dataFinancial[index];
                startFinancialColumn++;
                if (startFinancialColumn == 18)
                {
                    Row2++;
                    startFinancialColumn = 3;
                }
            }
            Console.WriteLine("Done Q2 write");
            mWorkBook.SaveAs(destinationPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public static void writeQ1ToExcel(string[] dataDetail, string[] dataFinancial, string destinationPath, int Row, int Row2, string startQuarter, StreamWriter w)
        {
            w.WriteLine("Writing Q1's Detail and Financial to Excel.", w);
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;
            oXL.DisplayAlerts = false;
            mWorkBook = oXL.Workbooks.Open(destinationPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            mWorkSheets = mWorkBook.Worksheets;
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Detail");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
            mWSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Financial " + startQuarter);
            Microsoft.Office.Interop.Excel.Range range2 = mWSheet2.UsedRange;
            int startDetailColumn = 5;
            int startFinancialColumn = 3;
            int counterFinancialRow = Row2;
            for (int index = 0; index < dataDetail.Length; index++)
            {
                mWSheet1.Cells[Row, startDetailColumn] = dataDetail[index];
                startDetailColumn++;
            }
            for (int index = 0; index < dataFinancial.Length; index++)
            {
                mWSheet2.Cells[Row2, startFinancialColumn] = dataFinancial[index];
                startFinancialColumn++;
                if (startFinancialColumn == 18)
                {
                    Row2++;
                    startFinancialColumn = 3;
                }
            }
            mWorkBook.SaveAs(destinationPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        static void Main()
        {
            //Prep lookup dictionaries for Detail and Financial quarter tabs.
            //(Internal ID, excelRow)
            Dictionary<int, int> DetailMap = new Dictionary<int, int>();
            DetailMap.Add(1, 3);
            DetailMap.Add(2, 4);
            Dictionary<int, int> financialMap = new Dictionary<int, int>();
            financialMap.Add(1, 3);
            financialMap.Add(2, 8);

            using (StreamWriter w = File.AppendText("log.txt"))
            {
                w.WriteLine("\r\n--Application started : {0} {1}", DateTime.Now.ToLongTimeString(),
                    DateTime.Now.ToLongDateString());
                try
                {
                    string directory = Directory.GetCurrentDirectory() + @"\";
                    Console.WriteLine("Directory : " + directory);

                    string[] dirs = Directory.GetFiles(directory);
                    string[] xlsxFiles = Directory.GetFiles(directory, "*.xlsx")
                                             .Select(path => Path.GetFileName(path))
                                             .ToArray();
                    Console.WriteLine("The number of excel files found : {0}.", xlsxFiles.Length);
                    foreach (string xls in xlsxFiles)
                    {
                        string destinationPath = directory + "Destination.xlsx";  //Change if desired
                        string sourcePath = directory;
                        string sourceFilename = xls;
                        string[] validQuarters = new[] { "Q4", "Q3", "Q2", "Q1" };
                        string startQuarter = "";
                        string id = sourcePath + sourceFilename;
                        if (sourceFilename.Contains("~$"))
                        {
                            w.WriteLine("Accessing : " + sourceFilename + "\n");
                            w.WriteLine("Temporary file found. No action performed.", w);
                        }
                        else
                        {
                            Console.WriteLine("Starting import for : " + sourcePath + sourceFilename);
                            w.WriteLine("Accessing : " + id + "\n");
                            int rowNumDetail;
                            int rowNumFinancial;
                            if (validQuarters.Any(sourceFilename.Contains))
                            {
                                Application excel = new Application();
                                Workbook workbook = excel.Workbooks.Open(id);
                                Worksheet sheet = workbook.Sheets["General"];
                                string PrgmID = Convert.ToString(sheet.Cells[2, 2].Value2);  //Internal ID (Column,Row)
                                workbook.Close(false, Missing.Value, Missing.Value);
                                int companyIDInt = Convert.ToInt32(PrgmID);   //String to Int convert
                                Console.WriteLine("ID : " + companyIDInt);
                                //Find destination row from Company ID Dictionaries
                                rowNumDetail = DetailMap[companyIDInt];  //returns excel row to write to in Destination File Detail Tab
                                rowNumFinancial = financialMap[companyIDInt];   //returns excel row to write to in Destination file Financial tab
                                if (sourceFilename.Contains("Q4"))
                                {
                                    startQuarter = "Q4";
                                    Log("Starting Quarter : " + startQuarter + "\n", w);
                                    string[] dataSetQ4GeneralData = prepGeneralData(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ4Summary = prepQuarterSummary(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ4Financial = prepQuarterFinancial(sourcePath, sourceFilename, startQuarter, w);
                                    writeGeneralDataToExcel(dataSetQ4GeneralData, destinationPath, rowNumDetail, startQuarter, w);
                                    writeQ4ToExcel(dataSetQ4Summary, dataSetQ4Financial, destinationPath, rowNumDetail, rowNumFinancial, startQuarter, w);
                                    startQuarter = "Q3";
                                    string[] dataSetQ3Summary = prepQuarterSummary(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ3Financial = prepQuarterFinancial(sourcePath, sourceFilename, startQuarter, w);
                                    writeQ3ToExcel(dataSetQ3Summary, dataSetQ3Financial, destinationPath, rowNumDetail, rowNumFinancial, startQuarter, w);
                                    startQuarter = "Q2";
                                    string[] dataSetQ2Summary = prepQuarterSummary(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ2Financial = prepQuarterFinancial(sourcePath, sourceFilename, startQuarter, w);
                                    writeQ2ToExcel(dataSetQ2Summary, dataSetQ2Financial, destinationPath, rowNumDetail, rowNumFinancial, startQuarter, w);
                                    startQuarter = "Q1";
                                    string[] dataSetQ1Summary = prepQuarterSummary(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ1Financial = prepQuarterFinancial(sourcePath, sourceFilename, startQuarter, w);
                                    writeQ1ToExcel(dataSetQ1Summary, dataSetQ1Financial, destinationPath, rowNumDetail, rowNumFinancial, startQuarter, w);

                                }
                                else if (sourceFilename.Contains("Q3"))
                                {
                                    startQuarter = "Q3";
                                    Log("Starting Quarter : " + startQuarter + "\n", w);
                                    string[] dataSetQ3GeneralData = prepGeneralData(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ3Summary = prepQuarterSummary(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ3Financial = prepQuarterFinancial(sourcePath, sourceFilename, startQuarter, w);
                                    writeGeneralDataToExcel(dataSetQ3GeneralData, destinationPath, rowNumDetail, startQuarter, w);
                                    writeQ3ToExcel(dataSetQ3Summary, dataSetQ3Financial, destinationPath, rowNumDetail, rowNumFinancial, startQuarter, w);
                                    startQuarter = "Q2";
                                    string[] dataSetQ2Summary = prepQuarterSummary(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ2Financial = prepQuarterFinancial(sourcePath, sourceFilename, startQuarter, w);
                                    writeQ2ToExcel(dataSetQ2Summary, dataSetQ2Financial, destinationPath, rowNumDetail, rowNumFinancial, startQuarter, w);
                                    startQuarter = "Q1";
                                    string[] dataSetQ1Summary = prepQuarterSummary(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ1Financial = prepQuarterFinancial(sourcePath, sourceFilename, startQuarter, w);
                                    writeQ1ToExcel(dataSetQ1Summary, dataSetQ1Financial, destinationPath, rowNumDetail, rowNumFinancial, startQuarter, w);

                                }
                                else if (sourceFilename.Contains("Q2"))
                                {
                                    startQuarter = "Q2";
                                    Log("Starting Quarter : " + startQuarter + "\n", w);
                                    string[] dataSetQ2GeneralData = prepGeneralData(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ2Summary = prepQuarterSummary(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ2Financial = prepQuarterFinancial(sourcePath, sourceFilename, startQuarter, w);
                                    writeGeneralDataToExcel(dataSetQ2GeneralData, destinationPath, rowNumDetail, startQuarter, w);
                                    writeQ2ToExcel(dataSetQ2Summary, dataSetQ2Financial, destinationPath, rowNumDetail, rowNumFinancial, startQuarter, w);
                                    startQuarter = "Q1";
                                    string[] dataSetQ1Summary = prepQuarterSummary(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ1Financial = prepQuarterFinancial(sourcePath, sourceFilename, startQuarter, w);
                                    writeQ1ToExcel(dataSetQ1Summary, dataSetQ1Financial, destinationPath, rowNumDetail, rowNumFinancial, startQuarter, w);
                                }
                                else if (sourceFilename.Contains("Q1"))
                                {
                                    startQuarter = "Q1";
                                    Log("Starting Quarter : " + startQuarter + "\n", w);
                                    string[] dataSetQ1GeneralData = prepGeneralData(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ1Summary = prepQuarterSummary(sourcePath, sourceFilename, startQuarter, w);
                                    string[] dataSetQ1Financial = prepQuarterFinancial(sourcePath, sourceFilename, startQuarter, w);
                                    writeGeneralDataToExcel(dataSetQ1GeneralData, destinationPath, rowNumDetail, startQuarter, w);
                                    writeQ1ToExcel(dataSetQ1Summary, dataSetQ1Financial, destinationPath, rowNumDetail, rowNumFinancial, startQuarter, w);
                                }
                                else { Log("Quarter specified not found. Moving on", w); }
                            }
                            else { Log("Quarter specified not found. Moving on" + "\n", w); }
                        }
                    }
                    Log("********************************************", w);
                    Console.WriteLine("\nImport completed. Press Enter to close application. Review log for any errors.");
                    Console.ReadLine();
                }
                catch (Exception e)
                {
                    Log("**Exception caught: " + e, w);
                    Console.WriteLine("An error has occur (most likely missing destination file or locating the ID). Review the log.txt file for specifics. Please Enter to close the program");
                }
            }
        }
        public static void Log(string logMessage, TextWriter w)
        {
            w.WriteLine(" :{0}", logMessage);
        }
        public static void DumpLog(StreamReader r)
        {
            string line;
            while ((line = r.ReadLine()) != null)
            {
                Console.WriteLine(line);
            }
        }
    }
}