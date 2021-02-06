using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automated_Maintenance_Reminder
{
    class Program
    {
        // Excel declarations for the App, Workbook, Worksheet and Range
        public static Excel.Application xlApp; 
        public static Excel.Workbook xlWorkbook;
        public static Excel.Worksheet xlWorksheet1;
        public static Excel.Worksheet xlWorksheet2;
        public static Excel.Range xlRange1;
        public static Excel.Range xlRange2;

        public static int rowCount_Sheet1;                                      //Number of rows in used range for worksheet 1
        public static int rowCount_Sheet2;                                      //Number of rows in used range for worksheet 2
        public int daysElapsed;                                                 //Days since last maintenance
        public int daysLeft;                                                    //Days until next maintenance
        public string serialNumber;                                             //Fixture serial number
        public string partNumber;                                               //Fixture part number
        public HashSet<string> serialNumbersList = new HashSet<string>();       //Set for unique serial numbers
        public HashSet<DateTime> maintained_Dates = new HashSet<DateTime>();    //Set for unique previous maintenance dates
        public DateTime latestDate;                                             //Last maintenance date
        public DateTime nextMaintenance;                                        //Next maintenance date

        static void Main(string[] args)
        {
            //When the Excel file is in it's location with correct file name
            
            string filePath = @"\\mk-fs01\vol1\Test Eng\Documents\Test Fixture Maintenance\Test Fixture Maintenance Log.xls";
            Emailer email = new Emailer();
            FixturesChecker Tracker = new FixturesChecker();

            xlApp = new Excel.Application();

            if (File.Exists(filePath))                                          //Check if the Log file exists and is in correct extension
            {
                //Initializations
                xlWorkbook = xlApp.Workbooks.Open(filePath);
                xlWorksheet1 = xlWorkbook.Sheets[1];
                xlWorksheet2 = xlWorkbook.Sheets[2];
                xlRange1 = xlWorksheet1.UsedRange;
                xlRange2 = xlWorksheet2.UsedRange;

                rowCount_Sheet1 = xlRange1.Rows.Count;
                rowCount_Sheet2 = xlRange2.Rows.Count;

                string fullEmailToSend = "";
                string BD_Flags = "";
                string HST_Flags = "";

                fullEmailToSend += Tracker.HeaderInfo();                            //Email header information
                fullEmailToSend += Tracker.FixtureChecker("Breakdown Fixtures");    //Run object for Breakdown fixtures
                BD_Flags = Tracker.getFlag();                                       //Resultant flags from Breakdown fixtures object
                fullEmailToSend += Tracker.FixtureChecker("HST Fixtures");          //Run object for HST fixtures
                HST_Flags = Tracker.getFlag();                                      //Resultant flags from HST fixtures object

                email.setEmailText(fullEmailToSend);                                //Email object setter


                /* If Breakdown or HST fixtures contain a yellow or red fixture, set that as the email frequency.
                 * Otherwise, if none are red or yellow, set the email frequency to green 
                 */
                if (BD_Flags == "redFlag" || BD_Flags == "yellowFlag")
                {
                    email.emailFrequency(Tracker.getFlag());
                }
                else if (HST_Flags == "redFlag" || HST_Flags == "yellowFlag")
                {
                    email.emailFrequency(Tracker.getFlag());
                }
                else if (BD_Flags != "redFlag" && BD_Flags != "yellowFlag" && HST_Flags != "redFlag" && HST_Flags != "yellowFlag")
                {
                    email.emailFrequency("greenFlag");
                }

                /* Cleaning up unused memory space. Release and close all excel worksheets and application objects */
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(xlRange1);
                Marshal.ReleaseComObject(xlRange2);
                xlWorkbook.Close(0);
                Marshal.ReleaseComObject(xlWorkbook);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            else
            {
                //If Excel object cannot locate or read the Excel log file, send an error email.
                string errorEmail = "";

                errorEmail = Tracker.FileMissing();
                email.setEmailText(errorEmail);
                email.emailFrequency("redFlag");
            }
        }
    }
}
