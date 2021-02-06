using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automated_Maintenance_Reminder
{
    class FixturesChecker : Program
    {
        private bool yellowFlag = false;
        private bool redFlag = false;

        //Flags getter
        public string getFlag()
        {
            if (redFlag == true)
                return "redFlag";
            else if (yellowFlag == true)
                return "yellowFlag";

            return "greenFlag";
        }

        //Header information in the email
        public string HeaderInfo()
        {
            string headerText = "";

            headerText = "This is an automated email reminder and tracker for test fixtures maintenance.<br/>" +
                "Please follow WI-718 for the 16-week checks.<br/>For any questions, please contact Amr Elshenawy or Peter Abercrombie.<br/>" +
                "<br/>Please do NOT reply to this email as it is not monitored.<br/><br/>";

            return headerText;
        }

        //Email body if Excel object can't open/read Excel Log file
        public string FileMissing()
        {
            string errorText = "";

            errorText = "This is an automated email reminder and tracker for test fixtures maintenance.<br/>" +
                "<br/><br/><b><font color=red>ERROR - can't find/access Excel file for Maintenance Log. " +
                @"Please ensure the file is G:\Test Eng\Documents\Test Fixture Maintenance\Test Fixture Maintenance Log.xls</b></font><br/><br/>";

            return errorText;
        }

        public string FixtureChecker(string type)
        {
            int rowLimit = 0;
            string emailBody = "";
            Excel.Worksheet worksheet = null;
            Excel.Range range = null;

            if (type == "Breakdown Fixtures")
            {
                emailBody +=
                    "**************************************************************" +              //Formatting
                    "<br/><div style = margin-left:150px'><b> BREAKDOWN FIXTURES </b></div>";

                rowLimit = rowCount_Sheet1;
                worksheet = xlWorksheet1;
                range = xlRange1;
            }
            else if (type == "HST Fixtures")
            {
                emailBody +=
                    "**************************************************************" +              //Formatting
                    "<br/><div style = margin-left:180px'><b> HST FIXTURES </b></div>";

                rowLimit = rowCount_Sheet2;
                worksheet = xlWorksheet2;
                range = xlRange2;
            }

            /* Scan through the entire worksheet and dump all unique serial numbers in list */
            for (int j = 4; j < rowLimit; j++)
            {
                if (worksheet.Cells[j, 1] != null && worksheet.Cells[j, 1].Value2 != null)
                {
                    serialNumbersList.Add(range.Cells[j, 5].Value2);
                }
            }

            /* Take each unique serial number from the list and scan the worksheet for 
               for all maintenance dates for that number. Add those dates to maintained_Dates list.*/
            foreach (string part in serialNumbersList)
            {
                for (int i = 4; i < rowLimit; i++)
                {
                    if (worksheet.Cells[i, 1] != null && worksheet.Cells[i, 1].Value2 != null)
                    { 
                        serialNumber = range.Cells[i, 5].Value2;
                        double d = double.Parse(range.Cells[i, 1].Value2.ToString());
                        DateTime converted = DateTime.FromOADate(d);

                        if (serialNumber == part)
                        {
                            partNumber = range.Cells[i, 4].Value2;
                            maintained_Dates.Add(converted);
                        }
                    }
                }

                latestDate = maintained_Dates.Max();                //Find the latest date amongst list of dates for that fixture
                daysElapsed = (DateTime.Now - latestDate).Days;     
                nextMaintenance = latestDate.AddDays(112);
                daysLeft = (nextMaintenance - DateTime.Now).Days;
                maintained_Dates.Clear();                           //Clear the list for new dates for the next serial number

                if (daysLeft > 30)
                {
                    emailBody +=
                    "**************************************************************" +
                    "<br/><b>Part Number: </b>" + partNumber +
                    "<br/><b>Serial Number: </b>" + part +
                    "<br/><b>Last time maintained: </b>" + latestDate.ToShortDateString() +
                    "<br/><b>Time elapsed since last maintenance: </b>" + daysElapsed + " days OR " + daysElapsed / 7 + " weeks." +
                    "<br/><b>Next Maintenance: </b>" + nextMaintenance.ToShortDateString() + " (" + daysLeft + ")" + " days left.<br/>";
                }
                else if (daysLeft > 14 && daysLeft <= 30)
                {
                    emailBody +=
                    "**************************************************************" +
                    "<font color=#d17117><b>" +
                    "<br/><color='yellow'><b>Part Number: </b>" + partNumber +
                    "<br/><b>Serial Number: </b>" + part +
                    "<br/><b>Last time maintained: </b>" + latestDate.ToShortDateString() +
                    "<br/><b>Time elapsed since last maintenance: </b>" + daysElapsed + " days OR " + daysElapsed / 7 + " weeks." +
                    "<br/><b>Next Maintenance: </b>" + nextMaintenance.ToShortDateString() + " (" + daysLeft + ")" + " days left.<br/>" +
                    "</b></font>";

                    yellowFlag = true;          //Indicate email frequency
                }
                else if (daysLeft <= 14)
                {
                    emailBody +=
                    "**************************************************************" +
                    "<font color=red><b>" +
                    "<br/><b>Part Number: </b>" + partNumber +
                    "<br /><b>Serial Number: </b>" + part +
                    "<br/><b>Last time maintained: </b>" + latestDate.ToShortDateString() +
                    "<br/><b>Time elapsed since last maintenance: </b>" + daysElapsed + " days OR " + daysElapsed / 7 + " weeks." +
                    "<br/><b>Next Maintenance: </b>" + nextMaintenance.ToShortDateString() + " (" + daysLeft + ")" + " days left.<br/>" +
                    "</b></font>";

                    redFlag = true;             //Indicate email frequency
                }
                
            }

            emailBody += "<br/><b><font color=green>===>>> TOTAL: " + serialNumbersList.Count + " " + type + "</font></b><br/>";
            //"<p style='color: green; font-size:50px; margin-left:100px'>different font and color</p>";#bd7e13#bd7e13

            serialNumbersList.Clear();  //Clear the list for new serial numbers for HST/Breakdown
            maintained_Dates.Clear();
            return emailBody;
        }
    }
}
