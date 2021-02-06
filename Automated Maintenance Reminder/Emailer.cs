using System;
using System.Net.Mail;
using System.IO;

namespace Automated_Maintenance_Reminder
{
    class Emailer
    {
        private string fullEmailToSend = "";

        //Email body setter
        public void setEmailText(string fullEmailToSend)
        {
            this.fullEmailToSend = fullEmailToSend;
        }

        /* Based on the flag type, dictate the frequency emails should be sent out.
         * redFlag = daily
         * yellowFlag = twice a week
         * greenFlag = once a week
         */
        public void emailFrequency(string flag)
        {
            StreamReader dataInfo = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + @"\SentEmailsHistory.dat");    //Open email history data file
            DateTime date = Convert.ToDateTime(dataInfo.ReadLine());                                                        //Read when was the last email sent
            int period = (DateTime.Now - date).Days;

            if (flag == "redFlag" && period >= 0)
            {
                dataInfo.Close();
                SendEmail(fullEmailToSend);
            }
            else if (flag == "yellowFlag" && period >= 3)
            {
                dataInfo.Close();
                SendEmail(fullEmailToSend);
            }
            else if (flag == "greenFlag" && period > 7)
            {
                dataInfo.Close();
                if(DateTime.Now.DayOfWeek == DayOfWeek.Monday)          //If all fixtures are greenFlag and none are over due, reset email frequency to be on Monday weekly
                {
                    SendEmail(fullEmailToSend);
                }    
            }
        }

        /* Email method that is called by emailFrequency() based on the flag conditions.
         * Sets addresses, ccAddresses
         * After sending out email, write the date in the data file.
         */
        private void SendEmail(string body)
        {
            StreamWriter dataFile = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + @"\SentEmailsHistory.dat");
            string addresses = "x";
            string ccAddresses = "x";
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("x");

            mail.From = new MailAddress("x");
            mail.To.Add(addresses);
            mail.CC.Add(ccAddresses);
            mail.Subject = "Test Fixtures Reminder & Tracker";
            mail.Body = body;
            mail.IsBodyHtml = true;
            SmtpServer.Send(mail);
            dataFile.WriteLine(DateTime.Now.ToShortDateString());
            dataFile.Close();
        }
    }
}
