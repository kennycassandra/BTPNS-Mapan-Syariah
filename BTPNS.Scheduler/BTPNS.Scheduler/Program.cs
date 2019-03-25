using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BTPNS.Scheduler
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string TestSendEmail = ConfigurationManager.AppSettings["TestSendEmail"].ToString();
                Console.WriteLine("SMTP : {0}", ConfigurationManager.AppSettings["SMTP"].ToString());
                Console.WriteLine("From : {0}", ConfigurationManager.AppSettings["From"].ToString());
                Console.WriteLine("Send To : {0}", ConfigurationManager.AppSettings["TestSendTo"].ToString());
                Console.WriteLine("SQL Connection String : {0}", ConfigurationManager.ConnectionStrings["cnstr"].ToString());

                if (TestSendEmail == "1")
                {
                    new MailHelper().TestSendEmail();
                    Console.WriteLine("Test Send email Done");
                    Console.ReadLine();
                    return;
                }

                new CleansingHelper().CleansingFiles();
                new RDLCHelper().GeneratePDF();
                new RDLCHelper().GenerateSMS();
                new RDLCHelper().GenerateExcelSummaryReport_Detail1();
                new RDLCHelper().GenerateExcelSummaryReport_Detail2();
                new RDLCHelper().GenerateExcelDetailReport();
                new RDLCHelper().GenerateExcelLogReport();
                Console.WriteLine("Process Done");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadLine();
            }


            #region Matrix Grade
            //int Start_PUS = 11; int Start_PRS =21;
            //int End_PUS = 50; int End_PRS = 50;
            //int Tenor = 24;
            //int OriStartPRS = Start_PRS;
            //string Grade = "D";
            //while (Start_PUS <= End_PUS)
            //{
            //    Start_PRS = OriStartPRS;
            //    while (Start_PRS <= End_PRS)
            //    {
            //        using (StreamWriter writer = new StreamWriter("C:\\log3.txt", true))
            //        {
            //            writer.WriteLine("insert into MstGrade (Tenor, PUS, PRS, Grade) select " + Tenor.ToString() + "," + Start_PUS.ToString() + "," + Start_PRS.ToString() + ",'" + Grade + "'");
            //        }
            //        Start_PRS++;
            //    }
            //    Start_PUS++;
            //}
            #endregion
        }

    }
}
