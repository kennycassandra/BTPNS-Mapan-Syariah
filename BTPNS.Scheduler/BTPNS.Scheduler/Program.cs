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
                string OutputFolder = AppDomain.CurrentDomain.BaseDirectory;
                Console.WriteLine("SMTP : {0}", ConfigurationManager.AppSettings["SMTP"].ToString());
                Console.WriteLine("From : {0}", ConfigurationManager.AppSettings["From"].ToString());
                Console.WriteLine("Send To : {0}", ConfigurationManager.AppSettings["TestSendTo"].ToString());
                Console.WriteLine("SQL Connection String : {0}", ConfigurationManager.ConnectionStrings["cnstr"].ToString());
                //Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory);
                Console.WriteLine("Output Location : {0}", OutputFolder);
                Console.WriteLine("Last Update : {0}", File.GetLastWriteTime(OutputFolder + "BTPNS.Scheduler.exe"));
                if (TestSendEmail == "1")
                {
                    new MailHelper().TestSendEmail();
                    Console.WriteLine("Test Send email Done");
                    Console.ReadLine();
                    return;
                }

                new SharePointHelper().CreateDocLib2(OutputFolder, "AP3R", "IF Mapan Syariah Generate PDF");
                new SharePointHelper().CreateDocLib2(OutputFolder, "CIF", "IF Mapan Syariah Generate PDF");
                new SharePointHelper().CreateDocLib2(OutputFolder, "Pembiayaan", "IF Mapan Syariah Generate PDF");
                new SharePointHelper().CreateDocLib2(OutputFolder, "Persetujuan_Pembiayaan", "IF Mapan Syariah Generate PDF");
                new SharePointHelper().CreateDocLib2(OutputFolder, "SMS_Notification", "IF Mapan Syariah Generate PDF");


                new SharePointHelper().CreateDocLib(OutputFolder, "AP3R", "IF Mapan Syariah Generate PDF");
                new SharePointHelper().CreateDocLib(OutputFolder, "CIF", "IF Mapan Syariah Generate PDF");
                new SharePointHelper().CreateDocLib(OutputFolder, "Pembiayaan", "IF Mapan Syariah Generate PDF");
                new SharePointHelper().CreateDocLib(OutputFolder, "Persetujuan_Pembiayaan", "IF Mapan Syariah Generate PDF");
                new SharePointHelper().CreateDocLib(OutputFolder, "SMS_Notification", "IF Mapan Syariah Generate PDF");

                new CleansingHelper().CleansingSPFiles(OutputFolder);
                new CleansingHelper().CleansingLocalFiles(OutputFolder);

                new RDLCHelper().GeneratePDF(OutputFolder);
                new RDLCHelper().GenerateSMS(OutputFolder);
                new RDLCHelper().GenerateExcelSummaryReport_Detail1(OutputFolder);
                new RDLCHelper().GenerateExcelSummaryReport_Detail2(OutputFolder);
                new RDLCHelper().GenerateExcelDetailReport(OutputFolder);
                new RDLCHelper().GenerateExcelLogReport(OutputFolder);
                new GenerateTxt().GenerateTxtCIF(OutputFolder);
                new GenerateTxt().GenerateTxtPembiayaan(OutputFolder);
                System.Threading.Thread.Sleep(5000);
                Console.WriteLine("Process Done");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);               
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
