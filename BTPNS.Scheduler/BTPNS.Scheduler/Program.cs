using System;
using System.Collections.Generic;
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
            //new MailHelper().email_send();
            //new SharePointHelper().UploadFileToDocLib("C:\\Log2.txt");
            new RDLCHelper().GeneratePDF();
            new RDLCHelper().GenerateSMS();

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

        }

    }
}
