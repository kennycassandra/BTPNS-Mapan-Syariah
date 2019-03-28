using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;

namespace BTPNS.Scheduler
{
    public class CleansingHelper
    {
        public void CleansingSPFiles()
        {
            try
            {

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        public int GetCleansingDays()
        {
            try
            {
                return Convert.ToInt32(ConfigurationManager.AppSettings["CleansingDays"].ToString());
            }
            catch
            {
                return 8;
            }
        }
        public void CleansingLocalFiles(string OutputFolder)
        {
            try
            {
                string[] folders = Directory.GetDirectories(OutputFolder + "Output");
                foreach (string fol in folders)
                {
                    string[] filePaths = Directory.GetFiles(fol);
                    foreach (string f in filePaths)
                    {
                        DateTime dtModified = File.GetLastWriteTime(f);
                        DateTime dtCurrent = DateTime.Now;

                        if ((dtCurrent - dtModified).TotalDays >= GetCleansingDays())
                        {
                            File.Delete(f);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "CleansingLocalFiles");
            }

        }
    }
}
