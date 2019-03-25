using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace BTPNS.Scheduler
{
    public class CleansingHelper
    {
        public void CleansingFiles()
        {

            string[] folders = Directory.GetDirectories(Path.Combine(Environment.CurrentDirectory, "Output"));
            foreach(string fol in folders)
            {
                string[] filePaths = Directory.GetFiles(fol);
                foreach(string f in filePaths)
                {
                    DateTime dtModified = File.GetLastWriteTime(f);
                    DateTime dtCurrent = DateTime.Now;

                    if ((dtCurrent - dtModified).TotalDays >= 8)
                    {
                        File.Delete(f);
                    }
                }
            }
        }
    }
}
