using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using System.Configuration;
using System.IO;

namespace BTPNS.Scheduler
{
    public class SharePointHelper
    {
        public void UploadFileToDocLib(string SelectedfilePath, string DocLib)
        {
            try
            {
                SPSite Site = new SPSite(ConfigurationManager.AppSettings["SharePointOnPremURL"].ToString());
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPWeb web = Site.OpenWeb())
                    {
                        web.AllowUnsafeUpdates = true;
                        if (!System.IO.File.Exists(SelectedfilePath))
                            throw new FileNotFoundException("File not found.", SelectedfilePath);
                        SPFolder libFolder = web.Folders[DocLib];

                        // Prepare to upload
                        string fileName = System.IO.Path.GetFileName(SelectedfilePath);
                        FileStream fileStream = File.OpenRead(SelectedfilePath);

                        //Check the existing File out if the Library Requires CheckOut
                        if (libFolder.RequiresCheckout)
                        {
                            try
                            {
                                SPFile fileOld = libFolder.Files[fileName];
                                fileOld.CheckOut();
                            }
                            catch { }
                        }

                        // Upload document
                        SPFile spfile = libFolder.Files.Add(fileName, fileStream, true);
                        libFolder.Update();
                        try
                        {
                            fileStream.Close();
                        }
                        catch (Exception) { }
                        web.AllowUnsafeUpdates = false;

                    }
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

        }
    }
}
