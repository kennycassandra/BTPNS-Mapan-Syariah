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
        public bool ListExists(SPWeb web, string listName)
        {
            return web.Lists.Cast<SPList>().Any(list => string.Equals(list.Title, listName));
        }
        public void CreateDocLib(string OutputFolder, string DocLib, string Desc)
        {
            try
            {
                using (SPSite oSPsite = new SPSite(ConfigurationManager.AppSettings["SharePointOnPremURL"].ToString()))
                {

                    using (SPWeb oSPWeb = oSPsite.OpenWeb())
                    {
                        oSPWeb.AllowUnsafeUpdates = true;
                        if (!ListExists(oSPWeb, DocLib))
                        {
                            /*create list from custom ListTemplate present within ListTemplateGalery */
                            //SPListTemplateCollection lstTemp = oSPsite.GetCustomListTemplates(oSPWeb);
                            //SPListTemplate template = lstTemp["custom template name"];
                            oSPWeb.Lists.Add(DocLib, Desc, SPListTemplateType.DocumentLibrary);
                            oSPWeb.Update();
                            oSPWeb.AllowUnsafeUpdates = false;
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "CreateDocLib - " + DocLib);
            }
          
        }
        public string UploadFileToDocLib(string OutputFolder, string SelectedfilePath, string DocLib)
        {
            try
            {
                string url = "";
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
                                url = fileOld.Url;
                                fileOld.CheckOut();
                            }
                            catch {
                            }
                        }

                        // Upload document
                        SPFile spfile = libFolder.Files.Add(fileName, fileStream, true);
                        libFolder.Update();
                        url = spfile.Url;
                        try
                        {
                            fileStream.Close();
                        }
                        catch (Exception)
                        {
                        }
                        web.AllowUnsafeUpdates = false;
                    }
                });
                return url;
            }
            catch (Exception ex)
            {
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "UploadFileToDocLib - " + DocLib);
                return "";
            }

        }
    }
}
