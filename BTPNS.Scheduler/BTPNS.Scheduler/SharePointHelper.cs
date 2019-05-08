using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using System.Configuration;
using System.IO;
using Microsoft.SharePoint.Client;
using System.Security;

namespace BTPNS.Scheduler
{
    public class SharePointHelper
    {
        //https://social.msdn.microsoft.com/Forums/sharepoint/en-US/465c2a18-e63b-43b1-bed2-b1bf1934f0d1/need-to-get-actual-created-and-modified-time-of-the-document-while-uploading-file?forum=sharepointdevelopmentlegacy
        public ClientContext Auth(string OutputFolder, String uname, String pwd, string siteURL)
        {
            ClientContext context = new ClientContext(siteURL);
            Web web = context.Web;
            SecureString passWord = new SecureString();
            foreach (char c in pwd.ToCharArray()) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(uname, passWord);
            try
            {
                context.Load(web);
                context.ExecuteQuery();
                Console.WriteLine("Olla! from " + web.Title + " site");
                return context;
            }
            catch (Exception ex)
            {
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "Authentication SP On Prem");
                return null;
            }
        }
        public string UploadFileToDocLibClientSide(string OutputFolder, string SelectedfilePath, string DocLib, ClientContext context)
        {
            try
            {
                string fileUrl = "";
                string UrlSPOnPrem = ConfigurationManager.AppSettings["SharePointOnPremURL"].ToString();
                using (var fs = new FileStream(SelectedfilePath, FileMode.Open))
                {
                    var fi = new FileInfo(SelectedfilePath);
                    var list = context.Web.Lists.GetByTitle(DocLib);
                    context.Load(list.RootFolder);
                    context.ExecuteQuery();
                    fileUrl = String.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, fi.Name);

                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(context, fileUrl, fs, true);
                }
                return UrlSPOnPrem + fileUrl;
            }
            catch (Exception ex)
            {
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "UploadFileToDocLib - " + DocLib);
                return "";
            }
        }

        public bool ListExists(SPWeb web, string listName)
        {
            return web.Lists.Cast<SPList>().Any(list => string.Equals(list.Title, listName));
        }

        public bool ListExistsNew(ClientContext cl, string listName)
        {
            ListCollection listCollection = cl.Web.Lists;
            cl.Load(listCollection, lists => lists.Include(list => list.Title).Where(list => list.Title == listName));
            cl.ExecuteQuery();

            if (listCollection.Count > 0)
            {
                Console.WriteLine("List " + listName + " already exists");
                return true;
            }
            else
            {
                Console.WriteLine("List " + listName + " not exists");
                return false;
            }
        }

        public void CreateDocLib2(string OutputFolder, string DocLib, string Desc, ClientContext cl)
        {
            try
            {
                if (!ListExistsNew(cl, DocLib))
                {
                    using (ClientContext clientCTX = cl)
                    {
                        ListCreationInformation lci = new ListCreationInformation();
                        lci.Description = Desc;
                        lci.Title = DocLib;
                        lci.TemplateType = 101;
                        List newLib = clientCTX.Web.Lists.Add(lci);
                        clientCTX.Load(newLib);
                        clientCTX.ExecuteQuery();
                    }
                }

            }
            catch (Exception ex)
            {
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "CreateDocLib2 - " + DocLib);
            }
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

        public void CleansingFiles(string OutputFolder, string DocLib)
        {
            try
            {
                int CleansingDays = new CleansingHelper().GetCleansingDays();
                SPSite Site = new SPSite(ConfigurationManager.AppSettings["SharePointOnPremURL"].ToString());
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPWeb web = Site.OpenWeb())
                    {
                        web.AllowUnsafeUpdates = true;
                        SPFolder libFolder = web.Folders[DocLib];
                        SPFileCollection file = libFolder.Files;

                        foreach (SPFile f in file)
                        {
                            SPListItem item = f.Item;
                            int diff_days = (DateTime.Now - f.TimeCreated).Days;
                            if (diff_days > CleansingDays) item.Delete();
                        }
                        web.AllowUnsafeUpdates = false;

                    }
                });
            }
            catch (Exception ex)
            {
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "CleansingFiles SharePoint - " + DocLib);
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
                        FileStream fileStream = System.IO.File.OpenRead(SelectedfilePath);

                        //Check the existing File out if the Library Requires CheckOut
                        if (libFolder.RequiresCheckout)
                        {
                            try
                            {
                                SPFile fileOld = libFolder.Files[fileName];
                                url = fileOld.Url;
                                fileOld.CheckOut();
                            }
                            catch
                            {
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
