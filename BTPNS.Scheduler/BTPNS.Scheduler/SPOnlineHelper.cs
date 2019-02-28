using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Configuration;
namespace BTPNS.Scheduler
{
    public class SPOnlineHelper
    {
        public void UploadFile(string SourceFile, string OutputFileName)
        {
            try
            {
                using (ClientContext cnx = new ClientContext("https://test.sharepoint.com/sites/test01"))
                {
                    string password = ConfigurationManager.AppSettings["PassAccount"].ToString();
                    string account = ConfigurationManager.AppSettings["LoginAccount"].ToString();
                    var secret = new SecureString();
                    foreach (char c in password)
                    {
                        secret.AppendChar(c);
                    }
                    cnx.Credentials = new SharePointOnlineCredentials(account, secret);

                    Web web = cnx.Web;

                    FileCreationInformation newFile = new FileCreationInformation();
                    newFile.Content = System.IO.File.ReadAllBytes("document.pdf");

                    //file url is name
                    newFile.Url = @"document.pdf";
                    List docs = web.Lists.GetByTitle("Contact");

                    //get folder and add to that
                    Folder folder = docs.RootFolder.Folders.GetByUrl("demo");
                    File uploadFile = folder.Files.Add(newFile);

                    cnx.Load(docs);
                    cnx.Load(uploadFile);
                    cnx.ExecuteQuery();
                    Console.WriteLine("done UploadFile");
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
    }
}
