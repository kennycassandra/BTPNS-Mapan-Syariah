using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;
using Microsoft.SharePoint.Client;
using System.Data.SqlClient;

namespace BTPNS.Scheduler
{
    public class CleansingHelper
    {
        DataBaseManager db = new DataBaseManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlDataReader reader = null;
        Utility util = new Utility();

        public static CamlQuery CreateAllFilesQuery()
        {
            var qry = new CamlQuery();
            qry.ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq></Where></Query></View>";
            return qry;
        }

        public void CleansingLogExcelData(string OutputFolder)
        {
            try
            {
                db.OpenConnection(ref sqlConn);
                db.cmd.CommandText = "usp_Cleansing";
                db.cmd.ExecuteNonQuery();
                db.CloseConnection(ref sqlConn);
            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "CleansingLogExcelData");

            }
        }
        public void CleansingSPFiles(string OutputFolder, Microsoft.SharePoint.Client.ClientContext cl)
        {
            try
            {
                int CleansingDays = GetCleansingDays();
                List<string> DocLibs = new List<string>();
                DocLibs.Add("AP3R");
                DocLibs.Add("CIF");
                DocLibs.Add("Pembiayaan");
                DocLibs.Add("Persetujuan_Pembiayaan");
                DocLibs.Add("SMS_Notification");

                string UrlSPOnPrem = ConfigurationManager.AppSettings["SharePointOnPremURL"].ToString();

                using (var ctx = cl)
                {

                    var results = new Dictionary<string, IEnumerable<Microsoft.SharePoint.Client.File>>();
                    var lists = ctx.LoadQuery(ctx.Web.Lists.Where(l => l.BaseType == BaseType.DocumentLibrary));
                    ctx.ExecuteQuery();
                    foreach (var list in lists)
                    {
                        var items = list.GetItems(CreateAllFilesQuery());
                        ctx.Load(items, icol => icol.Include(i => i.File));
                        results[list.Title] = items.Select(i => i.File);
                    }
                    ctx.ExecuteQuery();

                    foreach (var result in results)
                    {

                        Console.WriteLine("List: {0}", result.Key);
                        if (DocLibs.Contains(result.Key))
                        {
                            foreach (var file in result.Value)
                            {
                                DateTime dtModified = file.TimeCreated;
                                DateTime dtCurrent = DateTime.Now;

                                if ((dtCurrent - dtModified).TotalDays >= CleansingDays)
                                {
                                    //System.IO.File.Delete(f);
                                    file.DeleteObject();
                                    ctx.ExecuteQuery();
                                }


                                Console.WriteLine("File: {0}-{1}-{2}", file.Name, file.TimeLastModified, file.TimeCreated);
                            }
                        }
                    }
                }

                //File file = context.Web.GetFileByLinkingUrl(UrlSPOnPrem);

                //context.Load(file, fv => fv.Name, fv => fv.Exists, fv => fv.TimeLastModified);
                //context.ExecuteQuery();
                //FileVersionCollection fileVersionCollection = file.Versions;
                //context.Load(fileVersionCollection);
                //context.ExecuteQuery();

                //foreach (FileVersion fileVersion in fileVersionCollection)
                //{
                //    context.Load(fileVersion, fv => fv.Created);
                //    context.ExecuteQuery();
                //    DateTime ModifiedTime = fileVersion.Created;
                //    Console.WriteLine("File : {0} {1} {2}", fileVersion.Url, fileVersion.Context.Url, fileVersion.Created);
                //}
            }
            catch (Exception ex)
            {
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "CleansingSPFiles");
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
                int CleansingDays = GetCleansingDays();

                string OutputFolderTXT = OutputFolder + @"Output\TXT";
                string[] folders = Directory.GetDirectories(OutputFolder + "Output");
                foreach (string fol in folders)
                {
                    string[] filePaths = Directory.GetFiles(fol);
                    foreach (string f in filePaths)
                    {
                        DateTime dtModified = System.IO.File.GetLastWriteTime(f);
                        DateTime dtCurrent = DateTime.Now;

                        if ((dtCurrent - dtModified).TotalDays >= CleansingDays)
                        {
                            System.IO.File.Delete(f);
                        }
                    }
                }


                folders = Directory.GetDirectories(OutputFolderTXT);
                foreach (string fol in folders)
                {
                    string[] filePaths = Directory.GetFiles(fol);
                    foreach (string f in filePaths)
                    {
                        DateTime dtModified = System.IO.File.GetLastWriteTime(f);
                        DateTime dtCurrent = DateTime.Now;

                        if ((dtCurrent - dtModified).TotalDays >= CleansingDays)
                        {
                            System.IO.File.Delete(f);
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
