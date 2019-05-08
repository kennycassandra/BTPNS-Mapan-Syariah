using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices.AccountManagement;
using System.Data;
using System.DirectoryServices;
using System.DirectoryServices.Protocols;
using System.Data.SqlClient;

namespace BTPNS.Scheduler
{
    public class User
    {
        public string DisplayName { get; set; }
        public string UserName { get; set; }
        public string Mail { get; set; }
        public string EmployeeId { get; set; }
    }
    public class ActiveDirectoryHelper
    {
        DataBaseManager db = new DataBaseManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlDataReader reader = null;
        Utility util = new Utility();

        public string getAttributeValue(SearchResult sr, string attrName)
        {
            try
            {
                return (string)sr.Properties[attrName][0];
            }
            catch 
            {
                return "";
            }
        }

        public string GetEmailAD(string MailAddress)
        {
            string email = "";
            try
            {
                db.OpenConnection(ref sqlConn);
                email = db.GetValueFromQuery("select * from officer_AD where employeeId = '" + 
                    MailAddress.Replace("@mail.btpnsyariah.com","") + "'", "Mail");
                db.CloseConnection(ref sqlConn);
                return email;
            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                return MailAddress;
            }
        }

        public void GetADUsers(string OutputFolder, string LDAPPath)
        {
            try
            {
                List<User> AdUsers = new List<User>();
                string domainPath = LDAPPath;
                DirectoryEntry searchroot = new DirectoryEntry(domainPath);
                DirectorySearcher search = new DirectorySearcher(searchroot);
                search.Filter = "(&(objectClass=user)(objectCategory=person))";
                search.PropertiesToLoad.Add("samaccountname");
                search.PropertiesToLoad.Add("displayname");
                search.PropertiesToLoad.Add("employeeID");
                search.PropertiesToLoad.Add("mail");
                SearchResult result;
                SearchResultCollection resultCol = search.FindAll();
                if (resultCol != null)
                {
                    db.OpenConnection(ref sqlConn, true);
                    for (int i = 0; i < resultCol.Count; i++)
                    {
                        result = resultCol[i];
                        User adUser = new User();
                        adUser.DisplayName = getAttributeValue(result, "displayname");
                        adUser.UserName = getAttributeValue(result, "samaccountname");
                        adUser.Mail = getAttributeValue(result, "mail");
                        adUser.EmployeeId = getAttributeValue(result, "employeeID");
                        AdUsers.Add(adUser);

                        if(!string.IsNullOrEmpty(adUser.EmployeeId))
                        {
                            //@EmployeeId varchar(50),
                            //@DisplayName varchar(100),
                            //@AccountUserName varchar(255),
                            //@Mail varchar(255)                            
                            db.cmd.CommandText = "usp_Officer_AD_Insert";
                            db.cmd.CommandType = CommandType.StoredProcedure;
                            db.cmd.Parameters.Clear();
                            db.AddInParameter(db.cmd, "EmployeeId", adUser.EmployeeId);
                            db.AddInParameter(db.cmd, "DisplayName", adUser.DisplayName);
                            db.AddInParameter(db.cmd, "AccountUserName", adUser.UserName);
                            db.AddInParameter(db.cmd, "Mail", adUser.Mail);

                            db.cmd.ExecuteNonQuery();
                        }

                    }
                    db.CloseConnection(ref sqlConn, true);
                }
            }
            catch (Exception ex)
            {
                db.CloseConnection(ref sqlConn);
                new GenerateTxt().GenerateTxtLogError(OutputFolder, ex.Message, "GetADUsers");
            }
        }
        public bool AuthenticateUser(string Domain, string Username, string Password, string LDAP_Path, ref string Errmsg)
        {
            Errmsg = "";
            string domainAndUsername = Domain + "\\" + Username;
            //DirectoryEntry entry = new DirectoryEntry(LDAP_Path, domainAndUsername, Password);
            DirectoryEntry entry = new DirectoryEntry(LDAP_Path);
            //entry.AuthenticationType = AuthenticationTypes.Secure;
            entry.AuthenticationType = AuthenticationTypes.ServerBind;
            try
            {
                DirectorySearcher search = new DirectorySearcher(entry);

                search.Filter = "(SAMAccountName=" + Username + ")";

                search.PropertiesToLoad.Add("cn");
                search.PropertiesToLoad.Add("mail");
                search.PropertiesToLoad.Add("employeeid");

                SearchResult result = search.FindOne();

                if (result == null)
                {
                    return false;
                }
                // Update the new path to the user in the directory

                LDAP_Path = result.Path;

                string _filterAttribute = (String)result.Properties["cn"][0];

                foreach(SearchResult rs in search.FindAll())
                {
                    Console.WriteLine((String)rs.Properties["cn"][0] + " - " + (String)rs.Properties["mail"][0] + (String)rs.Properties["employeeid"][0]);
                }

                //Console.WriteLine(_filterAttribute);
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Errmsg = ex.Message;
                return false;
                throw new Exception("Error authenticating user." + ex.Message);
            }

            return true;
        }
        public void GetActiveDirectoryList()
        {
            // bind to your domain
           PrincipalContext pricipalContext = new PrincipalContext(ContextType.Domain, "demodomain", "DC=demodomain,DC=com, OU=users");

            // find the user by identity (or many other ways)
            //UserPrincipal user = UserPrincipal.FindByIdentity(pricipalContext, "cn=leonard");
            //Console.WriteLine("{0} - {1}", user.EmailAddress, user.DisplayName);
            //Console.ReadLine();
            string msg = "";
            AuthenticateUser("demodomain", "kenny", "pass@word1", "LDAP://demodomain.com", ref msg);
            Console.WriteLine(msg);
            Console.ReadLine();
        }
    }
}
