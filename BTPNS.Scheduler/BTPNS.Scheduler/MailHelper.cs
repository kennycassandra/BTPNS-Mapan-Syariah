using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace BTPNS.Scheduler
{
    class MailHelper
    {
        public void email_send(List<string> file_attachment, string Subject, string MailTo)
        {
            MailMessage mail = new MailMessage();

            SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["SMTP"].ToString());
            mail.From = new MailAddress(ConfigurationManager.AppSettings["From"].ToString());
            mail.To.Add(MailTo);
            mail.Subject = Subject;
            mail.Body = "mail with attachment";
            mail.IsBodyHtml = true;

            Attachment attachment;
            foreach (string s in file_attachment)
            {
                attachment = new System.Net.Mail.Attachment(s);
                mail.Attachments.Add(attachment);
            }
            SmtpServer.Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"].ToString());

            string SMTPUser = ConfigurationManager.AppSettings["SMTPUser"].ToString();
            string SMTPPass = ConfigurationManager.AppSettings["SMTPPass"].ToString();

            if (!string.IsNullOrEmpty(SMTPUser))
            {
                SmtpServer.Credentials = new System.Net.NetworkCredential(SMTPUser, SMTPPass);
            }
            SmtpServer.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableSsl"].ToString());

            SmtpServer.Send(mail);

        }

        public void TestSendEmail()
        {
            try
            {
                MailMessage mail = new MailMessage();

                SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["SMTP"].ToString());
                mail.From = new MailAddress(ConfigurationManager.AppSettings["From"].ToString());
                mail.To.Add(ConfigurationManager.AppSettings["TestSendTo"].ToString());
                mail.Subject = "Test Send Email";
                mail.Body = "This Is Body Email";
                mail.IsBodyHtml = true;

                SmtpServer.Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"].ToString());

                string SMTPUser = ConfigurationManager.AppSettings["SMTPUser"].ToString();
                string SMTPPass = ConfigurationManager.AppSettings["SMTPPass"].ToString();

                if (!string.IsNullOrEmpty(SMTPUser))
                {
                    SmtpServer.Credentials = new System.Net.NetworkCredential(SMTPUser, SMTPPass);
                }
                SmtpServer.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableSsl"].ToString());

                SmtpServer.Send(mail);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
