using System;
using System.Collections.Generic;
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
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            mail.From = new MailAddress("vioser7@gmail.com");
            mail.To.Add(MailTo);
            mail.Subject = Subject;
            mail.Body = "mail with attachment";
            mail.IsBodyHtml = true;

            System.Net.Mail.Attachment attachment;
            foreach (string s in file_attachment)
            {
                attachment = new System.Net.Mail.Attachment(s);
                mail.Attachments.Add(attachment);
            }
            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("vioser7@gmail.com", "b1223smz");
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);

        }
    }
}
