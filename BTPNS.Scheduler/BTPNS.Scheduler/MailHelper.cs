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
        public void email_send()
        {
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            mail.From = new MailAddress("vioser7@gmail.com");
            mail.To.Add("kennycassandra@outlook.com");
            mail.Subject = "Test Mail - 1";
            mail.Body = "mail with attachment";

            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment("c:\\log.txt");
            mail.Attachments.Add(attachment);

            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("vioser7@gmail.com", "b1223smz");
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);

        }
    }
}
