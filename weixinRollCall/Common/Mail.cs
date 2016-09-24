using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net.Mail;

namespace weixinRollCall.Common
{
    public class Mail
    {
        static public int send(string to,string subject,string attachpath)
        {
            Attachment attachment = new Attachment(attachpath);
            SmtpClient smtp = new SmtpClient("smtp.zjut.edu.cn", 25);
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.Credentials = new System.Net.NetworkCredential("cyq", "@chu230628");
            MailMessage message = new MailMessage("cyq@zjut.edu.cn", to,subject, "名单以及点名情况在附件中，请查收！");
            message.Attachments.Add(attachment);
            try {
                smtp.Send(message);
                attachment.Dispose();
                return 1;
            }
            catch
            {
                return 0;
            }

        }        
    }
}