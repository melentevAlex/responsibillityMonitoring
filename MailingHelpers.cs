using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Linq;
using System.Net.Mail;


namespace responsibillityMonitoring
{
    public static class MailingHelpers
    {
        /// <summary>
        /// Список получателей рассылки.
        /// </summary>
        private static readonly string[] _receivers =
        {
                     "aleksey.melentev@company.ru"
              };

        /// <summary>
        /// Отправка сообщения.
        /// </summary>
        public static void SendMail(string name, string subject, string body, string[] receivers = null, string[] cc = null, string[] attachments = null)
        {
            MailMessage message = new MailMessage();
            message.From = new MailAddress("noreply@company.ru", name);

            if (receivers == null) { receivers = _receivers; }
            message.To.Add(receivers.Select(x => $"{x.ToString()}").Aggregate((current, a) => current + ',' + a));
            if (cc != null) { message.CC.Add(cc.Select(x => $"{x.ToString()}").Aggregate((current, a) => current + ',' + a)); }

            message.IsBodyHtml = true;
            message.Subject = subject;
            message.Body = body;

            if (attachments != null && attachments.Length > 0)
            {
                foreach (string attachment in attachments)
                {
                    message.Attachments.Add(new Attachment(attachment));
                }
            }

            SmtpClient client = new SmtpClient("mail.company.ru");
            client.Send(message);
            message.Dispose();
        }
    }

}
