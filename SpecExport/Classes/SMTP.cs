using System;
using System.Net.Mail;
using SpecExport.Properties;

namespace SpecExport.Classes
{
    class SMTP
    {
        private string FromMailAddres { get { return Settings.Default.FromMailAddres; } }
        private string DisplayName { get { return Settings.Default.DisplayName; } }
        private string ToMailAddres { get { return Settings.Default.ToMailAddres; } }
        private string MailBody { get { return Settings.Default.MailBody; } }
        private string MailSubject { get { return Settings.Default.MailSubject; } }
        private string DrawingsDirectory { get { return Settings.Default.DrawingsDirectory; } }
        private string NameDoc { get { return $@"{DrawingsDirectory}\Отчет за {DateTime.Now.ToShortDateString().Replace(".", "_")}.xlsx"; } }

        public void SendMail()
        {
            SendMailMessage();
        }
        private System.Security.SecureString GetPassword()
        {
            System.Security.SecureString password = new System.Security.SecureString();
            Console.WriteLine($"Введите пароль от почты {FromMailAddres}");
            foreach (var c in Console.ReadLine())
            {
                password.AppendChar(c);
            }
            return password;
        }

        private void SendMailMessage()
        {
            Program.log = NLog.LogManager.GetCurrentClassLogger();
            try
            {
                MailMessage mm = new MailMessage();
                mm.From = new MailAddress(FromMailAddres, DisplayName);
                mm.Sender = new MailAddress(ToMailAddres);
                mm.Subject = MailSubject;
                mm.Body = MailBody;
                mm.Attachments.Add(new Attachment(NameDoc));
                //// письмо представляет код html
                //mm.IsBodyHtml = true;
                // адрес smtp-сервера и порт, с которого будем отправлять письмо
                SmtpClient smtp = new SmtpClient($"smtp.{new System.Text.RegularExpressions.Regex("@(.*)").Match(FromMailAddres).Groups[1].Value}", 587);
                // логин и пароль
                smtp.Credentials = new System.Net.NetworkCredential(FromMailAddres, GetPassword());
                smtp.EnableSsl = true;
                smtp.Send(mm);
            }
            catch (SmtpException ex)
            {
                Program.log.Error(ex, "Не удалось отправить письмо");
            }
        }
    }
}
