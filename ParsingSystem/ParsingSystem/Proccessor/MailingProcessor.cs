
using System;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Windows;

namespace ParsingSystem.Proccessor
{
    public class MailingProcessor
    {
        public void Send(string address, string subject = "", string body = "", string attachmentFileName = "")
        {
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress("test123@gmail.com");
                    mail.To.Add(address);
                    mail.Subject = subject;
                    mail.Body = body;

                    if (attachmentFileName != null)
                    {
                        Attachment attachment = new Attachment(attachmentFileName, MediaTypeNames.Application.Octet);
                        ContentDisposition disposition = attachment.ContentDisposition;
                        disposition.CreationDate = File.GetCreationTime(attachmentFileName);
                        disposition.ModificationDate = File.GetLastWriteTime(attachmentFileName);
                        disposition.ReadDate = File.GetLastAccessTime(attachmentFileName);
                        disposition.FileName = Path.GetFileName(attachmentFileName);
                        disposition.Size = new FileInfo(attachmentFileName).Length;
                        disposition.DispositionType = DispositionTypeNames.Attachment;
                        mail.Attachments.Add(attachment);
                    }

                    using (SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com", 587))
                    {
                        SmtpServer.UseDefaultCredentials = false; //Need to overwrite this
                        SmtpServer.Credentials = new System.Net.NetworkCredential("harish.1138@gmail.com", "SamplePWD");
                        SmtpServer.EnableSsl = true;
                        SmtpServer.Send(mail);
                    }
                }

                MessageBox.Show("Mail Sent");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
