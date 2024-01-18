using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using System.IO;

namespace ProvaN3_SolicitaSuporte
{
    public class Email
    {
        public string FromEmail { get; set; }
        public string Password { get; set; }

        private readonly string _emailFrom = "prova_n3hiper@outlook.com";
        private readonly string _password = "Hiper123*";

        public Email( )
        {
            FromEmail = _emailFrom;
            Password = _password;
        }

        public void SendImageToSuport(string toEmail, string subject, string body, string attachmentPath)
        {
            try
            {
                MailMessage message = new MailMessage();

                string[] to = toEmail.Split(';');
                foreach (string destinatario in to)
                {
                    message.To.Add(destinatario.Trim());
                }
                message.From = new MailAddress(FromEmail);
                message.Subject = string.IsNullOrEmpty(subject)  ? "Assunto teste" : subject;
                message.Body = string.IsNullOrEmpty(body) ? "Mensagem padrão de teste\nAtenciosamente,\nMarcos Kohler" : body;



                Attachment anexo = new Attachment(attachmentPath, MediaTypeNames.Application.Octet);
                ContentDisposition disposition = anexo.ContentDisposition;
                disposition.CreationDate = File.GetCreationTime(attachmentPath);
                disposition.ModificationDate = File.GetLastWriteTime(attachmentPath);
                disposition.ReadDate = File.GetLastAccessTime(attachmentPath);
                message.Attachments.Add(anexo);

                SmtpClient clienteSmtp = new SmtpClient("smtp.office365.com");
                clienteSmtp.Port = 587;
                clienteSmtp.Credentials = new NetworkCredential(FromEmail, Password);
                clienteSmtp.EnableSsl = true;

                clienteSmtp.Send(message);
                MessageBox.Show("E-mail enviado com sucesso!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao enviar e-mail: {ex.Message}\nStackTrace: {ex.StackTrace}");
            }
        }
    }




}
