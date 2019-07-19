using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;

namespace ExcelPasswordAdder
{
   class EmailSender
   {
      public class EmailManager
      {
         private string m_HostName; // Aquí se asignará el host name  

         public EmailManager(string hostName)
         {
            m_HostName = hostName;
         }

         public void SendMail(EmailSendConfigure emailConfig, EmailContent content)
         {
            //Construimos el Mensaje de email para ser enviado
            MailMessage msg = ConstructEmailMessage(emailConfig, content);
            Send(msg, emailConfig);
         }

         // Ponemos las propiedades del email incluyendo "to", "cc", "from", "subject" y "email body"  
         private MailMessage ConstructEmailMessage(EmailSendConfigure emailConfig, EmailContent content)
         {
            //Configuramos el mensaje email obteniendo los "to" y "cc"
            MailMessage msg = new MailMessage();
            foreach (string to in emailConfig.TOs)
            {
               if (!string.IsNullOrEmpty(to))
               {
                  msg.To.Add(to);
               }
            }

            foreach (string cc in emailConfig.CCs)
            {
               if (!string.IsNullOrEmpty(cc))
               {
                  msg.CC.Add(cc);
               }
            }

            msg.From = new MailAddress(emailConfig.From,
                                       emailConfig.FromDisplayName,
                                      Encoding.UTF8);
            msg.IsBodyHtml = content.IsHtml;
            msg.Body = content.Content;
            msg.Priority = emailConfig.Priority;
            msg.Subject = emailConfig.Subject;
            msg.BodyEncoding = Encoding.UTF8;
            msg.SubjectEncoding = Encoding.UTF8;


            //Si no hay archivos, no se agrega ninguno.
            if (content.AttachFileName != null)
            {
               Attachment data = new Attachment(content.AttachFileName,
                                                MediaTypeNames.Application.Zip);
               msg.Attachments.Add(data);
            }

            return msg;
         }

         //Enviamos el mail usando SMTP  
         private void Send(MailMessage message, EmailSendConfigure emailConfig)
         {
            SmtpClient client = new SmtpClient();
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential(
                                  emailConfig.ClientCredentialUserName,
                                  emailConfig.ClientCredentialPassword);
            client.Host = m_HostName;
            client.Port = 25;  // es importante colocar el puerto
            client.EnableSsl = true;  // habilitamos la seguridad SSL

            try
            {
               client.Send(message);
            }
            catch (Exception e)
            {
               Console.WriteLine("Error al enviar el correo: {0}", e.Message);
               throw;
            }
            message.Dispose();
         }

      }

      public class EmailSendConfigure
      {
         public string[] TOs { get; set; }
         public string[] CCs { get; set; }
         public string From { get; set; }
         public string FromDisplayName { get; set; }
         public string Subject { get; set; }
         public MailPriority Priority { get; set; }
         public string ClientCredentialUserName { get; set; }
         public string ClientCredentialPassword { get; set; }
         public EmailSendConfigure()
         {
         }
      }

      public class EmailContent
      {
         public bool IsHtml { get; set; }
         public string Content { get; set; }
         public MemoryStream AttachFileName { get; set; }
      }
   }
}
