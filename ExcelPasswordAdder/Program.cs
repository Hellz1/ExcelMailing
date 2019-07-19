using OfficeOpenXml;
using System;
using System.IO;

//Mailing
using System.Net.Mime;
using System.Net.Mail;
using static ExcelPasswordAdder.EmailSender;

namespace ExcelPasswordAdder
{

   public class Program
   {
      static public void initExcelProcess(byte[] fileArray)
      {

         //Generamos un MemoryStream del arreglo de bytes recibido por ExcelGenerator
         MemoryStream ms = new MemoryStream(fileArray);


         using (ExcelPackage excel = new ExcelPackage(ms))
         {
            {
               //Seleccionamos la hoja de trabajo
               var currentExcelWorksheet = excel.Workbook.Worksheets["Reporte"];

               //Agregamos contraseña
               excel.Encryption.Password = "12";

               //Guardamos en stream y asignamos el smtp que se usará
               var stream = new MemoryStream(excel.GetAsByteArray());
               string smtpServer = "smtp-mail.outlook.com";

               //Invocamos la función de envío de mail
               SendEmail(smtpServer, stream);
            }

         }
      }

      static void SendEmail(string smtpServer, MemoryStream filePath)
      {
         //Creamos un objeto EmailManager con el servidor outlook  
         EmailManager mailMan = new EmailManager(smtpServer);

         //Creamos una configuración de envío
         EmailSendConfigure myConfig = new EmailSendConfigure();

         // coloca tu email que funcionará como host para enviar  
         myConfig.ClientCredentialUserName = "tu-email@outlook.com";
         // coloca tu contraseña del correo que funcionarpa como host
         myConfig.ClientCredentialPassword = "tucontraseña";

         //Configuramos los envíos "PARA" en un arreglo de string, puedes agregar N correos o generar un método que obtenga los correos Se puede enviar a diferentes dominios.
         myConfig.TOs = new string[] { "correo-a-enviar@gmail.com", "segundo-correo@outlook.com", "tercer-correo@yahoo.com" };

         //Puedes agregar Copias con la misma estructura
         myConfig.CCs = new string[] { };

         //Este es el correo del cual aparecerá "DE" puede ser una mascara
         myConfig.From = "no-reply@mycompany.mx";
         myConfig.FromDisplayName = "Nombre a mostrar";

         //COlocamos un orden de prioridad normal
         myConfig.Priority = MailPriority.Normal;

         //Titulo del correo
         myConfig.Subject = "Esta es una prueba de mailing";

         //Generamos un objeto de contenido de mail
         EmailContent myContent = new EmailContent();

         //Puede ser texto plano o un html
         myContent.Content = "The following URLs were down - 1. Foo, 2. bar";

         //El attachFileName puede ser un archivo con un PATH fisico o un stream, como es el caso
         myContent.AttachFileName = filePath;

         /**
          * Puedes agregar al contenido html, para eso debe asignarse IsHtml = true; 
          */
         //myContent.IsHtml = true;

         //Enviamos el eMail
         mailMan.SendMail(myConfig, myContent);
      }
      
   }
}
