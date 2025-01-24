using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Xml.Linq;

namespace Company.Email
{
    public class Mailtest
    {
        private void Test()
        {
            // Erstellen einer neuen E-Mail-Nachricht
            // Create a new mail message
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("sender@example.com");
            mail.To.Add("test@example.de");
            mail.Subject = "Test";
            mail.Body = "Das ist eine Test eMail.";

            // Speichern der E-Mail in einer .eml-Datei
            // Save the mail later in an .eml file
            string emlFilePath = "mail.eml";

            using (var ms = new MemoryStream())
            using (var smtpClient = new SmtpClient())
            {
                // Bild speichern in MemoryStream
                // Load the attachment in a MemoryStream
                Properties.Resources.dog.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                ms.Position = 0;  // WICHTIG: Stream-Position zurücksetzen
				// Important: reset the stream position

                // Anhang erstellen
                // create the attachement
                var attachment = new Attachment(ms, "Anhang.jpg");
                attachment.ContentType = new System.Net.Mime.ContentType("image/jpeg");
                attachment.Name = "Anhang.jpg";
                attachment.ContentId = "Anhang.jpg";
                attachment.ContentDisposition.FileName = "Anhang.jpg";
                attachment.ContentDisposition.DispositionType = System.Net.Mime.DispositionTypeNames.Attachment;
                attachment.ContentDisposition.Size = ms.Length;
                attachment.ContentDisposition.CreationDate = DateTime.Now;
                attachment.ContentDisposition.ModificationDate = DateTime.Now;
                attachment.TransferEncoding = System.Net.Mime.TransferEncoding.Base64;
                mail.Attachments.Add(attachment);

                // SMTP Client Konfiguration
                // configure the smtp client
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                smtpClient.PickupDirectoryLocation = Environment.CurrentDirectory;
                smtpClient.Send(mail);

                // Die erstellte Datei finden und verschieben
                // Find the .eml file and move it to the desired position
                var tempDirectory = new DirectoryInfo(smtpClient.PickupDirectoryLocation);
                var file = tempDirectory.GetFiles().OrderByDescending(f => f.LastWriteTime).First();

                if (File.Exists(emlFilePath))
                    File.Delete(emlFilePath);
                file.MoveTo(emlFilePath);
            }

            // X-Unsent Header hinzufügen
            // Add X-Unsent header so Outlook will open it as draft
            var allLines = File.ReadAllLines(emlFilePath, Encoding.UTF8).ToList();
            allLines.Insert(0, "X-Unsent: 1");

            File.WriteAllLines(emlFilePath, allLines, Encoding.UTF8);

            // Die .eml-Datei öffnen
            // open the .eml file
            Process.Start(new ProcessStartInfo(emlFilePath) { UseShellExecute = true });
        }
    }
}