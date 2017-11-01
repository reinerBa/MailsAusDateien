using MailKit.Net.Smtp;
using MimeKit;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Collections.Specialized;

namespace MailAusDatei
{
    class Program
    {
        static NameValueCollection Get = System.Configuration.ConfigurationManager.AppSettings;
        // Prüft ob wat da iss
        static void Main(string[] args)
        {
            string path = "Kunden";
            if (Directory.Exists(path))
            {
                // This path is a directory
                ProcessDirectory(path);
            }
        }

        // knöpft sich jede File einzeln vor zum verschicken
        public static void ProcessDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory, "*.json");
            using (var client = new SmtpClient())
            {
                // For demo-purposes, accept all SSL certificates (in case the server supports STARTTLS)
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                client.Connect(Get["smptUrl"], 
                    Convert.ToInt32(Get["smptPort"]), false);

                // Note: since we don't have an OAuth2 token, disable
                // the XOAUTH2 authentication mechanism.
                client.AuthenticationMechanisms.Remove("XOAUTH2");

                // Note: only needed if the SMTP server requires authentication
                client.Authenticate(
                    Get["smtpBenutzername"], 
                    Get["smptPasswort"]);

                foreach (string fileName in fileEntries)
                    ProcessFile(fileName, client);

                client.Disconnect(true);
            }
       }

        // 
        public static void ProcessFile(string path, SmtpClient client)
        {
        JObject ConfigFile = JObject.Parse(File.ReadAllText(path));

            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(
                Get["Absendername"], ConfigFile["MarkusEmail"].ToString()));
            message.To.Add(
                new MailboxAddress(ConfigFile["NameEmail"].ToString(), ConfigFile["EmfaengerEmail"].ToString()));

            message.Subject = ConfigFile["Betreff"].ToString();

            message.Body = new TextPart(Get["emailContent"])
            {
                Text = ConfigFile["Text"].ToString()
            };

            client.Send(message);
        }
    }
}
