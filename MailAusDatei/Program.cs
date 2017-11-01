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

        static void Main(string[] args)
        {
            EasyTimer.SetInterval(
                ProcessIfDict
                , Convert.ToInt32(Get["sendeIntervalStunden"]) * 60 * 60 * 1000);
        }

        // Prüft ob wat da iss und macht dann.
        public static void ProcessIfDict()
        {
            string path = Get["Kunden"];
            if (Directory.Exists(path))
                ProcessDirectory(path);    // This path is a directory
            
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
                    if(!fileName.EndsWith("Beispiel.json"))
                        ProcessFile(fileName, client);
                
                client.Disconnect(true);
            }
       }

        // macht aus dem json eine email
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

    /// <summary>
    ///  iss halt notwendig als interval helper
    /// </summary>
    public static class EasyTimer
    {
        public static IDisposable SetInterval(Action method, int delayInMilliseconds)
        {
            System.Timers.Timer timer = new System.Timers.Timer(delayInMilliseconds);
            timer.Elapsed += (source, e) =>
            {
                method();
            };

            timer.Enabled = true;
            timer.Start();

            // Returns a stop handle which can be used for stopping
            // the timer, if required
            return timer as IDisposable;
        }

    }
}
