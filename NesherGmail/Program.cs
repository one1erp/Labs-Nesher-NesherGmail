using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace NesherGmail
{
    public class NesherGmailCls
    {
        public static void Send(string recipientEmail, string subject, string body, List<string> AtachmentPathes)
        {

            try
            {



                // Sender's email address and password
                string senderEmail = GetUserNamePwd("Sender");// "micro.bio2123@gmail.com";
                string senderPassword = GetUserNamePwd("Password");// "hqfz amgw nnnp zjmz";
                string SenderDisplayName = GetUserNamePwd("SenderDisplayName");// המכון למיקרוביולוגיה


                MailAddress mf = new MailAddress(senderEmail, SenderDisplayName);
                // Create the MailMessage object
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = mf;
             
                mailMessage.Subject = subject;
                mailMessage.Body = body;


                foreach (var address in recipientEmail.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    mailMessage.To.Add(address);
                }

                if (AtachmentPathes != null)
                {
                    foreach (var item in AtachmentPathes)
                    {
                        if (File.Exists(item))
                        {
                            mailMessage.Attachments.Add(new Attachment(item));
                        }
                    }
                }
                // Create the SmtpClient
                SmtpClient smtpClient = new SmtpClient("smtp.gmail.com")
                {
                    Port = 587,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(senderEmail, senderPassword),
                    EnableSsl = true,
                };

                try
                {
                    // Send the email
                    smtpClient.Send(mailMessage);
                    Common.Logger.WriteLogFile("Email sent successfully.");
                }
                catch (Exception ex)
                {
                    Common.Logger.WriteLogFile($"Error sending email: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Common.Logger.WriteLogFile(ex);
            }
        }


        private static string GetUserNamePwd(string key)
        {


            try
            {


                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                ExeConfigurationFileMap map = new ExeConfigurationFileMap();
                map.ExeConfigFilename = assemblyPath + ".config";
                Configuration cfg = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                var appSettings = cfg.AppSettings;

                string s = appSettings.Settings[key].Value;

                return s;
            }
            catch (Exception e)
            {

                return null;
            }
        }

    }


}

