using System;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Xml.Serialization;
using System.IO;
using System.Collections.Generic;

namespace EWS
{
    class Program
    {
        static int Main(string[] args)
        {
            // Test if input arguments were supplied:
            if (args.Length != 2){
                Console.WriteLine("Please enter a config file");
                return 1;
            }
            string fileCfg = args[0];
            string fileLog = args[1];
            Logger(fileLog, "File cfg: " + fileCfg);
            Logger(fileLog, "File log: " + fileLog);

            try
            {
                StreamReader file = new StreamReader(@fileCfg);
                XmlSerializer reader = new XmlSerializer(typeof(EWS_1C));
                EWS_1C overview = (EWS_1C)reader.Deserialize(file);
                file.Close();
                
                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                service.UseDefaultCredentials = false;
                service.Credentials = new WebCredentials(overview.Login, overview.Password, overview.Domain);
                service.Url = new Uri(overview.url);

                //service.EnableSs
                Logger(fileLog, service.Url.ToString());

                service.TraceEnabled = false;
                service.TraceFlags = TraceFlags.All;

                // emails
                Logger(fileLog, "Писем для отправки: " + overview.emails.Length);
                for (int i = 0; i < overview.emails.Length; i++)
                {
                    EmailOut eml = overview.emails[i];
                    //System.Console.WriteLine("- Subject: " + eml.Subject);

                    EmailMessage message = new EmailMessage(service);
                    message.Subject = eml.Subject;
                    message.Body = eml.Body;
                    for (int j = 0; j < eml.Recipient.Length; j++)
                    {
                        message.ToRecipients.Add(eml.Recipient[j]);
                    }
                    for (int j = 0; j < eml.File.Length; j++)
                    {
                        if (!String.IsNullOrEmpty(eml.File[j]))
                        {
                            message.Attachments.AddFileAttachment(eml.File[j]);
                        }
                    }
                    
                    message.SendAndSaveCopy();
                    Logger(fileLog, "Email № "+ (i+1) +" send recipient = " + eml.Recipient.ToString());
                }

                // Bind the Inbox folder to the service object.
                Folder inbox = Folder.Bind(service, WellKnownFolderName.Inbox);
                // The search filter to get unread email.
                //SearchFilter sf = new SearchFilter.SearchFilterCollection(
                //    LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
                SearchFilter sf =
                    new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived, DateTime.Now.AddDays(-5));

                ItemView view = new ItemView(30); //Больше 30 писем не получать

                // Fire the query for the unread items.
                // This method call results in a FindItem call to EWS.
                FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, sf, view);

                //var emailProps = new PropertySet(ItemSchema.MimeContent, ItemSchema.Body,
                //    ItemSchema.InternetMessageHeaders);
                
                var dirOut = Path.GetDirectoryName(fileCfg);
                Logger(fileLog, "dirOut = " + dirOut);

                List<EmailIn> list = new List<EmailIn>();

                Logger(fileLog, "Писем получено: " + findResults.TotalCount.ToString());
                foreach (EmailMessage i in findResults)
                {
                    //var email_ews = EmailMessage.Bind(service, i.Id, emailProps);
                    i.Load();
                    if (!i.IsRead) {
                        i.IsRead = true;
                        i.Update(ConflictResolutionMode.AutoResolve);
                    }

                    var eml = new EmailIn();
                    eml.IdObj = Guid.NewGuid().ToString();
                    eml.Subject = i.Subject;
                    eml.Body = i.Body;
                    List<String> listRecipients = new List<String>();
                    foreach (EmailAddress ToRecipient in i.ToRecipients)
                    {
                        listRecipients.Add(ToRecipient.Name + " <" + ToRecipient.Address + ">");
                    }
                    eml.Recipient = listRecipients.ToArray();
                    eml.From = i.From.Name + " <"+i.From.Address+">";
                    eml.Id = i.InternetMessageId;
                    eml.DateSend = i.DateTimeSent;
                    Logger(fileLog, "id email: "+eml.Id);
                    Logger(fileLog, "id: " + eml.IdObj);


                    if (i.HasAttachments)
                    {
                        string dirEmail = Path.Combine(dirOut, "files", eml.IdObj);
                        if (!Directory.Exists(dirEmail))
                        {
                            Directory.CreateDirectory(dirEmail);
                        }
                        List<String> listFiles = new List<String>();
                        foreach (Attachment attachment in i.Attachments)
                        {
                            if (attachment is FileAttachment)
                            {
                                FileAttachment fileAttachment = attachment as FileAttachment;
                                string fileEmail = Path.Combine(dirEmail, fileAttachment.Name);
                                fileAttachment.Load(fileEmail);
                                listFiles.Add(fileEmail);
                            }
                            else
                            {
                                //Вложенные письма нам не нужны
                                //ItemAttachment itemAttachment = attachment as ItemAttachment;
                                //itemAttachment.Load();
                            }
                        }

                        eml.File = listFiles.ToArray();
                    }

                    list.Add(eml);
                }
                
                var outData = new Messages();
                outData.emails = list.ToArray();

                var writer = new System.Xml.Serialization.XmlSerializer(typeof(Messages));
                var wfile = new System.IO.StreamWriter(@dirOut + "\\messages.xml");
                writer.Serialize(wfile, outData);
                wfile.Close();


                Logger(fileLog, "Done!");
            }
            catch (Exception e)
            {
                Logger(fileLog, e.Message);
            }

            return 0;
        }
        private static bool CertificateValidationCallBack(
            object sender,
            System.Security.Cryptography.X509Certificates.X509Certificate certificate,
            System.Security.Cryptography.X509Certificates.X509Chain chain,
            System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }
        public static void Logger(string strFile, string lines)
        {
            string logMessage = String.Format("{0} ({1}) - {2}",
                DateTime.Now.ToShortDateString(),
                DateTime.Now.ToShortTimeString(),
                lines);
            
            if (String.IsNullOrEmpty(strFile))
            {
                Console.WriteLine(logMessage);
            }else{
                // Write the string to a file.append mode is enabled so that the log
                // lines get appended to  test.txt than wiping content and writing the log

                StreamWriter file = new StreamWriter(strFile, true);
                file.WriteLine(logMessage);

                file.Close();
            }
        }
    }
    public class EWS_1C
    {
        public String Login;
        public String Password;
        public String Domain;
        public String url;
        public EmailOut[] emails;
        public bool SendNot;
        public bool GetNot;
    }
    public class EmailOut
    {
        public String Subject;
        public String Body;
        public String[] Recipient;
        public String[] File;
    }
    public class Messages
    {
        public EmailIn[] emails;
    }
    public class EmailIn
    {
        public String Subject;
        public String Body;
        public String From;
        public String[] Recipient;
        public String[] File;
        public String Id;
        public String IdObj;
        public DateTime DateSend;
    }
}
