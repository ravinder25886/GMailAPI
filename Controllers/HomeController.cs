using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace GoogleGmailAPI.App.Controllers
{
    public class HomeController : Controller
    {   // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/gmail-dotnet-quickstart.json
        static string[] Scopes = {
            GmailService.Scope.GmailSend,
            GmailService.Scope.GmailReadonly,
            GmailService.Scope.GmailLabels,
            GmailService.Scope.GmailCompose,
            GmailService.Scope.GmailInsert,
            GmailService.Scope.GmailModify,
            GmailService.Scope.MailGoogleCom,
            //GmailService.Scope.GmailMetadata

        };
        static string ApplicationName = "Gmail API .NET Quickstart";
        
        public ActionResult Index()
        {
            //SendMail("ravinder25886@gmail.com", "test", "This is test message");
            return View(LoadInbox());
        }
        public void SendMail(string txtTo, string txtSubject, string txtMessage)
        {
            UserCredential credential;
            //read credentials file
            using (FileStream stream = new FileStream(Server.MapPath("credentials.json"), FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/gmail-dotnet-quickstart.json");
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
            }

          

            string plainText = $"To: {txtTo}\r\n" +
                               $"Subject: {txtSubject}\r\n" +
                               "Content-Type: text/html; charset=utf-8\r\n\r\n" +
                               $"<h1>{txtMessage}</h1>";

            //call gmail service
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            var newMsg = new Google.Apis.Gmail.v1.Data.Message();
            newMsg.Raw = Base64UrlEncode(plainText.ToString());
            service.Users.Messages.Send(newMsg, "me").Execute();
            //MessageBox.Show("Your email has been successfully sent !", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public List<MyInbox> LoadInbox()
        {
            string yourEmailId = "me";
            UserCredential credential;

            using (var stream =new FileStream(Server.MapPath("credentials.json"), FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/gmail-dotnet-quickstart.json");
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,Scopes,"user",CancellationToken.None,
                new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Gmail API service.   
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            var inboxlistRequest = service.Users.Messages.List(yourEmailId);
            inboxlistRequest.LabelIds = "INBOX";
            inboxlistRequest.IncludeSpamTrash = false;
            inboxlistRequest.MaxResults = 10;
            //get our emails   
            List<MyInbox> myInboxes = new List<MyInbox>();
            var emailListResponse = inboxlistRequest.Execute();
            if (emailListResponse != null && emailListResponse.Messages != null)
            {
                //loop through each email and get what fields you want...   
               
                foreach (var email in emailListResponse.Messages)
                {
                    var emailInfoRequest = service.Users.Messages.Get(yourEmailId, email.Id);
                    var emailInfoResponse = emailInfoRequest.Execute();
                    if (emailInfoResponse != null)
                    {
                        String from = "";
                        String date = "";
                        String subject = "";
                        //loop through the headers to get from,date,subject, body  
                        foreach (var mParts in emailInfoResponse.Payload.Headers)
                        {
                            if (mParts.Name == "Date")
                            {
                                date = mParts.Value;
                            }
                            else if (mParts.Name == "From")
                            {
                                from = mParts.Value;
                            }
                            else if (mParts.Name == "Subject")
                            {
                                subject = mParts.Value;
                            }
                            if (date != "" && from != "")
                            {
                                if (emailInfoResponse.Payload.Parts != null)
                                {
                                    foreach (MessagePart p in emailInfoResponse.Payload.Parts)
                                    {
                                        if (p.MimeType == "text/html")
                                        {
                                            byte[] data = FromBase64ForUrlString(p.Body.Data);
                                            string decodedString = Encoding.UTF8.GetString(data);
                                        }
                                    }
                                }
                                myInboxes.Add(new MyInbox { From = from, Date = date, Subject = subject });
                            }
                            
                        }
                    }
                }
            }
            return myInboxes;
        }
        public string Base64UrlEncode(string input)
        {
            var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(inputBytes).Replace("+", "-").Replace("/", "_").Replace("=", "");
        }
        public  byte[] FromBase64ForUrlString(string base64ForUrlInput)
        {
            int padChars = (base64ForUrlInput.Length % 4) == 0 ? 0 : (4 - (base64ForUrlInput.Length % 4));
            StringBuilder result = new StringBuilder(base64ForUrlInput, base64ForUrlInput.Length + padChars);
            result.Append(String.Empty.PadRight(padChars, '='));
            result.Replace('-', '+');
            result.Replace('_', '/');
            return Convert.FromBase64String(result.ToString());
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        
    }
    public class MyInbox {
        public string From { get; set; }
        public string Date { get; set; }
        public string Subject { get; set; }
    }
}
