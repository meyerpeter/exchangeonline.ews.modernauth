using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ewstest
{
    class Program
    {
        static async System.Threading.Tasks.Task Main(string[] args)
        {
            string myAddress = "MY@ADRESS.DE";
            SecureString myPassword = new NetworkCredential("", "MYPASSWORD").SecurePassword; // Example. Use a better way to get your password!

            // Using Microsoft.Identity.Client 4.22.0
            // Configure the MSAL client to get tokens
            var pcaOptions = new PublicClientApplicationOptions
            {
                ClientId = ConfigurationManager.AppSettings["clientId"],
                TenantId = ConfigurationManager.AppSettings["tenantId"]
            };

            var pca = PublicClientApplicationBuilder
                .CreateWithApplicationOptions(pcaOptions).Build();

           
            // The permission scope required for EWS access
            var ewsScopes = new string[] { "https://outlook.office365.com/EWS.AccessAsUser.All" };

            var authResult = await pca.AcquireTokenByUsernamePassword(ewsScopes, myAddress, myPassword).ExecuteAsync();

            var token = authResult.AccessToken;

            ExchangeService ews = new ExchangeService();
            ews.Credentials = new OAuthCredentials(token);

            ews.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");


            // Sending Mail
            EmailMessage email = new EmailMessage(ews);
            email.ToRecipients.Add("SEND@MAILTO.DE");
            email.Subject = "HelloWorld";
            email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
            email.Send();

            // Receiving Mail
            FolderId SharedMailbox = new FolderId(WellKnownFolderName.Inbox, myAddress);
            ItemView itemView = new ItemView(1000);
            var mails = ews.FindItems(SharedMailbox, itemView);
            foreach(var mail in mails)
            {
                Console.WriteLine(mail.Subject);
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
