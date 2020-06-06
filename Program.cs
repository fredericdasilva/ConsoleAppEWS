using Microsoft.Exchange.WebServices.Data;
using System;
using System.Configuration;

namespace ConsoleAppEWS
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Hello World!");
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            service.Credentials = new WebCredentials("XXXXXXXX@hotmail.com", "Password");
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;
            service.AutodiscoverUrl("XXXXXXXX@hotmail.com", RedirectionUrlValidationCallback);

            //sendMail(service);
            readMail(service);



        }

        private static void sendMail(ExchangeService service)
        {
            EmailMessage email = new EmailMessage(service);
            email.ToRecipients.Add("recipient1@free.fr");
            email.Subject = "HelloWorld";
            email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
            email.Send();
        }

        private static void readMail(ExchangeService service)
        {
            //var items = service.FindItems(
            //    //Find Mails from Inbox of the given Mailbox
            //    new FolderId(WellKnownFolderName.Inbox, new Mailbox(ConfigurationManager.AppSettings["MailBox"].ToString())),
            //    //Filter criterion
            //    new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter[] {
            //       new SearchFilter.ContainsSubstring(ItemSchema.Subject, ConfigurationManager.AppSettings["ValidEmailIdentifier"].ToString()),
            //       new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false)
            //    }),
        
            //    //View Size as 15
            //    new ItemView(15));

            var items = service.FindItems(
               //Find Mails from Inbox of the given Mailbox
               new FolderId(WellKnownFolderName.Inbox),
               //View Size as 15
               new ItemView(15));

            foreach (EmailMessage msg in items)
            {
                //Retrieve Additional data for Email
                EmailMessage message = EmailMessage.Bind(service,
                     (EmailMessage.Bind(service, msg.Id)).Id,
                      new PropertySet(BasePropertySet.FirstClassProperties,
                      new ExtendedPropertyDefinition(0x1013, MapiPropertyType.Binary)));

                Console.Write(message.Subject);
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
