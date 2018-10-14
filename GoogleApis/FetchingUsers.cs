using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography.X509Certificates;
using Google.Apis;
using Google.Apis.Admin.Directory.directory_v1;
using Google.Apis.Admin.Directory.directory_v1.Data;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;

namespace GoogleApis
{
    static class FetchingUsers
    {
        private static string _serviceAccount = "groupid@adminproject-219211.iam.gserviceaccount.com";
        private static string _impersonatedUser = "umair@bibelotz.com";

        public static void Run()
        {
            AccountFetchAll();
        }

        private static void AccountFetchAll()
        {
            var certificate = GetCertificateFromFile();
            var scopes = new[] { DirectoryService.Scope.AdminDirectoryUserReadonly };
            var creds = GetServiceAccountCredentials(_serviceAccount, _impersonatedUser, scopes, certificate);
            var service = GetDirectoryService(creds);
            var userListRaw = GetUsersListRaw(service);
            var userModelList = GetMailBoxUsers(userListRaw);
            PrintMailBoxUsers(userModelList);      
        }

        private static void PrintMailBoxUsers(IEnumerable<MailBoxUser> list)
        {
            Console.WriteLine($" MAILBOX USERS FOUND: {list.Count()}\n\n");
            int count = 0;
            foreach (var mailbox in list)
            {
                Console.WriteLine($" Mailbox {count++}:");
                Console.WriteLine($"\tPrimary Email: {mailbox.PrimaryEmail}");
                Console.Write($"\tNon-Editable Aliases: ");
                if(mailbox.NonEditableAliases.Count > 0)
                {
                    Console.WriteLine($"{mailbox.NonEditableAliases.Count}:");
                    foreach (var alias in mailbox.NonEditableAliases)
                    {
                        Console.WriteLine($"\t\t{alias}");
                    }
                }
                else
                {
                    Console.WriteLine("No Non-Editable aliases found.");
                }

                Console.Write($"\tAliases: ");
                if (mailbox.NonEditableAliases.Count > 0)
                {
                    Console.WriteLine($"{mailbox.Aliases.Count}:");
                    foreach (var alias in mailbox.Aliases)
                    {
                        Console.WriteLine($"\t\t{alias}");
                    }
                }
                else
                {
                    Console.WriteLine("No aliases found.");
                }

                Console.Write($"\tEmails: ");
                if (mailbox.Emails.Count > 0)
                {
                    Console.WriteLine($"{mailbox.Emails.Count}:");
                    foreach (var email in mailbox.Emails)
                    {
                        Console.WriteLine($"\t\t{email}");
                    }
                }
                else
                {
                    Console.WriteLine("No other emails found.");
                }
            }
        }

        private static List<MailBoxUser> GetMailBoxUsers(IEnumerable<User> userListRaw)
        {
            var mailboxList = new List<MailBoxUser>();
            foreach (var user in userListRaw)
            {
                var isMailbox = user.IsMailboxSetup ?? false;
                if(!isMailbox)
                {
                    continue;
                }
                var mailbox = new MailBoxUser();
                mailbox.PrimaryEmail = user.PrimaryEmail;
                mailbox.NonEditableAliases = new List<string>();
                if(user.NonEditableAliases != null)
                {
                    foreach (var nonEditableEmail in user.NonEditableAliases)
                    {
                        mailbox.NonEditableAliases.Add(nonEditableEmail);
                    }
                }
                mailbox.Aliases = new List<string>();
                if(user.Aliases != null)
                {
                    foreach (var alias in user.Aliases)
                    {
                        mailbox.Aliases.Add(alias);
                    }
                }
                mailbox.Emails = new List<string>();
                if(mailbox.Emails != null)
                {
                    foreach (var email in user.Emails)
                    {
                        mailbox.Emails.Add(email.Address);
                    }
                }
                mailboxList.Add(mailbox);
            }
            return mailboxList;
        }

        private static IEnumerable<User> GetUsersListRaw(DirectoryService service, string customerId = "my_customer")
        {
            var request = service.Users.List();
            request.Customer = customerId;
            var response = request.Execute();
            var users = response.UsersValue;
            return users;
        }

        private static DirectoryService GetDirectoryService(ServiceAccountCredential credentials, string appName = "Test Application")
        {
            var service = new DirectoryService(new BaseClientService.Initializer()
            {
                ApplicationName = appName,
                HttpClientInitializer = credentials
            });
            return service;
        }

        private static ServiceAccountCredential GetServiceAccountCredentials(string serviceAccount, string impersonatedAccount, string[] scopes, X509Certificate2 certificate)
        {
            var creds = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(serviceAccount)
            {
                User = impersonatedAccount,
                Scopes = scopes
            }.FromCertificate(certificate));
            return creds;
        }

        private static X509Certificate2 GetCertificateFromFile(string path = "key.p12", string pwd = "notasecret")
        {
            var certificate = new X509Certificate2(path, pwd, X509KeyStorageFlags.Exportable);
            return certificate;
        }
    }
}
