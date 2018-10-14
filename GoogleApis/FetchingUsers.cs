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
    /// <summary>
    /// Demonstrating the basic features of G-Suite Admin Directory API
    /// </summary>
    static class FetchingUsers
    {
        private static string _serviceAccount = "groupid@adminproject-219211.iam.gserviceaccount.com";
        private static string _impersonatedUser = "umair@bibelotz.com";

        /// <summary>
        /// Public method to run the private methods
        /// </summary>
        public static void Run()
        {
            AccountFetchAll();
        }

        /// <summary>
        /// Fetches all the users for the service account
        /// </summary>
        private static void AccountFetchAll()
        {
            X509Certificate2 certificate = GetCertificateFromFile();
            string[] scopes = new[] { DirectoryService.Scope.AdminDirectoryUserReadonly };
            ServiceAccountCredential creds = GetServiceAccountCredentials(_serviceAccount, _impersonatedUser, scopes, certificate);
            DirectoryService service = GetDirectoryService(creds);
            IEnumerable<User> userListRaw = GetUsersListRaw(service);
            //List<MailBoxUser> userModelList = GetMailBoxUsers(userListRaw);
            //PrintMailBoxUsers(userModelList);      
            Dictionary<string, bool> mailboxDictionary = GetMailboxDictionary(userListRaw);
            PrintMailboxDictionary(mailboxDictionary);
        }

        /// <summary>
        /// Prints the mailbox users stored in the list
        /// </summary>
        /// <param name="list">The list of mailbox to print</param>
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

        /// <summary>
        /// Prints all the emails by removing the duplicates
        /// </summary>
        /// <param name="dict">dictionary of emails</param>
        private static void PrintMailboxDictionary(Dictionary<string, bool> dict)
        {
            Console.WriteLine($" EMAIL ADDRESSES FOUND: {dict.Count}:");
            foreach (var item in dict)
            {
                Console.WriteLine($"\t\t{item.Key}, {item.Value}");
            }
        }

        /// <summary>
        /// Gets the dictionary of emails from the raw email list.
        /// </summary>
        /// <param name="userListRaw">Raw list returned by executing the List request on DirectoryRequest</param>
        /// <returns>
        /// A dictionary containing the email address as the key and bool as value representing the type of email. If the value is true,
        /// it means the email address is either the primary email or a non-editable alias. The false value represents that email is
        /// neither the primary email nor a non-editable alias. Note that the duplicate emails are not removed in this dictionary.
        /// </returns>
        private static Dictionary<string, bool> GetMailboxDictionary(IEnumerable<User> userListRaw)
        {
            var dict = new Dictionary<string, bool>();
            foreach (var user in userListRaw)
            {
                if (user.PrimaryEmail != null && !dict.ContainsKey(user.PrimaryEmail))
                {
                    dict.Add(user.PrimaryEmail, true);
                }

                if (user.NonEditableAliases != null)
                {
                    foreach (var item in user.NonEditableAliases)
                    {
                        if (!dict.ContainsKey(item))
                        {
                            dict.Add(item, true);
                        }
                    }
                }

                if (user.Aliases != null)
                {
                    foreach (var item in user.Aliases)
                    {
                        if(!dict.ContainsKey(item))
                        {
                            dict.Add(item, false);
                        }
                    } 
                }

                if(user.Emails != null)
                {
                    foreach (var item in user.Emails)
                    {
                        if(!dict.ContainsKey(item.Address))
                        {
                            dict.Add(item.Address, false);
                        }
                    }
                }
            }
            return dict;
        }


        /// <summary>
        /// Gets the list of MailBoxUser from the raw email list.
        /// </summary>
        /// <param name="userListRaw">Raw list returned by executing the List request on DirectoryRequest</param>
        /// <returns>List of MailBoxUser. Note that the individual list objects may contain duplicate emails.</returns>
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

        /// <summary>
        /// Executes the List request on the given DirectoryService and returns the UsersValue response.
        /// </summary>
        /// <param name="service">DirectoryService on which the request is to be executed</param>
        /// <param name="customerId">The customer id to user for request</param>
        /// <returns></returns>
        private static IEnumerable<User> GetUsersListRaw(DirectoryService service, string customerId = "my_customer")
        {
            var request = service.Users.List();
            request.Customer = customerId;
            var response = request.Execute();
            var users = response.UsersValue;
            return users;
        }

        /// <summary>
        /// Gets the DirectoryService from the given credentials and application name.
        /// </summary>
        /// <param name="credentials">ServiceAccountCredentials to use for the service</param>
        /// <param name="appName">Application name to use for the service</param>
        /// <returns></returns>
        private static DirectoryService GetDirectoryService(ServiceAccountCredential credentials, string appName = "Test Application")
        {
            var service = new DirectoryService(new BaseClientService.Initializer()
            {
                ApplicationName = appName,
                HttpClientInitializer = credentials
            });
            return service;
        }

        /// <summary>
        /// Gets the ServiceAccountCredential from the given service account, account to impersonate, scopes, and the certificate
        /// </summary>
        /// <param name="serviceAccount">Service account name</param>
        /// <param name="impersonatedAccount">Account that the service account will impersonate</param>
        /// <param name="scopes">Scopes to specify for the credentials</param>
        /// <param name="certificate">PKCS12 private key</param>
        /// <returns></returns>
        private static ServiceAccountCredential GetServiceAccountCredentials(string serviceAccount, string impersonatedAccount, string[] scopes, X509Certificate2 certificate)
        {
            var creds = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(serviceAccount)
            {
                User = impersonatedAccount,
                Scopes = scopes
            }.FromCertificate(certificate));
            return creds;
        }

        /// <summary>
        /// Creates the certificate from PKCS12 file
        /// </summary>
        /// <param name="path">Path of the PKCS12 file</param>
        /// <param name="pwd">Password to read PKCS12 file</param>
        /// <returns></returns>
        private static X509Certificate2 GetCertificateFromFile(string path = "key.p12", string pwd = "notasecret")
        {
            var certificate = new X509Certificate2(path, pwd, X509KeyStorageFlags.Exportable);
            return certificate;
        }
    }
}
