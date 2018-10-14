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

            var request = service.Users.List();
            var response = request.Execute();
            var users = response.UsersValue;
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

        private static X509Certificate2 GetCertificateFromFile(string path = "key.12", string pwd = "notasecret")
        {
            var certificate = new X509Certificate2(path, pwd, X509KeyStorageFlags.Exportable);
            return certificate;
        }
    }
}
