using System;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using System.Linq;
using System.Security;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;

namespace ReportsData
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Retrieving SharePoint Usage Data ...");

            // todo: fill out 
            ConfigurationModel configurationModel = new ConfigurationModel
            {
                AuthenticationMode = "MasterUser",
                AuthorityUri = "https://login.microsoftonline.com/organizations/",
                ClientId = "ClientID",
                ClientSecret = "",
                PbiPassword = "Password",
                PbiUsername = "pbiMainUser@xxxxxx.onmicrosoft.com",
                ReportId = "",
                Scope = new string[] { "https://reports.office.com/user_impersonation" },
                TenantId = "TenantID",
                WorkspaceId = ""
            };


            string accessToken = GetAccessToken(configurationModel);

            Console.WriteLine("Access Token Retrieved!");

            Console.WriteLine(accessToken);

            string jsonString = GetSharePointUsageData(accessToken);

            TenantSharePointUsageModel spum = JsonSerializer.Deserialize<TenantSharePointUsageModel>(jsonString);

            Console.WriteLine("Reports Data Retrieved!");

            Console.WriteLine("Done!");
        }


        private static string GetAccessToken(ConfigurationModel configuration) {

            // Logic for retrieving the access token 

            AuthenticationResult authenticationResult = null;
            
            // Create a public client to authorize the app with the AAD app
            IPublicClientApplication clientApp = PublicClientApplicationBuilder.Create(configuration.ClientId).WithAuthority(configuration.AuthorityUri).Build();
            var userAccounts = clientApp.GetAccountsAsync().Result;
            try
            {
                // Retrieve Access token from cache if available
                authenticationResult = clientApp.AcquireTokenSilent(configuration.Scope, userAccounts.FirstOrDefault()).ExecuteAsync().Result;
            }
            catch (MsalUiRequiredException)
            {
                try
                {
                    SecureString password = new SecureString();
                    foreach (var key in configuration.PbiPassword)
                    {
                        password.AppendChar(key);
                    }
                    authenticationResult = clientApp.AcquireTokenByUsernamePassword(configuration.Scope, configuration.PbiUsername, password).ExecuteAsync().Result;
                }
                catch (MsalException)
                {
                    throw;
                }
            }

            try
            {
                return authenticationResult.AccessToken;
            }
            catch (Exception)
            {
                throw;
            }
        }


        private static string GetSharePointUsageData(string accessToken) {

            // Logic for retrieving the reports data using accesstoken retrieved

            HttpClient client = new HttpClient();

            client.BaseAddress = new Uri("https://reports.office.com");
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));

            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {accessToken}");

            var result = Task.Run(async () => await client.GetAsync("/pbi/v1.0/f543e2b1-0a5e-4c64-9596-24a9d40c26ed/TenantSharePointUsage").ConfigureAwait(false)).Result;
            var data = Task.Run(async () => await result.Content.ReadAsStringAsync().ConfigureAwait(false)).Result;

            Console.WriteLine(data);

            return data;


        }
    }
}
