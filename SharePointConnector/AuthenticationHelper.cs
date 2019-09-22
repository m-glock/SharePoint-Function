using System;
using System.Diagnostics;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;

    class AuthenticationHelper
    {
        public static ConfidentialClientApplication IdentityAppOnlyApp = new ConfidentialClientApplication(Constants.ClientIdForAppAuthn, Constants.AuthorityUri, Constants.RedirectUriForAppAuthn, new ClientCredential(Constants.ClientSecret), new TokenCache(), new TokenCache());

        private static GraphServiceClient graphClient = null;
        public static string token;

        
        public static GraphServiceClient GetAuthenticatedClientForApp()
        {

            // Create Microsoft Graph client.
            try
            {
                graphClient = new GraphServiceClient(
                    "https://graph.microsoft.com/v1.0",
                    new DelegateAuthenticationProvider(
                        async (requestMessage) =>
                        {
                            token = await GetTokenForAppAsync();
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                        }));
                return graphClient;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Could not create a graph client: " + ex.Message);
            }


            return graphClient;
        }

        public static async Task<string> GetTokenForAppAsync()
        {
            AuthenticationResult authResult;
        authResult = await IdentityAppOnlyApp.AcquireTokenForClientAsync(new string[] { "https://graph.microsoft.com/.default" });
            return authResult.AccessToken;
        }
    }