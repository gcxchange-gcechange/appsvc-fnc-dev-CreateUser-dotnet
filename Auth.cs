using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    class Auth
    {
        public GraphServiceClient graphAuth(ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            var scopes = new string[] { "https://graph.microsoft.com/.default" };
            var clientSecret = "m2.Oy7ws71T299-GqoB5lO~3Bj-V.6H_tH"; // Or some other secure place.
            var clientId = "4bea31f8-4b45-4faa-bc60-e5bffb898a37"; // Or some other secure place.
            var tenantid = "ddbd240e-11ba-47a6-abeb-e1a6be847a17";
            
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithTenantId(tenantid)
            .WithClientSecret(clientSecret)
            .Build();

            // Build the Microsoft Graph client. As the authentication provider, set an async lambda
            // which uses the MSAL client to obtain an app-only access token to Microsoft Graph,
            // and inserts this access token in the Authorization header of each API request. 
            GraphServiceClient graphServiceClient =
                new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) => {

                    // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                    var authResult = await confidentialClientApplication
                        .AcquireTokenForClient(scopes)
                        .ExecuteAsync();

                    // Add the access token in the Authorization header of the API request.
                    requestMessage.Headers.Authorization =
                        new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                })
                );
            return graphServiceClient;
        }
      
    }
}
