using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    class Auth
    {
        public static GraphServiceClient GetGraphClient(ILogger log)
        {
            var config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: true)
                .AddEnvironmentVariables()
                .Build();

            var clientId = config["clientId"];
            var clientSecret = config["clientSecret"];
            var tenantId = config["tenantid"];
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            return new GraphServiceClient(credential, scopes);
        }
    }
}