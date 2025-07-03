using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    public class Auth
    {
        private readonly IConfiguration _config;
        private readonly ILogger _logger;

        public Auth(IConfiguration config, ILogger logger)
        {
            _config = config;
            _logger = logger;
        }

        public GraphServiceClient GetGraphClient()
        {
            _logger.LogInformation("Creating GraphServiceClient with TokenCredentialAuthProvider...");

            var clientId = _config["clientId"];
            var tenantId = _config["tenantid"];
            var clientSecret = _config["clientSecret"];
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            var clientSecretCredential = new ClientSecretCredential(
                tenantId,
                clientId,
                clientSecret
            );

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            return graphClient;
        }
    }
}

