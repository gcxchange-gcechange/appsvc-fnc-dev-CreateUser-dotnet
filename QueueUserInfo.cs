using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.WindowsAzure.Storage.Queue;
using Microsoft.WindowsAzure.Storage;
using Microsoft.Extensions.Configuration;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    public static class QueueUserInfo
    {
        [FunctionName("QueueUserInfo")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.System, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            IConfiguration config = new ConfigurationBuilder()

            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .AddEnvironmentVariables()
            .Build();

            string EmailWork = req.Query["EmailWork"];
            string EmailCloud = req.Query["EmailCloud"];
            string FirstName = req.Query["FirstName"];
            string LastName = req.Query["LastName"];
            string Department = req.Query["Department"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            EmailWork = EmailWork ?? data?.EmailWork;
            EmailCloud = EmailCloud ?? data?.EmailCloud;
            FirstName = FirstName ?? data?.FirstName;
            LastName = LastName ?? data?.LastName;
            Department = Department ?? data?.Department;

            string ResponsQueue = "";
            log.LogInformation($"Start process with {EmailCloud}");
            if (String.IsNullOrEmpty(EmailCloud) || String.IsNullOrEmpty(EmailWork) || String.IsNullOrEmpty(FirstName) || String.IsNullOrEmpty(LastName))
            {
                ResponsQueue = $"Missing field: {EmailCloud} - {EmailWork} - {FirstName} - {LastName}";
                return new BadRequestObjectResult(ResponsQueue);
            }
            else
            {
                // Check if user exists
                var userExists = await CheckUserExists(EmailWork, log);

                if(userExists)
                {
                    ResponsQueue = $"{EmailWork} already registered";
                    return new BadRequestObjectResult(ResponsQueue);
                }

                var connectionString = config["AzureWebJobsStorage"];
                log.LogInformation($"Queue for {EmailCloud}");
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
                CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                CloudQueue queue = queueClient.GetQueueReference("userrequestaccess");
              
                ResponsQueue = InsertMessageAsync(queue, EmailCloud, EmailWork, FirstName, LastName, Department,  log).GetAwaiter().GetResult();

                if (String.Equals(ResponsQueue, "Queue create"))
                {
                    log.LogInformation("Response queue");
                    return new OkObjectResult(ResponsQueue);
                }
                else
                {
                    log.LogInformation("Response queue error");

                    return new BadRequestObjectResult(ResponsQueue);
                }
            }
        }
       
        static async Task<string> InsertMessageAsync(CloudQueue theQueue, string emailcloud, string emailwork, string firstname, string lastname, string department, ILogger log)
        {
            string response = "";
            UserInfo info = new UserInfo();

            info.emailcloud = emailcloud;
            info.emailwork = emailwork;
            info.firstname = firstname;
            info.lastname = lastname;
            info.department = department;

            string serializedMessage = JsonConvert.SerializeObject(info);
            if (await theQueue.CreateIfNotExistsAsync())
            {
                log.LogInformation("The queue was created.");
            }

            CloudQueueMessage message = new CloudQueueMessage(serializedMessage);
            try
            {
                log.LogInformation("create queue");

                await theQueue.AddMessageAsync(message);
                response = "Queue create";
            }
            catch(Exception ex)
            {
                log.LogInformation($"Error in the queue {ex}");
                response = "Queue error";

            }

            return response;

        }

        static async Task<bool> CheckUserExists(string Email, ILogger log)
        {
            bool result = false;

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            var user = await graphAPIAuth.Users.Request().Filter($"mail eq '{Email}'").GetAsync();

            if(user.Count > 0)
            {
                result = true;
            }

            return result;
        }
    }
}
