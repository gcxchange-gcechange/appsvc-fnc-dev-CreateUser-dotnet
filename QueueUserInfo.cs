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
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            IConfiguration config = new ConfigurationBuilder()

            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .AddEnvironmentVariables()
            .Build();

            string Email = req.Query["Email"];
            string FirstName = req.Query["FirstName"];
            string LastName = req.Query["LastName"];
            string JobTitle = req.Query["JobTitle"];
            string Department = req.Query["Department"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            Email = Email ?? data?.Email;
            FirstName = FirstName ?? data?.FirstName;
            LastName = LastName ?? data?.LastName;
            JobTitle = JobTitle ?? data?.JobTitle;
            Department = Department ?? data?.Department;

            string ResponsQueue = "";

            if (String.IsNullOrEmpty(Email) || String.IsNullOrEmpty(FirstName) || String.IsNullOrEmpty(LastName))
            {
                ResponsQueue = "Missing field";
                return new BadRequestObjectResult(ResponsQueue);
            }
            else
            {
                var connectionString = config["AzureWebJobsStorage"];
                log.LogInformation($"Queue for {Email}");
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
                CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                CloudQueue queue = queueClient.GetQueueReference("userrequestaccess");
              
                ResponsQueue = InsertMessageAsync(queue,  Email, FirstName, LastName, JobTitle, Department,  log).GetAwaiter().GetResult();

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
       
        static async Task<string> InsertMessageAsync(CloudQueue theQueue, string email, string firstname, string lastname, string jobtitle, string department, ILogger log)
        {
            string response = "";
            UserInfo info = new UserInfo();

            info.email = email;
            info.firstname = firstname;
            info.lastname = lastname;
            info.jobtitle = jobtitle;
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
    }
}
