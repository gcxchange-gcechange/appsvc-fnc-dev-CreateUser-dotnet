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
using Microsoft.Graph;
using System.Collections.Generic;

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

            string listId = config["listId"];
            string siteId = config["siteId"];

            string EmailWork = req.Query["EmailWork"];
            string EmailCloud = req.Query["EmailCloud"];
            string FirstName = req.Query["FirstName"];
            string LastName = req.Query["LastName"];
            string Department = req.Query["Department"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            string ResponsQueue = "";
            dynamic data = JsonConvert.DeserializeObject(requestBody);

            EmailWork = EmailWork ?? data?.EmailWork;
            EmailCloud = EmailCloud ?? data?.EmailCloud;
            FirstName = FirstName ?? data?.FirstName;
            LastName = LastName ?? data?.LastName;
            Department = Department ?? data?.Department;

            log.LogInformation($"Start process with {EmailCloud}");
            if (String.IsNullOrEmpty(EmailCloud) || String.IsNullOrEmpty(EmailWork) || String.IsNullOrEmpty(FirstName) || String.IsNullOrEmpty(LastName))
            {
                ResponsQueue = $"Missing field: {nameof(EmailCloud)}: {EmailCloud} - {nameof(EmailWork)}: {EmailWork} - {nameof(FirstName)}: {FirstName} - {nameof(LastName)}: {LastName}";
                return new BadRequestObjectResult(ResponsQueue);
            }
            else
            {
                // Check if user exists
                var userExists = await CheckUserExists(EmailWork, log);

                if (userExists)
                {
                    ResponsQueue = $"{EmailWork} already registered";
                    return new BadRequestObjectResult(ResponsQueue);
                }
                else
                {
                    // Check if department has been synced
                    var isSynced = await CheckB2BSync(siteId, listId, Department, log);

                    if (isSynced)
                    {
                        ResponsQueue = $"{Department} already synced. User email:{EmailWork}";
                        return new BadRequestObjectResult(ResponsQueue);
                    }

                }

                var connectionString = config["AzureWebJobsStorage"];
                log.LogInformation($"Queue for {EmailCloud}");
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
                CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                CloudQueue queue = queueClient.GetQueueReference("userrequestaccess");

                ResponsQueue = InsertMessageAsync(queue, EmailCloud, EmailWork, FirstName, LastName, Department, log).GetAwaiter().GetResult();

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

            var user = await graphAPIAuth.Users.Request().Filter($"mail eq '{Email.Replace("'", "''")}'").GetAsync();

            if(user.Count > 0)
            {
                result = true;
            }

            return result;
        }

        static async Task<bool> CheckB2BSync(string siteId, string listId, string RGCode, ILogger log)
        {
            bool result = false;

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            try
            {
                var queryOptions = new List<Option>()
                {
                    new QueryOption("expand", "fields(select=Id, Legal_x0020_Title, RG_x0020_Code, B2B)"),

                    //new HeaderOption("ConsistencyLevel", "eventual"),
                    //new QueryOption("$filter", "RG_x0020_Code eq '19'")
                    //new QueryOption("count", "true")
                };

                var deptItems = await graphAPIAuth.Sites[siteId].Lists[listId].Items.Request(queryOptions).GetAsync();

                foreach (var dept in deptItems)
                {
                    if (dept.Fields.AdditionalData.ContainsKey("RG_x0020_Code") && dept.Fields.AdditionalData["RG_x0020_Code"].ToString() == RGCode)
                    {
                        if (dept.Fields.AdditionalData.ContainsKey("B2B") && dept.Fields.AdditionalData["B2B"].ToString() == "YES")
                        {
                            result = true;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                log.LogError($"e.Message = {e.Message}");
                if (e.InnerException != null) log.LogError($"e.InnerException = {e.InnerException}");
            }
            return result;
        }










    }
}
