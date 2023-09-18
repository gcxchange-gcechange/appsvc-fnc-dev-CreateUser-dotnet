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

            string UserSender = config["userSender"];
            string recipientAddress = config["recipientAddress"];
            string ResponsQueue = "";

            string EmailWork = req.Query["EmailWork"];
            string EmailCloud = req.Query["EmailCloud"];
            string FirstName = req.Query["FirstName"];
            string LastName = req.Query["LastName"];
            string Department = req.Query["Department"];
            string B2B = req.Query["B2B"];
            string RGCode = req.Query["RGCode"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);

            EmailWork = EmailWork ?? data?.EmailWork;
            EmailCloud = EmailCloud ?? data?.EmailCloud;
            FirstName = FirstName ?? data?.FirstName;
            LastName = LastName ?? data?.LastName;
            Department = Department ?? data?.Department;
            B2B = B2B ?? data?.B2B;
            RGCode = RGCode ?? data?.RGCode;

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
                    var isSynced = (B2B == "YES");

                    if (isSynced)
                    {
                        sendEmail(recipientAddress, UserSender, EmailWork, Department, DateTime.Now, log);
                        ResponsQueue = $"{Department} already synced. User email:{EmailWork}";
                        return new BadRequestObjectResult(ResponsQueue);
                    }
                }

                var connectionString = config["AzureWebJobsStorage"];
                log.LogInformation($"Queue for {EmailCloud}");
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
                CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                CloudQueue queue = queueClient.GetQueueReference("userrequestaccess");

                ResponsQueue = InsertMessageAsync(queue, EmailCloud, EmailWork, FirstName, LastName, RGCode, log).GetAwaiter().GetResult();

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
       
        static async Task<string> InsertMessageAsync(CloudQueue theQueue, string emailcloud, string emailwork, string firstname, string lastname, string rgcode, ILogger log)
        {
            string response = "";
            UserInfo info = new UserInfo();

            info.emailcloud = emailcloud;
            info.emailwork = emailwork;
            info.firstname = firstname;
            info.lastname = lastname;
            info.rgcode = rgcode;

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

        public static async void sendEmail(string recipientAddress, string senderUserId, string userEmail, string userDept, DateTime incidentDate, ILogger log)
        {
            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            var format = "yyyy-MM-dd hh:mm:ss tt";

            var submitMsg = new Message
            {
                Subject = "GCX - Registration attempt from synced department",

                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = @$"<p>The following user from a synced department attempted to register for GCX:</p>
                                 Date: {incidentDate.ToString(format)}<br />
                                 User email: {userEmail}<br />
                                 User department: {userDept}<br />"
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                           Address = $"{recipientAddress}"
                        }
                    }
                },
            };

            try
            {
                await graphAPIAuth.Users[senderUserId].SendMail(submitMsg).Request().PostAsync();

            }
            catch (ServiceException e)
            {
                log.LogInformation($"Error: {e.Message}");
            }
        }
    }
}