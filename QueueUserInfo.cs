using System.Net;
using System.Text;
using System.Text.Json;
using Azure.Storage.Queues;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    public class QueueUserInfo
    {
        private readonly IConfiguration _config;
        private readonly ILogger<QueueUserInfo> _log;

        public QueueUserInfo(IConfiguration config, ILogger<QueueUserInfo> log)
        {
            _config = config;
            _log = log;
        }

        [Function("QueueUserInfo")]
        public async Task<HttpResponseData> RunAsync(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req,
            FunctionContext context)
        {
            try
            {
                _log.LogInformation("Processing QueueUserInfo request");

                string UserSender = _config["delegatedUserName"];
                string recipientAddress = _config["recipientAddress"];
                string ResponsQueue = "";

                var query = System.Web.HttpUtility.ParseQueryString(req.Url.Query);
                string EmailWork = query["EmailWork"];
                string EmailCloud = query["EmailCloud"];
                string FirstName = query["FirstName"];
                string LastName = query["LastName"];
                string Department = query["Department"];
                string B2B = query["B2B"];
                string RGCode = query["RGCode"];

                string body = await new StreamReader(req.Body).ReadToEndAsync();
                var data = JsonSerializer.Deserialize<UserInfo>(body);

                EmailWork ??= data?.emailwork;
                EmailCloud ??= data?.emailcloud;
                FirstName ??= data?.firstname;
                LastName ??= data?.lastname;
                Department ??= data?.rgcode;
                B2B ??= data?.rgcode;
                RGCode ??= data?.rgcode;

                if (string.IsNullOrWhiteSpace(EmailCloud) || string.IsNullOrWhiteSpace(EmailWork)
                    || string.IsNullOrWhiteSpace(FirstName) || string.IsNullOrWhiteSpace(LastName))
                {
                    ResponsQueue = $"Missing field: {nameof(EmailCloud)}: {EmailCloud}, {nameof(EmailWork)}: {EmailWork}, {nameof(FirstName)}: {FirstName}, {nameof(LastName)}: {LastName}";
                    var badRequest = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                    await badRequest.WriteStringAsync(ResponsQueue);
                    return badRequest;
                }

                var graphClient = Auth.GetGraphClient(_log);

                if (await CheckUserExists(graphClient, EmailWork))
                {
                    ResponsQueue = $"{EmailWork} already registered";
                    var badRequest = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                    await badRequest.WriteStringAsync(ResponsQueue);
                    return badRequest;
                }

                if (B2B == "YES")
                {
                    await SendEmail(graphClient, recipientAddress, UserSender, EmailWork, Department, DateTime.UtcNow);
                    ResponsQueue = $"{Department} already synced. User email: {EmailWork}";
                    var badRequest = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                    await badRequest.WriteStringAsync(ResponsQueue);
                    return badRequest;
                }

                var queueClient = new QueueClient(_config["AzureWebJobsStorage"], "userrequestaccess");
                ResponsQueue = await InsertMessageAsync(queueClient, EmailCloud, EmailWork, FirstName, LastName, RGCode);

                var response = req.CreateResponse(
                    ResponsQueue == "Queue create" ? System.Net.HttpStatusCode.OK : System.Net.HttpStatusCode.BadRequest);

                await response.WriteStringAsync(ResponsQueue);
                return response;
            }
            catch (Exception ex)
            {
                _log.LogError(ex.Message);
                _log.LogError(ex.InnerException?.Message);
                _log.LogError(ex.StackTrace);

                var badResponse = req.CreateResponse(HttpStatusCode.BadRequest);
                await badResponse.WriteStringAsync(ex.Message);
                return badResponse;
            }
            
        }

        private async Task<string> InsertMessageAsync(QueueClient queueClient, string emailcloud, string emailwork, string firstname, string lastname, string rgcode)
        {
            var userInfo = new UserInfo
            {
                emailcloud = emailcloud,
                emailwork = emailwork,
                firstname = firstname,
                lastname = lastname,
                rgcode = rgcode
            };

            string serializedMessage = JsonSerializer.Serialize(userInfo);
            await queueClient.CreateIfNotExistsAsync();

            try
            {
                var base64Message = Convert.ToBase64String(Encoding.UTF8.GetBytes(serializedMessage));
                await queueClient.SendMessageAsync(base64Message);
                _log.LogInformation("Message added to queue");
                return "Queue create";
            }
            catch (Exception ex)
            {
                _log.LogError($"Queue error: {ex.Message}");
                return "Queue error";
            }
        }

        private async Task<bool> CheckUserExists(GraphServiceClient graphClient, string email)
        {
            try
            {
                var result = await graphClient.Users
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Filter = $"mail eq '{email.Replace("'", "''")}'";
                    });

                return result?.Value?.Count > 0;
            }
            catch (Exception ex)
            {
                _log.LogError($"Graph check error: {ex.Message}");
                return false;
            }
        }

        private async Task SendEmail(GraphServiceClient graphClient, string recipientAddress, string senderUserId, string userEmail, string userDept, DateTime incidentDate)
        {
            var message = new Message
            {
                Subject = "GCX - Registration attempt from synced department",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = @$"<p>The following user from a synced department attempted to register for GCX:</p>
                                Date: {incidentDate:yyyy-MM-dd hh:mm:ss tt}<br />
                                User email: {userEmail}<br />
                                User department: {userDept}<br />"
                },
                ToRecipients = new List<Recipient>
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipientAddress
                        }
                    }
                }
            };

            try
            {
                await graphClient.Users[senderUserId].SendMail.PostAsync(new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
                {
                    Message = message,
                    SaveToSentItems = false
                });
            }
            catch (Exception ex)
            {
                _log.LogError($"SendMail failed: {ex.Message}");
            }
        }
    }
}
