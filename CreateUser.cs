using System;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using Microsoft.Azure.Storage;
using Microsoft.Azure.Storage.Queue;
using Newtonsoft.Json;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    public static class CreateUser
    {
        [FunctionName("CreateUser")]
        public static async Task RunAsync(
            [QueueTrigger("UserRequestAccess")] UserInfo user,
            ILogger log)
        {
            
            log.LogInformation("C# HTTP trigger function processed a request.");
            IConfiguration config = new ConfigurationBuilder()

            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .AddEnvironmentVariables()
            .Build();

            string redirectLink = config["redirectLink"];
            string EmailWork = user.emailwork;
            string EmailCloud = user.emailcloud;
            string FirstName = user.firstname;
            string LastName = user.lastname;
            string RGCode = user.rgcode;

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            log.LogInformation($"Create user {EmailCloud}");
            var createUser = await UserCreation(graphAPIAuth, EmailCloud, FirstName, LastName, redirectLink, log);

            if (String.Equals(createUser[0], "Invitation error"))
            {
                 throw new SystemException(createUser[0]);
            }
            else
            {
                var userupdate = await updateUser(graphAPIAuth, createUser, RGCode, FirstName, LastName, log);
                if (userupdate)
                {
                    string EmailUser = String.Equals(EmailCloud, EmailWork) ? EmailCloud : EmailWork;
                    string ResponsQueue = "";

                    var connectionString = config["AzureWebJobsStorage"];

                    CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
                    CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                    CloudQueue queue = queueClient.GetQueueReference("sendemail");
                     
                    ResponsQueue = AddQueueEmail(queue, EmailUser, FirstName, LastName, createUser, log).GetAwaiter().GetResult();

                    if (String.Equals(ResponsQueue, "Queue create"))
                    {
                        log.LogInformation("Response queue");
                    }
                    else
                    {
                    log.LogInformation("Response queue error");
                        throw new SystemException(ResponsQueue);
                    }
                }
                else
                {
                    throw new SystemException("Error in user update");
                }
            }
        }

        public static async Task<List<string>> UserCreation(GraphServiceClient graphServiceClient, string emailcloud, string firstname, string lastname, string redirectLink, ILogger log)
        {
            List<string> InviteInfo = new List<string>();

                try
                {
                    var invitation = new Invitation
                    {
                        SendInvitationMessage = false,
                        InvitedUserEmailAddress = emailcloud,
                        InvitedUserType = "Member",
                        InviteRedirectUrl = redirectLink,
                        InvitedUserDisplayName = $"{firstname} {lastname}",
                    };

                    var userInvite = await graphServiceClient.Invitations.Request().AddAsync(invitation);
                    InviteInfo.Add("Invitation success");
                    InviteInfo.Add(userInvite.InvitedUser.Id);
                    InviteInfo.Add(userInvite.InviteRedeemUrl);

                    log.LogInformation(@"User invite successfully - {userInvite.InvitedUser.Id}");
                }
                catch (ServiceException ex)
                {
                    log.LogInformation($"Error Creating User Invite : {ex.Message}");
                    InviteInfo.Add("Invitation error");
                };

            return InviteInfo;
        }

        public static async Task<bool> updateUser(GraphServiceClient graphServiceClient, List<string> userID, string RGCode, string firstName, string lastName, ILogger log)
        {
            bool result = false;
            try
            {
                var guestUser = new User
                {
                    Department = RGCode,
                    UserType = "Member"
                };

                await graphServiceClient.Users[userID[1]].Request().UpdateAsync(guestUser);
                log.LogInformation("User update successfully");

                result = true;
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error Updating User : {ex.Message}");
                result = false;
            }
            return result;
        }

        static async Task<string> AddQueueEmail(CloudQueue theQueue,  string EmailUser, string FirstName, string LastName, List<string> userID, ILogger log)
        {
            string response = "";
            UserEmail email = new UserEmail();

            email.emailUser = EmailUser;
            email.firstname = FirstName;
            email.lastname = LastName;
            email.userid = userID;

            string serializedMessage = JsonConvert.SerializeObject(email);
            if (await theQueue.CreateIfNotExistsAsync())
            {
                log.LogInformation("The queue was created.");
            }

            CloudQueueMessage message = new CloudQueueMessage(serializedMessage);
            try
            {
                log.LogInformation("create queue");

                theQueue.AddMessage(message, initialVisibilityDelay: TimeSpan.FromMinutes(5));
                response = "Queue create";
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error in the queue {ex}");
                response = "Queue error";

            }

            return response;

        }
    }
}
