using Azure.Storage.Queues;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    public class CreateUser
    {
        private readonly IConfiguration _config;
        private readonly ILogger _log;

        public CreateUser(IConfiguration config, ILogger<CreateUser> log)
        {
            _config = config;
            _log = log;
        }

        [Function("CreateUser")]
        public async Task RunAsync(
            [QueueTrigger("UserRequestAccess")] UserInfo user,
            FunctionContext context)
        {
            _log.LogInformation("Processing user creation request.");

            string redirectLink = _config["redirectLink"];
            string EmailWork = user.emailwork;
            string EmailCloud = user.emailcloud;
            string FirstName = user.firstname;
            string LastName = user.lastname;
            string RGCode = user.rgcode;

            Auth auth = new Auth(_config, _log);
            var graphAPIAuth = auth.GetGraphClient();

            _log.LogInformation($"Creating user {EmailCloud}");
            var createUser = await UserCreation(graphAPIAuth, EmailCloud, FirstName, LastName, redirectLink);

            if (string.Equals(createUser[0], "Invitation error"))
                throw new Exception(createUser[0]);

            bool userUpdated = await UpdateUser(graphAPIAuth, createUser, RGCode, FirstName, LastName);
            if (!userUpdated)
                throw new Exception("Error in user update");


            string EmailUser = String.Equals(EmailCloud, EmailWork) ? EmailCloud : EmailWork;
            string connectionString = _config["AzureWebJobsStorage"];

            var queueClient = new QueueClient(connectionString, "sendemail");
            string queueResult = await AddQueueEmail(queueClient, EmailUser, FirstName, LastName, createUser);

            if (queueResult == "Queue create")
            {
                _log.LogInformation("Message added to response queue.");
            }
            else
            {
                _log.LogError("Failed to queue response.");
                throw new Exception(queueResult);
            }
        }

        private async Task<List<string>> UserCreation(GraphServiceClient graphServiceClient, string emailcloud, string firstname, string lastname, string redirectLink)
        {
            var inviteInfo = new List<string>();
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

                var userInvite = await graphServiceClient.Invitations.PostAsync(invitation);

                inviteInfo.Add("Invitation success");
                inviteInfo.Add(userInvite.InvitedUser.Id);
                inviteInfo.Add(userInvite.InviteRedeemUrl);

                _log.LogInformation($"User invited successfully - {userInvite.InvitedUser.Id}");
            }
            catch (ServiceException ex)
            {
                _log.LogError($"Error creating user invite: {ex.Message}");
                inviteInfo.Add("Invitation error");
            }

            return inviteInfo;
        }

        private async Task<bool> UpdateUser(GraphServiceClient graphServiceClient, List<string> userID, string rgCode, string firstName, string lastName)
        {
            try
            {
                var guestUser = new User
                {
                    Department = rgCode,
                    UserType = "Member"
                };

                await graphServiceClient.Users[userID[1]].PatchAsync(guestUser);
                _log.LogInformation("User updated successfully");
                return true;
            }
            catch (Exception ex)
            {
                _log.LogError($"Error updating user: {ex.Message}");
                return false;
            }
        }

        private async Task<string> AddQueueEmail(QueueClient queueClient, string emailUser, string firstName, string lastName, List<string> userID)
        {
            try
            {
                await queueClient.CreateIfNotExistsAsync();
                var email = new UserEmail
                {
                    emailUser = emailUser,
                    firstname = firstName,
                    lastname = lastName,
                    userid = userID
                };

                string message = JsonConvert.SerializeObject(email);
                await queueClient.SendMessageAsync(Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(message)),
                                                   visibilityTimeout: TimeSpan.FromMinutes(5));

                return "Queue create";
            }
            catch (Exception ex)
            {
                _log.LogError($"Error adding message to queue: {ex}");
                return "Queue error";
            }
        }
    }
}
