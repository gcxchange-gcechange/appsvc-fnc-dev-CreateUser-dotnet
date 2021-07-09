using System;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;

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

            string Email = user.email;
            string FirstName = user.firstname;
            string LastName = user.lastname;
            string JobTitle = user.jobtitle;
            string Department = user.department;
            var domaine = "tbssctdev";
            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            var createUser = await UserCreation(graphAPIAuth, Email, FirstName, LastName, domaine, log);

            if (String.Equals(createUser[0], "Invitation error"))
            {
                 throw new SystemException(createUser[0]);
            }
            else
            {
                var userupdate = await updateUser(graphAPIAuth, createUser, JobTitle, Email, Department, log);
                if (userupdate)
                {
                    SendMail sendmail = new SendMail();
                    try
                    {
                        sendmail.send(graphAPIAuth, log, createUser, Email, "UserCreate");
                       //return new OkObjectResult("User mail with success");
                    }
                    catch (Exception ex)
                    {
                        throw new SystemException($"Error in invite user: {ex}");
                    }
                }
            }
            //throw new SystemException($"Something went wrong, please check the logs ");
        }

        public static async Task<List<string>> UserCreation(GraphServiceClient graphServiceClient, string email, string firstname, string lastname, string domaine, ILogger log)
        {
            List<string> InviteInfo = new List<string>();

            try
            {
                var invitation = new Invitation
                {
                    SendInvitationMessage = false,
                    InvitedUserEmailAddress = email,
                    InvitedUserType = "Member",
                    InviteRedirectUrl = $"https://{domaine}.sharepoint.com",
                    InvitedUserDisplayName = $"{firstname} {lastname}",
                };

                var userInvite = await graphServiceClient.Invitations.Request().AddAsync(invitation);
                InviteInfo.Add("Invitation success");
                InviteInfo.Add(userInvite.InvitedUser.Id);
                InviteInfo.Add(userInvite.InviteRedeemUrl);

                log.LogInformation("User invite successfully");
            }
            catch (ServiceException ex)
            {
                log.LogInformation($"Error Creating User Invite : {ex.Message}");
                InviteInfo.Add("Invitation error");
            };
            return InviteInfo;
        }

        public static async Task<bool> updateUser(GraphServiceClient graphServiceClient, List<string> userID, string jobTitle, string email, string department, ILogger log)
        {
            bool result = false;
            try
            {
                var guestUser = new User
                {
                    JobTitle = jobTitle,
                    Mail = email,
                    Department = department,
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

    }
}
