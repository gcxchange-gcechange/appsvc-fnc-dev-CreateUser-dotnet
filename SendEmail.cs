using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    class SendMail
    {
        public async void send(GraphServiceClient graphServiceClient, ILogger log, List<string> Redeem, string email, string UserSender)
        {
            var submitMsg = new Message();
            submitMsg = new Message
            {
                Subject = "Welcome",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = $"<a href='{Redeem[1]}'>Click here </a>"
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                           Address = $"{email}"
                        }
                    }
                },
            };
            try
            {
                await graphServiceClient.Users[UserSender]
                      .SendMail(submitMsg)
                      .Request()
                      .PostAsync();
                log.LogInformation($"User mail successfully {Redeem[2]}");

            }
            catch (ServiceException e)
            {
                log.LogInformation($"Error: {e.Message}");
            }
        }
    }
}
