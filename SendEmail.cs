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
using System.Configuration;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System.Collections.Generic;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    class SendMail
    {
        public async void send(GraphServiceClient graphServiceClient, ILogger log, List<string> Redeem, string email, string type)
        {
            var submitMsg = new Message();
            switch (type)
            {
                case "UserCreate":
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
                    Console.WriteLine("Case 1");
                    break;
                case "ErrorDepart":
                     submitMsg = new Message
                     {
                         Subject = "Error department not allow",
                         Body = new ItemBody
                         {
                             ContentType = BodyType.Html,
                             Content = $"Depart not in list, pls reach out to <a href='{Redeem[1]}'>HelpDesk </a>"
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
                    Console.WriteLine("Case 2");
                    break;
                default:
                    submitMsg = new Message
                    {
                        Subject = "Something went wrong",
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Html,
                            Content = $"Something somewhere happen. Please contact our helpdesk"
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
                    Console.WriteLine("Default case");
                    break;
            }
            
            try
            {
                await graphServiceClient.Users["6d6e092d-e8c1-4c97-b899-8a2626b0fccc"]
                   .SendMail(submitMsg)
                   .Request()
                   .PostAsync();
                log.LogInformation("User mail successfully");

            }
            catch (ServiceException e)
            {
                log.LogInformation($"Error: {e.Message}");
            }
        }
    }
}
