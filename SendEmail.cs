using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    public static class SendEmailQueueTrigger
    {
        [FunctionName("SendEmailQueueTrigger")]
        public static async Task RunAsync(
            [QueueTrigger("sendemail")] UserEmail email,
            ILogger log)
        {
            IConfiguration config = new ConfigurationBuilder()

            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .AddEnvironmentVariables()
            .Build();

            log.LogInformation("C# HTTP trigger function processed a request.");
            string UserSender = config["userSender"];
            string redirectLink = config["redirectLink"];
            string EmailUser = email.emailUser;
            string FirstName = email.firstname;
            string LastName = email.lastname;
            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            sendEmail(graphAPIAuth, EmailUser, UserSender, FirstName, LastName, redirectLink, log);
        }
        public static async void sendEmail(GraphServiceClient graphServiceClient, string email, string UserSender, string FirstName, string LastName, string redirectLink, ILogger log)
        {
            var submitMsg = new Message();
            submitMsg = new Message
            {
                Subject = "You're in! | Vous s'y êtes",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = @$"
                        (La version française suit)

                        Hi {FirstName} {LastName},<br><br>

                        We’re happy to announce that you now have access to gcxchange – the Government of Canada’s new digital workspace and modern intranet.<br><br>


                        Currently, there are two ways to use gcxchange: <br><br>

                        <ol><li><strong>Read articles, create and join GC-wide communities through your personalized homepage. Don’t forget to bookmark it: <a href='https://gcxgce.sharepoint.com/'>gcxgce.sharepoint.com/</a></strong></li>

                        <li><strong>Chat, call, and co-author with members of your communities using your Microsoft Teams and seamlessly toggle between gcxchange and your departmental environment. <a href='https://teams.microsoft.com/_?tenantId=f6a7234d-bc9b-4520-ad5f-70669c2c7a9c#/conversations/General?threadId=19:OXWdygF2pylAN26lrbZNN-GGzf8W9YEpe2EBawXtM0s1@thread.tacv2&ctx=channel'>Click here to find out how!</a></strong></li></ol>

                        We want to hear from you! Please take a few minutes to respond to our <a href=' https://questionnaire.simplesurvey.com/f/l/gcxchange-gcechange?idlang=EN'>survey</a> about the registration process.<br><br>

                        If you run into any issues along the way, please reach out to the support team at: <a href='mailto:support-soutien@gcx-gce.gc.ca'>support-soutien@gcx-gce.gc.ca</a><br><br>
                        
                        ---------------------------------------------------------------------------------<br><br>

                        (The English version precedes)<br><br>

                        Bonjour {FirstName} {LastName},<br><br>

                        Nous sommes heureux de vous annoncer que vous avez maintenant accès à gcéchange – le nouvel espace de travail numérique et intranet moderne du gouvernement du Canada.<br><br>


                        À l’heure actuelle, il y a deux façons d’utiliser gcéchange : <br><br>

                        <ol><li><strong>Lisez des articles, créez des communautés pangouvernementales et joignez-vous à celles-ci au moyen de votre page d’accueil personnalisée. N’oubliez pas d’ajouter cet espace dans vos favoris : <a href='https://gcxgce.sharepoint.com/'>gcxgce.sharepoint.com/</a></strong></li>

                        <li><strong>Clavardez et corédigez des documents avec des membres de vos communautés ou appelez ces membres au moyen de Microsoft Teams et passez facilement de gcéchange à votre environnement ministériel. <a href='https://teams.microsoft.com/_?tenantId=f6a7234d-bc9b-4520-ad5f-70669c2c7a9c#/conversations/General?threadId=19:OXWdygF2pylAN26lrbZNN-GGzf8W9YEpe2EBawXtM0s1@thread.tacv2&ctx=channel'>Cliquez ici pour savoir comment faire.</a></strong></li></ol>

                        Nous souhaitons connaître votre opinion! Veuillez prendre quelques minutes pour répondre à notre <a href='https://questionnaire.simplesurvey.com/f/l/gcxchange-gcechange?idlang=FR'>sondage</a> sur le processus d’inscription.<br><br>


                        Si vous éprouvez des problèmes en cours de route, veuillez communiquer avec l’équipe de soutien à l’adresse suivante : <a href='mailto:support-soutien@gcx-gce.gc.ca'>support-soutien@gcx-gce.gc.ca</a>"
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
                log.LogInformation($"User mail successfully {redirectLink}");

            }
            catch (ServiceException e)
            {
                log.LogInformation($"Error: {e.Message}");
            }
        }
    }
}
