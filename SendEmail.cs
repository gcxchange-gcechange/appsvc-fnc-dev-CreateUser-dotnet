using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    public class SendEmailQueueTrigger
    {
        private readonly IConfiguration _config;
        private readonly ILogger<SendEmailQueueTrigger> _log;

        public SendEmailQueueTrigger(IConfiguration config, ILogger<SendEmailQueueTrigger> log)
        {
            _config = config;
            _log = log;
        }

        [Function("SendEmailQueueTrigger")]
        public async Task RunAsync([QueueTrigger("sendemail")] UserEmail email)
        {
            _log.LogInformation("SendEmailQueueTrigger started");

            string userSender = _config["userSender"];
            string redirectLink = _config["redirectLink"];
            string welcomeGroup = _config["welcomeGroup"];
            string gcxAssigned = _config["gcxAssigned"];

            string emailUser = email.emailUser;
            string firstName = email.firstname;
            string lastName = email.lastname;
            List<string> userID = email.userid;

            var graphClient = Auth.GetGraphClient(_log);

            await SendEmailAsync(graphClient, emailUser, userSender, firstName, lastName, redirectLink);
            await AddUserToGroupsAsync(graphClient, userID, welcomeGroup, gcxAssigned);
        }

        private async Task SendEmailAsync(GraphServiceClient graphClient, string recipientEmail, string senderUserId, string firstName, string lastName, string redirectLink)
        {
            var message = new Message
            {
                Subject = "You now have access to GCXchange | Vous avez maintenant accés à GCÉchange",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = @$"
                        <div style='font-family: Helvetica'>
						<i>(La version française suit)</i> 
						<br><br>
						<b>Welcome to GC<b style='color: #1f9cf5'>X</b>change!</b>
						<br><br>
						Hi {firstName} {lastName},
						<br><br>
						We're happy to announce that you now have access to <a href='https://gcxgce.sharepoint.com/?gcxLangTour=en'>GCXchange</a> — the GC's new digital workspace and collaboration platform! No log-in or password is needed for GCXchange, since it uses a single sign-on from your government device.
						<br><br>
						<center><h2><a href='https://gcxgce.sharepoint.com/?gcxLangTour=en'>You can access GCXchange here</a></h2>
						<br>
						<b>Bookmark the above link to your personalized homepage, as well as to <a href='https://teams.microsoft.com/_?tenantId=f6a7234d-bc9b-4520-ad5f-70669c2c7a9c#/conversations/General?threadId=19:OXWdygF2pylAN26lrbZNN-GGzf8W9YEpe2EBawXtM0s1@thread.tacv2&ctx=channel'>GCXchange's MS Teams platform.</a></b></center>
						<br><br>
						GCXchange uses a combination of Sharepoint and MS Teams to allow users to collaborate across GC departments and agencies.
						<br><br>
						On the Sharepoint side of GCXchange you can:
						<ol>
						<li>read <a href='https://gcxgce.sharepoint.com/sites/news'>GC-wide news and stories</a></li>
						<li>join one of the many <a href='https://gcxgce.sharepoint.com/sites/Communities'>cross-departmental communities</a></li>
						<li>engage with thematic hubs that focus on issues relevant to the public service</li>
						<li>create a <a href='https://gcxgce.sharepoint.com/sites/Support/SitePages/Communities.aspx'>community</a> for interdepartmental collaboration with a dedicated page and Teams space</li>
						</ol>
						<br>
						On the Teams side of GCXchange you can engage with the communities you have joined, as well as co-autho documents and chat with colleagues in other departments and agencies. To learn how to switch between your departmental and GCXchange MS Teams accounts <a href='https://www.youtube.com/watch?v=71bULf1UqGw&list=PLWhPHFzdUwX98NKbSG8kyq5eW9waj3nNq&index=8'>watch a video tutorial</a> or <a href='https://gcxgce.sharepoint.com/sites/Support/SitePages/FAQ.aspx'>access the step-by-step guidance</a>.
						<br><br>
                        We want to hear from you! Please take a few minutes to respond to our <a href='https://can01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fquestionnaire.simplesurvey.com%2Ff%2Fl%2Fgcxchange-gcechange%3Fidlang%3DEN&data=05%7C01%7CJordana.Globerman%40tbs-sct.gc.ca%7C4e8c64422cfe447268d508da38fb0ae7%7C6397df10459540479c4f03311282152b%7C0%7C0%7C637884948107280711%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&sdata=Zk8xT7DoGA2pt48hYvplYfaTKWpY%2BjqJ7%2B60REgj3rE%3D&reserved=0'>survey</a> about the registration process.
                        <br><br>
						If you run into a problem or have a question, contact: <a href='mailto:support-soutien@gcx-gce.gc.ca'>support-soutien@gcx-gce.gc.ca</a>
						<br><br>
						Happy collaborating!
						<br><br>
						<hr>
						<br><br>
						<b>Bienvenue à GC<b style='color: #1f9cf5'>É</b>change!</b>
						<br><br>
						Bonjour {firstName} {lastName},
						<br><br>
						Nous sommes heureux de vous annoncer que vous avez maintenant accès à <a href='https://gcxgce.sharepoint.com/SitePages/fr/Home.aspx?gcxLangTour=fr'>GCÉchange</a>, la nouvelle plateforme de collaboration et de travail numérique du GC! Aucun nom d'utilisateur ni mot de passe n'est requis pour accéder à GCÉchange, puisque cette platforme est intégrée à la session unique que vous ouvrez à partir de votre appareil gouvernemental.
						<br><br>
						<center><h2><a href='https://gcxgce.sharepoint.com/SitePages/fr/Home.aspx?gcxLangTour=fr'>Vous pouvez accéder à GCÉchange ici</a></h2>
						<br>
						<b>Ajoutez le lien ci-dessus comme favori à votre page d'accueil personnalisée ainsi qu'à <a href='https://teams.microsoft.com/_?tenantId=f6a7234d-bc9b-4520-ad5f-70669c2c7a9c#/conversations/General?threadId=19:OXWdygF2pylAN26lrbZNN-GGzf8W9YEpe2EBawXtM0s1@thread.tacv2&ctx=channel'>la plateforme Microsoft Teams de gcéchange.</a></b></center>
						<br><br>
						GCÉchange utilise SharePoint et Teams pour permettre aux utilisateurs de collaborer avec l’ensemble des ministères et organismes du GC.
						<br><br>
						Du côté SharePoint de GCÉchange, vous pouvez :
						<ol>
						<li>lire <a href='https://gcxgce.sharepoint.com/sites/News/SitePages/fr/Home.aspx'></a>les nouvelles et les histoires du GC</li>
						<li>participer à l’une des nombreuses <a href='https://gcxgce.sharepoint.com/sites/Communities/SitePages/fr/Home.aspx'>participer à l’une des nombreuses collectivités interministérielles</a></li>
						<li>participer à des carrefours thématiques qui se concentrent sur ces enjeux pertinents pour la fonction publique</li>
						<li>créer une <a href='https://gcxgce.sharepoint.com/sites/Support/SitePages/fr/Communities.aspx'>collectivité</a> de collaboration interministérielle qui a sa page et son espace Teams</li>
						</ol>
						<br>
						Du côté Teams de GCÉchange, vous pouvez communiquer avec les collectivités desquelles vous êtes membre, corédiger des documents et clavarder avec des collègues d’autres ministères et organismes. Pour savoir comment passer d’un compte ministériel à un compte GCÉchange dans Teams, <a href='https://gcxgce.sharepoint.com/sites/Support/SitePages/fr/FAQ.aspx'>regardez un tutoriel vidéo ou accédez aux directives étape par étape.</a>
						<br><br>
                        Nous souhaitons connaître votre opinion! Veuillez prendre quelques minutes pour répondre à notre <a href='https://can01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fquestionnaire.simplesurvey.com%2Ff%2Fl%2Fgcxchange-gcechange%3Fidlang%3DEN&data=05%7C01%7CJordana.Globerman%40tbs-sct.gc.ca%7C4e8c64422cfe447268d508da38fb0ae7%7C6397df10459540479c4f03311282152b%7C0%7C0%7C637884948107280711%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&sdata=Zk8xT7DoGA2pt48hYvplYfaTKWpY%2BjqJ7%2B60REgj3rE%3D&reserved=0'>sondage</a> sur le processus d’inscription.
                        <br><br>
						Si vous avez un problème ou une question, écrivez à : <a href='mailto:support-soutien@gcx-gce.gc.ca'>support-soutien@gcx-gce.gc.ca</a>.
						<br><br>
						Bonne collaboration!
						</div>"
                },
                ToRecipients = new List<Recipient>
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipientEmail
                        }
                    }
                }
            };

            try
            {
                var sendMailRequest = new SendMailPostRequestBody
                {
                    Message = message,
                    SaveToSentItems = false
                };

                await graphClient.Users[senderUserId].SendMail.PostAsync(sendMailRequest);
                _log.LogInformation("User mail sent successfully.");
            }
            catch (Exception ex)
            {
                _log.LogError($"SendMail error: {ex.Message}");
            }
        }

        private async Task<bool> AddUserToGroupsAsync(GraphServiceClient graphClient, List<string> userID, string welcomeGroupId, string gcxAssignedGroupId)
        {
            try
            {
                var reference = new ReferenceCreate
                {
                    OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{userID[1]}"
                };

                await graphClient.Groups[gcxAssignedGroupId].Members.Ref.PostAsync(reference);
                _log.LogInformation("User added to GCX_Assigned group.");

                // Optional: Add to welcome group too
                if (!string.IsNullOrEmpty(welcomeGroupId))
                {
                    await graphClient.Groups[welcomeGroupId].Members.Ref.PostAsync(reference);
                    _log.LogInformation("User added to Welcome group.");
                }

                return true;
            }
            catch (Exception ex)
            {
                _log.LogError($"Group membership error: {ex.Message}");
                return false;
            }
        }
    }
}