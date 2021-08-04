using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    class SendMail
    {
        public async void send(GraphServiceClient graphServiceClient, ILogger log, List<string> Redeem, string email, string UserSender, string FirstName, string LastName)
        {
            var submitMsg = new Message();
            submitMsg = new Message
            {
                Subject = "You're in! | Vous s'y êtes",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = @$"
                        La version française suit<br><br>
                        Hi {FirstName} {LastName},<br><br>
                        We’re happy to announce that you now have access to gcxchange!<br><br>
                        <strong>But don't click it yet, the system needs a few minutes (5 minutes tops) to complete the activation.</strong> Patience is a virtue, and the reward is that when you finally do click the link, you'll be in right away.<br><br>
                        <a href='{Redeem[2]}'>Click here to get started!</a><br><br>
                        Can’t wait to explore the  new platform but don’t know where to start? We’ve got you covered. Check out the <a href='https://gcxgce.sharepoint.com/sites/support'>support centre</a> for guidance and tips on how to navigate gcxchange and explore the full potential of this exciting new tool.<br><br>  

                        ---------------------------------------------------------------------------------<br><br>

                        Bonjour {FirstName} {LastName},<br><br>
                        Nous sommes heureux de vous annoncer que vous avez maintenant accès à gcéchange!<br><br>
                        <strong>Toutefois, attendez avant de cliquer — le système a besoin de quelques minutes (au plus cinq minutes) pour terminer l’activation.</strong> La patience est mère de toutes les vertus, et la récompense, c’est que lorsque vous cliquerez finalement sur le lien, vous y serez déjà.<br><br>
                        <a href='{Redeem[2]}'>Cliquez ici pour commencer!</a><br><br>
                        Êtes-vous impatient d’exp lorer la nouvelle plateforme, mais ne savez pas par où commencer? Nous sommes là pour vous guider. Consultez le <a href='https://gcxgce.sharepoint.com/sites/support/sitepages/fr/home.aspx'>centre de soutien</a> pour obtenir des conseils et des astuces sur la façon de naviguer dans gcéchange et d’exploiter le plein potentiel de ce nouvel outil. "
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
