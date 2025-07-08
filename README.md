# appsvc-fnc-dev-CreateUser-dotnet
New registration process

This function work with a form in a webpage.
The function will receive data from the form and will create the user into the tenant, update the user info (and change the user type at the same time), add the user to our welcome group than send a custom email invitation to the user.
The webpage gater some information are EmailCloud, EmailWork, FirstName, LastName and Department. 

Step1. Check if all required information are receive.<br>
Step2. Create the queue with all information.<br>
Step3. Listen to that queue.  Start by creating the invitation.<br>
Step4. Update the user.<br>
Step5. Add the user to our group.<br>
Step6. Send email to the user with redeem link from step 3.<br>

## Required setting

clientId = App configuration client id<br>
secretId = App configuration client secret<br>
tenantId = Tenant id<br>
welcomeGroup = Id of the group<br>
redirectLink = Link that redirect the user when click on the link (Teams for us)<br>
delegatedUserName = the email address of the delegated user for graph calls/sending email<br>
delegatedUserSecret = the secret name in the key vault which contains the password for the delegatedUserName user<br>

