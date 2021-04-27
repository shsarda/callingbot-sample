using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Net.Http.Headers;

namespace CallingMeetingBot.Utility
{
    public static class GraphUtils
    {
        public static GraphServiceClient CreateGraphServiceClient_DelegatedAuth(BotOptions botOptions, string tenantId)
        {
            string graphApiResource = botOptions.GraphApiResourceUrl;
            Uri microsoftLogin = new Uri(botOptions.MicrosoftLoginUrl);

            // The authority to ask for a token: your azure active directory.
            string authority = new Uri(microsoftLogin, tenantId).AbsoluteUri;
            AuthenticationContext authenticationContext = new AuthenticationContext(authority);
            ClientCredential clientCredential = new ClientCredential(botOptions.AppId, botOptions.AppSecret);            

            var authProvider = new DelegateAuthenticationProvider(
                async (requestMessage) =>
                {
                    AuthenticationResult authenticationResult = authenticationContext.AcquireTokenAsync(graphApiResource, clientCredential).Result;
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", authenticationResult.AccessToken);
                });
            GraphServiceClient graphClient = new(authProvider);

            return graphClient;
        }
    }
}
