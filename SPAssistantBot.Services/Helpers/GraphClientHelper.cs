using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;

namespace SPAssistantBot.Services.Helpers
{
    static class GraphClientHelper
    {
        internal static GraphServiceClient GetGraphServiceClient(string aadApplicationId, string aadClientSecret, string spTenant)
        {
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder.Create(aadApplicationId)
                                                                                                                                                                          .WithTenantId(spTenant)
                                                                                                                                                                          .WithClientSecret(aadClientSecret)
                                                                                                                                                                          .Build();
            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);

            GraphServiceClient graphClient = new GraphServiceClient(authProvider);
            return graphClient;
        }
    }
}
