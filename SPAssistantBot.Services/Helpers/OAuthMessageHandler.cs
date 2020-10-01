using Microsoft.Identity.Client;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace SPAssistantBot.Services.Helpers
{
    public class OAuthMessageHandler : DelegatingHandler
    {
        private readonly string _applicationId;
        private readonly string _applicationSecret;
        private readonly string _tenant;

        
        public OAuthMessageHandler(string applicationId, string applicationSecret, string tenant, HttpMessageHandler innerHandler) : base(innerHandler)
        {
            _applicationId = applicationId;
            _applicationSecret = applicationSecret;
            _tenant = tenant;
        }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var accessToken = await GetAccessToken();
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            return await base.SendAsync(request, cancellationToken);
        }

        private async Task<string> GetAccessToken()
        {
            var clientApp = ConfidentialClientApplicationBuilder.Create(_applicationId).WithClientSecret(_applicationSecret).WithTenantId(_tenant).Build();
            var authResult = await clientApp.AcquireTokenForClient(new string[] { "https://graph.microsoft.com/.default" }).ExecuteAsync();
            var accessToken = authResult.AccessToken;
            return accessToken;
            
        }
    }
}
