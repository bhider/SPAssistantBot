using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SPAssistantBot.Services.Helpers
{
    public class OAuthMessageHandler : DelegatingHandler
    { 

        private readonly string _accessToken;
        public OAuthMessageHandler(string accessToken, HttpMessageHandler innerHandler) : base(innerHandler)
        {
            _accessToken = accessToken;
        }

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _accessToken);
            return base.SendAsync(request, cancellationToken);
        }    
    }
}
