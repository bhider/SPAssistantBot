using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace SPAssistantBot.Services.Helpers
{
    public class GraphServiceHttpClient : IDisposable
    {
        private readonly IConfiguration _configuration;
        
        private HttpClient httpClient;

        private bool disposedValue;
        public GraphServiceHttpClient(IConfiguration configuration, ILogger log)
        {
            _configuration = configuration;
            Log = log;
        }

        public async Task Init()
        {
            var aadApplicationId = _configuration["AADClientId"];
            var aadApplicationSecret = _configuration["AADClientSecret"];
            var spTenant = _configuration["Tenant"];
            var accessToken = await GetAccessToken(aadApplicationId, aadApplicationSecret, spTenant);
            
            httpClient = new HttpClient(new OAuthMessageHandler(accessToken, new HttpClientHandler()));
            
            
            //KVServiceCertIdentifier = configuration["KeyVaultSecretIdentifier"];
            //KVService = keyVaultService;
            //Log = log;
        }

        private ILogger Log { get;  set; }

        private async Task<string> GetAccessToken( string applicationId, string applicationSecret, string tenant)
        {
            IConfidentialClientApplication clientApp = ConfidentialClientApplicationBuilder.Create(applicationId).WithClientSecret(applicationSecret).WithTenantId(tenant).Build();
            var authResult = await clientApp.AcquireTokenForClient(new string[] { "https://graph.microsoft.com/.default"}).ExecuteAsync();
            var accessToken = authResult.AccessToken;
            return accessToken;
        }

        public async Task<JObject> ExecuteGet(string url)
        {
            JObject body = null;

            Log.LogInformation($"Executing Get - Request Url : {url}");

            var response = await httpClient.GetAsync(url);

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                
                if (!string.IsNullOrWhiteSpace(result)){
                    body = JObject.Parse(result);
                }
            }
            else
            {
                throw new Exception(response.ReasonPhrase);
            }
            
            return body;
        }

        public async Task<JObject> ExecutePostAsync(string url, JObject content)
        {
            JObject body = null;

            var response = await httpClient.PostAsync(url, new StringContent(content.ToString(), Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();

                if (!string.IsNullOrWhiteSpace(result))
                {
                    body = JObject.Parse(result);
                }
            }
            else
            {
                throw new Exception(response.ReasonPhrase);
            }

            return body;
        }

        public async Task<JObject> ExecutePutAsync(string url, JObject content)
        {
            JObject body = null;

            var response = await httpClient.PutAsync(url, new StringContent(content.ToString(), Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();

                if (!string.IsNullOrWhiteSpace(result))
                {
                    body = JObject.Parse(result);
                }
            }
            else
            {
                throw new Exception(response.ReasonPhrase);
            }

            return body;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposedValue)
            {
                return;
            }

            if (disposing)
            {
                httpClient.Dispose();
            }

            disposedValue = true;
        }

        public void Dispose()
        {
            Dispose(true);
        }
    }
}
