using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace SPAssistantBot.Services.Helpers
{
    public class GraphServiceHttpClient : IDisposable
    {
        private static HttpClient httpClient;

        private bool disposedValue;

       
        public GraphServiceHttpClient(IConfiguration configuration, ILogger log)
        {
            Init(configuration);
            Log = log;
        }


        public async Task<JObject> ExecuteGetAsync(string url)
        {
            JObject body = null;

            Log.LogDebug($"Get request to {url}");

            try
            {
                var response = await httpClient.GetAsync(url);

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
                    Log.LogError($"Get request error: {response.ReasonPhrase}");
                    throw new Exception(response.ReasonPhrase);
                }
            }
            catch (Exception ex)
            {
                Log.LogError($"Get request exception: {ex.Message}");
                throw;
            }
            
            return body;
        }

        public async Task<JObject> ExecutePostAsync(string url, JObject content)
        {
            JObject body = null;

            Log.LogDebug($"Post request to {url}");

            try
            {
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
                    Log.LogError($"Post request error: {response.ReasonPhrase}");
                    throw new Exception(response.ReasonPhrase);
                }
            }
            catch (Exception ex)
            {
                Log.LogError($"Post request exception: {ex.Message}");
                throw;
            }

            return body;
        }

        public async Task<JObject> ExecutePatchAsync(string url, JObject content)
        {
            JObject body = null;

            Log.LogDebug($"Patch request to {url}");

            try
            {
                var response = await httpClient.PatchAsync(url, new StringContent(content.ToString(), Encoding.UTF8, "application/json"));

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
                    Log.LogError($"Patch request error: {response.ReasonPhrase}");
                    throw new Exception(response.ReasonPhrase);
                }
            }
            catch (Exception ex)
            {
                Log.LogError($"Patch request exception: {ex.Message}");
                throw;
            }

            return body;
        }

        public async Task<bool> ExecuteDeleteAsync(string url)
        {
            var response = await httpClient.DeleteAsync(url);

            Log.LogDebug($"Delete request to {url}");

            try
            {
                if (response.IsSuccessStatusCode && response.StatusCode == HttpStatusCode.NoContent)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log.LogError($"Delete request exception: {ex.Message}");
                throw;
            }

            return false;
        }

        public async Task<JObject> ExecuteLongPollingPostAsync(string url, JObject content)
        {
            JObject body = null;

            Log.LogDebug($"Post request to {url}");
            Log.LogInformation("Executing Long Polling operation....");

            try
            {
                var response = await httpClient.PostAsync(url, new StringContent(content.ToString(), Encoding.UTF8, "application/json"));

                if (response.IsSuccessStatusCode)
                {
                    if (response.StatusCode == System.Net.HttpStatusCode.Accepted)
                    {
                        var location = response.Headers.Location;
                        var monitorUrl = $"https://graph.microsoft.com/v1.0{location}";
                        var statusCode = HttpStatusCode.NotFound;
                        HttpResponseMessage opResponse = null;

                        while (statusCode != HttpStatusCode.OK)
                        {
                            opResponse = await httpClient.GetAsync(monitorUrl);
                            statusCode = opResponse.StatusCode;

                        }

                        var result = await opResponse.Content.ReadAsStringAsync();

                        if (!string.IsNullOrWhiteSpace(result))
                        {
                            body = JObject.Parse(result);
                        }
                    }

                }
                else
                {
                    Log.LogError($"Error executing long running operation: {response.ReasonPhrase}");
                    throw new Exception(response.ReasonPhrase);
                }
            }
            catch (Exception ex)
            {
                Log.LogError($"Long polling operation exception: {ex.Message}");
                throw;
            }

            return body;
        }

        public async Task<JObject> ExecutePutAsync(string url, JObject content)
        {
            JObject body = null;

            Log.LogDebug($"Put request to {url}");

            try
            {
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
                    Log.LogError($"Put request error: {response.ReasonPhrase}");
                    throw new Exception(response.ReasonPhrase);
                }
            }
            catch (Exception ex)
            {
                Log.LogError($"Put request exception: {ex.Message}");
                throw;
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

        private static void Init(IConfiguration configuration)
        {
            var aadApplicationId = configuration["AADClientId"];
            var aadApplicationSecret = configuration["AADClientSecret"];
            var spTenant = configuration["Tenant"];
            //var accessToken = await GetAccessToken(aadApplicationId, aadApplicationSecret, spTenant);

            httpClient = new HttpClient(new OAuthMessageHandler(aadApplicationId, aadApplicationSecret, spTenant, new HttpClientHandler()));


            //KVServiceCertIdentifier = configuration["KeyVaultSecretIdentifier"];
            //KVService = keyVaultService;
            //Log = log;
        }

        private ILogger Log { get; set; }

        private async Task<string> GetAccessToken(string applicationId, string applicationSecret, string tenant)
        {
            try
            {
                IConfidentialClientApplication clientApp = ConfidentialClientApplicationBuilder.Create(applicationId).WithClientSecret(applicationSecret).WithTenantId(tenant).Build();
                var authResult = await clientApp.AcquireTokenForClient(new string[] { "https://graph.microsoft.com/.default" }).ExecuteAsync();
                var accessToken = authResult.AccessToken;
                return accessToken;
            }
            catch (Exception ex)
            {
                Log.LogError($"Get Access token error: {ex.Message}");
                throw;
            }
        }


        public void Dispose()
        {
            Dispose(true);
        }
    }
}
