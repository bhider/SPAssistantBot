using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace SPAssistantBot.Services
{
    class GraphServiceHelper
    {
        private readonly string accessToken = string.Empty;
        private HttpClient httpClient = null;
        private readonly JsonSerializerSettings jsonSettings =
           new JsonSerializerSettings { ContractResolver = new CamelCasePropertyNamesContractResolver() };
        private static readonly string teamsEndpoint = Environment.GetEnvironmentVariable("TeamsEndpoint");

        private ILogger logger = null;

        public GraphServiceHelper(string accessToken, ILogger log = null)
        {
            this.accessToken = accessToken;
            httpClient = new HttpClient();

            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", accessToken);
            httpClient.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));

            logger = log;
        }

        private async Task<HttpResponseMessage> MakeGraphCall(HttpMethod method, string uri, object body = null, int retries = 0)
        {
            string version = string.IsNullOrEmpty(teamsEndpoint) ? "beta" : teamsEndpoint;
            // Initialize retry delay to 3 secs
            int retryDelay = 3;

            string payload = string.Empty;

            if (body != null && (method != HttpMethod.Get || method != HttpMethod.Delete))
            {
                // Serialize the body
                payload = JsonConvert.SerializeObject(body, jsonSettings);
            }

            if (logger != null)
            {
                logger.LogInformation($"MakeGraphCall Request: {method} {uri}");
                logger.LogInformation($"MakeGraphCall Payload: {payload}");
            }

            do
            {
                var request = new HttpRequestMessage(method, $"{Constants.GraphResource}{version}{uri}");
                if (method.Method.ToUpper() == "PATCH")
                {
                    //httpClient.DefaultRequestHeaders.Add("X-HTTP-Method-Override", "PATCH");
                    request = new HttpRequestMessage(method, $"{Constants.GraphResource}v1.0{uri}");
                    //request.Headers.Add("Content-type", "application/json");
                    //method = HttpMethod.Post;
                }
                // Create the request



                if (!string.IsNullOrEmpty(payload))
                {
                    request.Content = new StringContent(payload);//, Encoding.UTF8, "application/json");
                }



                // Send the request
                var response = await httpClient.SendAsync(request, HttpCompletionOption.ResponseContentRead);

                if (!response.IsSuccessStatusCode)
                {
                    if (logger != null)
                        logger.LogInformation($"MakeGraphCall Error: {response.StatusCode}");
                    if (retries > 0)
                    {
                        if (logger != null)
                            logger.LogInformation($"MakeGraphCall Retrying after {retryDelay} seconds...({retries} retries remaining)");
                        Thread.Sleep(retryDelay * 1000);
                        // Double the retry delay for subsequent retries
                        retryDelay += retryDelay;
                    }
                    else
                    {
                        // No more retries, throw error
                        var error = await response.Content.ReadAsStringAsync();
                        throw new Exception(error);
                    }
                }
                else
                {
                    var r = await response.Content.ReadAsStringAsync();
                    return response;
                }
            }
            while (retries-- > 0);

            return null;
        }
    }
}
