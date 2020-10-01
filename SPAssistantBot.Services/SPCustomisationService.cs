using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

using System.Threading.Tasks;

namespace SPAssistantBot.Services
{
    public class SPCustomisationService
    {
        private readonly ILogger _log;

        public SPCustomisationService(ILogger<SPCustomisationService> log)
        {
            _log = log;
        }
        public  async Task<bool> CustomiseAsync(string templateSiteUrl, string targetSiteUrl)
        {
            _log.LogInformation("Calling function to customise site");

            var success = false;

            var customisationServiceUrl = Environment.GetEnvironmentVariable("CustomisationServiceUrl");
            _log.LogDebug($"Customisation Function Url: {customisationServiceUrl}");

            var customiseSiteFromTemplateInfo = new { TemplateSiteUrl = templateSiteUrl, TargetSiteUrl = targetSiteUrl};

           //var customizeSiteFromTemplateInfo = new { TemplateSite}
            using (var client = new HttpClient())
            using (var request = new HttpRequestMessage(HttpMethod.Post, customisationServiceUrl))
            using (var httpContent = CreateHttpContent(customiseSiteFromTemplateInfo))
            {
                request.Content = httpContent;

                var responseMessage = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                    .ConfigureAwait(false);
                if (responseMessage.IsSuccessStatusCode && responseMessage.StatusCode == System.Net.HttpStatusCode.Accepted)
                {
                    var statusCheckUri = responseMessage.Headers.Location;

                    var count = 50;
                    var queryResponse = await client.GetAsync(statusCheckUri);
                    var queryResult = await queryResponse.Content.ReadAsStringAsync();
                    var result = JObject.Parse(queryResult);
                    var status = result.Value<string>("runtimeStatus");
                    var isComplete = status == "Completed" || status == "Failed";
                    while (count > 0 && !isComplete)
                    {
                        await Task.Delay(10000);
                        count--;
                        queryResponse = await client.GetAsync(statusCheckUri);
                        queryResult = await queryResponse.Content.ReadAsStringAsync();
                        result = JObject.Parse(queryResult);
                        status = result.Value<string>("runtimeStatus");
                        isComplete = status == "Completed" || status == "Failed";// result.Value<string>("runtimeStatus") == "Completed";
                    }
                    if ((isComplete) && (status == "Completed"))
                    {
                        _log.LogInformation("Customisation function completed successfully...");
                        success = result.Value<bool>("output");
                    }
                    
                }
                //else
                //{
                //    teamSiteUrl = await responseMessage.Content.ReadAsStringAsync();
                //}
                //teamSiteUrl = await responseMessage.Content.ReadAsStringAsync();
            }

            return success;
        }

        public static void SerializeJsonIntoStream(object value, Stream stream)
        {
            using (var sw = new StreamWriter(stream, new UTF8Encoding(false), 1024, true))
            using (var jtw = new JsonTextWriter(sw) { Formatting = Formatting.None })
            {
                var js = new JsonSerializer();
                js.Serialize(jtw, value);
                jtw.Flush();
            }
        }

        private static HttpContent CreateHttpContent(object content)
        {
            HttpContent httpContent = null;

            if (content != null)
            {
                var ms = new MemoryStream();
                SerializeJsonIntoStream(content, ms);
                ms.Seek(0, SeekOrigin.Begin);
                httpContent = new StreamContent(ms);
                httpContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            }

            return httpContent;
        }
    }
}
