using System.Net.Http;
using System.Threading.Tasks;

namespace SPAssistantBot.Services.Helpers
{
    public static class HttpClientExtensions
    {
        public static async Task<HttpResponseMessage> PatchAsync(this HttpClient httpClient, string requestUri, HttpContent content)
        {
            var method = new HttpMethod("PATCH");

            var request = new HttpRequestMessage(method, requestUri)
            {
                Content = content
            };

            return await httpClient.SendAsync(request);
        }
    }
}
