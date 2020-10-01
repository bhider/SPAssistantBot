using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using SPAssistant.SPServices.Functions.Helper;

namespace SPAssistant.SPServices.Functions
{
    public  class CustomiseSPSite
    {
        static CustomiseSPSite()
        {
            ApplicationHelper.Startup();
        }

        [FunctionName("CustomiseSPSite")]
        public  static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, [DurableClient] IDurableOrchestrationClient starter,  TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            string requestBody = await new StreamReader(await req.Content.ReadAsStreamAsync()).ReadToEndAsync();
            var instanceId = await starter.StartNewAsync<string>("O_CustomiseSPSite", requestBody);
            log.Info($"Started orchestration with Id {instanceId}");

            return starter.CreateCheckStatusResponse(req, instanceId);
            
        }
    }
}
