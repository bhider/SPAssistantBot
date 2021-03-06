using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using SPAssistantBot.Services;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;

namespace SPAssistantBot.Functions
{
    public  class CreateSite
    {
        private readonly  SPService _spService;

        public CreateSite(SPService spService)
        {
            _spService = spService;
        }
        
        [FunctionName("CreateSite")]
        public  async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, [DurableClient]IDurableOrchestrationClient starter,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();

            var instanceId = await starter.StartNewAsync<string>("O_CreateSPSite", requestBody);

            log.LogInformation($"Started create site orchestration {instanceId}");

            return await starter.WaitForCompletionOrCreateCheckStatusResponseAsync(req, instanceId);
            
        }
    }
}
