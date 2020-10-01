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
    public  class DeleteTeam
    {
        private readonly TeamsService _teamsService;

        public DeleteTeam(TeamsService teamsService)
        {
            _teamsService = teamsService;
        }


        [FunctionName("DeleteTeam")]
        public  async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, 
            [DurableClient]IDurableOrchestrationClient starter,
            ILogger log, ExecutionContext ex)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string teamsList = req.Query["teamsList"];

            var instanceId = await starter.StartNewAsync<string>("O_DeleteTeams", teamsList);

            return starter.CreateCheckStatusResponse(req, instanceId);//;WaitForCompletionOrCreateCheckStatusResponseAsync(req, instanceId);
            
        }
    }
}
