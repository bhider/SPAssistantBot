using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SPAssistantBot.Services;

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
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string teamsList = req.Query["teamsList"];

            var deletedTeams = await _teamsService.DeleteTeamsAsync(teamsList);

            //string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            //dynamic data = JsonConvert.DeserializeObject(requestBody);
            //name = name ?? data?.name;

            string responseMessage = !string.IsNullOrEmpty(deletedTeams)
                ? $"The following teams were deleted successfully: {deletedTeams}"
                : $"No teams were deleted successfully";

            return new OkObjectResult(responseMessage);
        }
    }
}