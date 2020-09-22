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
using System.Runtime.CompilerServices;
using SPAssistantBot.Functions.Models;

namespace SPAssistantBot.Functions
{
    public  class CreateTeam
    {
        private readonly TeamsService _teamsService;
        public CreateTeam(TeamsService teamsService)
        {
            _teamsService = teamsService;
        }

        [FunctionName("CreateTeam")]
        public  async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var data = JsonConvert.DeserializeObject<CreateTeamRequest>(requestBody);

            if (data != null)
            {
                try
                {
                    if (data.UseTemplate)
                    {
                        await _teamsService.CloneTeam("92568ef0 - 8a32 - 4029 - a847 - c0c1add8103d", data.TeamName, data.Description, data.TeamType, data.OwnersUserEmailListAsString, data.MembersUserEmailListAsString);
                    }
                    else
                    {
                        var teamSiteUrl = await _teamsService.CreateTeam(data.TeamName, data.Description, data.TeamType, data.OwnersUserEmailListAsString, data.MembersUserEmailListAsString);
                    }
                     
                    string responseMessage = data.TeamName;
                    return new OkObjectResult(responseMessage);
                }
                catch (Exception ex)
                {
                    log.LogError(ex.Message);
                    throw;
                }
            }
            //name = name ?? data?.name;



            return new OkObjectResult("Site not created");
        }
    }
}
