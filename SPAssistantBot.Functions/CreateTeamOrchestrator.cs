using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SPAssistantBot.Functions.Models;
using SPAssistantBot.Services;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SPAssistantBot.Functions
{
    public static class CreateTeamOrchestrator
    {
        [FunctionName("O_CreateTeam")]
        public static async Task<string>  RunOrchestrator([OrchestrationTrigger] IDurableOrchestrationContext context, ILogger log)
        {
            var teamSiteUrl = string.Empty;

            try
            {
                var requestBody = context.GetInput<string>();
                teamSiteUrl = await context.CallActivityAsync<string>("A_CreateTeam", requestBody);
            }
            catch(Exception ex)
            {
                var message = ex.Message;
                throw;
            }

            

            return teamSiteUrl;
        }
    }

  
    public  class CreateTeamRequestProcessor
    {
        private readonly TeamsService _teamsService;

        //Constructor injection is required, argument/property injection does not work. 
        //Hence the constuctor and hence the class and the method cannot be static.
        public CreateTeamRequestProcessor(TeamsService teamsService)
        {
            _teamsService = teamsService;
        }

        [FunctionName("A_CreateTeam")]
        public  async Task<string> CreateTeam([ActivityTrigger]string input, ExecutionContext ec, ILogger log)//, TeamsService teamsService)
        {
            var teamSiteUrl = string.Empty;
            var createTeamRequest = JsonConvert.DeserializeObject<CreateTeamRequest>(input);

            if (createTeamRequest != null)
            {
                try
                {
                    if (createTeamRequest.UseTemplate)
                    {
                        teamSiteUrl = await _teamsService.CloneTeam("92568ef0 - 8a32 - 4029 - a847 - c0c1add8103d", createTeamRequest.TeamName, createTeamRequest.Description);
                    }
                    else
                    {
                        teamSiteUrl = await _teamsService.CreateTeam(createTeamRequest.TeamName, createTeamRequest.Description, createTeamRequest.OwnersUserEmailListAsString, createTeamRequest.MembersUserEmailListAsString);
                    }

                    string responseMessage = createTeamRequest.TeamName;
                    //return new OkObjectResult(responseMessage);
                }
                catch (Exception ex)
                {
                    log.LogError(ex.Message);
                    throw;
                }
            }

            return teamSiteUrl;
        }
    }
}
