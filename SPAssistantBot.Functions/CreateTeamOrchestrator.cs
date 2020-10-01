using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SPAssistantBot.Functions.Models;
using SPAssistantBot.Services;
using System;
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

        private readonly SPCustomisationService _spCustomisationService;

        private ILogger _log;

        //Constructor injection is required, argument/property injection does not work. 
        //Hence the constuctor and hence the class and the method cannot be static.
        public CreateTeamRequestProcessor(TeamsService teamsService, SPCustomisationService spCustomisationService)
        {
            _teamsService = teamsService;
            _spCustomisationService = spCustomisationService;
        }

        [FunctionName("A_CreateTeam")]
        public  async Task<string> CreateTeam([ActivityTrigger]string input, ExecutionContext ec, ILogger log)//, TeamsService teamsService)
        {
            _log = log;

            var newTeamId = string.Empty;
            var createTeamRequest = JsonConvert.DeserializeObject<CreateTeamRequest>(input);

            if (createTeamRequest != null)
            {
                try
                {
                    if (createTeamRequest.UseTemplate)
                    {
                        newTeamId = await _teamsService.CloneTeam("92568ef0-8a32-4029-a847-c0c1add8103d", createTeamRequest.TeamName, createTeamRequest.Description);
                        
                        if (!string.IsNullOrWhiteSpace(newTeamId))
                        {
                            var success = await CustomiseSharePointSiteFromTemplate("92568ef0-8a32-4029-a847-c0c1add8103d", newTeamId);
                        }
                    }
                    else
                    {
                        newTeamId = await _teamsService.CreateTeam(createTeamRequest.TeamName, createTeamRequest.Description, createTeamRequest.OwnersUserEmailListAsString, createTeamRequest.MembersUserEmailListAsString);
                    }

                    string responseMessage = createTeamRequest.TeamName;
                    
                }
                catch (Exception ex)
                {
                    log.LogError(ex.Message);
                    throw;
                }
            }

            return newTeamId;
        }

        private async Task<bool> CustomiseSharePointSiteFromTemplate(string templateTeamId, string newTeamId)
        {
            _log.LogInformation("Applying template to associated SharePoint site for cloned team");
            var success = false;

            var templateSiteUrl = await _teamsService.GetGroupUrlFromTeamId(templateTeamId);
            _log.LogInformation($"Template Site Url: {templateSiteUrl}");

            var teamSiteUrl = await _teamsService.GetGroupUrlFromTeamId(newTeamId);
            _log.LogInformation($"Team Site Url: {teamSiteUrl}");

            if (!string.IsNullOrWhiteSpace(templateSiteUrl) && !string.IsNullOrWhiteSpace(teamSiteUrl))
            {
                success = await _spCustomisationService.CustomiseAsync(templateSiteUrl, teamSiteUrl);
            }

            return success;
        }
        
    }
}
