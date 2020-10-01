using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using SPAssistantBot.Services;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SPAssistantBot.Functions
{
    
    public static class DeleteTeamOrchestrator
    {
        [FunctionName("O_DeleteTeams")]
        public static async Task<string> RunOrchestrator([OrchestrationTrigger]IDurableOrchestrationContext context, ILogger log)
        {
            var deletedTeams = string.Empty;
            try
            {
                var requestBody = context.GetInput<string>();
                deletedTeams = await context.CallActivityAsync<string>("A_DeleteTeams", requestBody);
            }
            catch(Exception ex)
            {
                log.LogError($"{nameof(DeleteTeamOrchestrator)} exception: {ex.Message}");
                throw ex;
            }

            return deletedTeams;

        }
    }

    public class DeleteTeamRequestProcessor
    {
        private readonly TeamsService _teamsService;
        public DeleteTeamRequestProcessor(TeamsService teamsService)
        {
            _teamsService = teamsService;
        }

        [FunctionName("A_DeleteTeams")]
        public async Task<string> DeleteTeams([ActivityTrigger]string input, ExecutionContext ec, ILogger log)
        {
            log.LogInformation($"Deleting teams: {input}");

            try
            {
                var deletedTeams = string.Empty;

                deletedTeams = await _teamsService.DeleteTeamsAsync(input);

                return deletedTeams;
            }
            catch (Exception ex)
            {
                log.LogError($"{nameof(DeleteTeamRequestProcessor)} exception: {ex.Message}");
                throw;
            }
        }
    }
}
