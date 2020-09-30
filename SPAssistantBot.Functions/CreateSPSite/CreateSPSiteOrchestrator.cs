using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using SPAssistantBot.Functions.Models;
using SPAssistantBot.Services;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SPAssistantBot.Functions
{
    public static class CreateSPSiteOrchestrator
    {
        [FunctionName("O_CreateSPSite")]
        public static async Task<string> RunOrchestrator([OrchestrationTrigger] IDurableOrchestrationContext context, ILogger log)
        {
            var teamSiteUrl = string.Empty;


            try
            {
                var requestBody = context.GetInput<string>();
                teamSiteUrl = await context.CallActivityAsync<string>("A_CreateSPSite", requestBody);
            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
                throw;
            }

            return teamSiteUrl;
        }
    }

    public class CreateSPSiteRequestProcessor
    {
        private readonly SPService _spService;
        public CreateSPSiteRequestProcessor(SPService spService)
        {
            _spService = spService;
        }

        [FunctionName("A_CreateSPSite")]
        public async  Task<string> CreateSPSite([ActivityTrigger]string input, ILogger log)
        {
            var data = JsonConvert.DeserializeObject<CreateSiteRequest>(input);

            string teamSiteUrl;

            try
            {
                teamSiteUrl = await _spService.CreateSite(data.SiteTitle, data.Description,  data.TemplateSiteUrl, data.OwnersUserEmailListAsString, data.MembersUserEmailListAsString);
            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
                throw;
            }

            return teamSiteUrl;

        }
    }
}
