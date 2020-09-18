using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SPAssistantBot.Functions.Models;
using SPAssistantBot.Services;

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
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, 
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            //string name = req.Query["name"];
            
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var data = JsonConvert.DeserializeObject<CreateSiteRequest>(requestBody);

            if (data != null)
            {
                try
                {
                    var teamSiteUrl = _spService.CreateSite(data.SiteTitle, data.Description, data.SiteType, data.OwnersUserEmailListAsString, data.MembersUserEmailListAsString);
                    string responseMessage = teamSiteUrl;
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
