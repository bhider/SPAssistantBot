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
using System.Security;
using SPAssistantBot.Services.Helpers;

namespace SPAssistantBot.Functions
{
    public class CreateSiteFromTemplate
    {
        private readonly SPService _spService;

        public CreateSiteFromTemplate(SPService sPService)
        {
            _spService = sPService;
        }

        [FunctionName("CreateSiteFromTemplate")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string name = req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            name = name ?? data?.name;

            var uri = new Uri("https://sptestbot.sharepoint.com/sites/TestBotSite6");
            var user = "rajB@sptestbot.onmicrosoft.com";
            var password =  GetSecureString("plassey#1");

            using(var authenticationManager = new AuthenticationManager())
            using(var context = authenticationManager.GetContext(uri, user, password))
            {
                context.Load(context.Web, p => p.Title);
                await context.ExecuteQueryAsync();
                var title = context.Web.Title;
            }

            string responseMessage = string.IsNullOrEmpty(name)
                ? "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."
                : $"Hello, {name}. This HTTP triggered function executed successfully.";

            return new OkObjectResult(responseMessage);
        }

        private static SecureString GetSecureString(string password)
        {
            var secureString = new SecureString();
            foreach(var ch in password)
            {
                secureString.AppendChar(ch);
            }

            return secureString;
        }
    }
}
