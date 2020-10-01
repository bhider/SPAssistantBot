using Microsoft.Bot.Builder.AI.Luis;
using Microsoft.Extensions.Configuration;

namespace SPAssistantBot.Services
{
    public class BotServices
    {
        public BotServices(IConfiguration configuration)
        {
            var luisApplication = new LuisApplication(configuration["LuisAppId"], configuration["LuisAppKey"], $"https://{configuration["LuisAPIHostName"]}.cognitiveservices.azure.com/");

            var luisRecognizerOptions = new LuisRecognizerOptionsV3(luisApplication)
            {
                PredictionOptions = new Microsoft.Bot.Builder.AI.LuisV3.LuisPredictionOptions { IncludeAllIntents = true, IncludeInstanceData = true } 
            };

            Dispatch = new LuisRecognizer(luisRecognizerOptions);
        }

        public LuisRecognizer Dispatch { get; private set; }
    }
}
