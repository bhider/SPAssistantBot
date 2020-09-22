using Microsoft.Graph;
using Newtonsoft.Json;

namespace SPAssistantBot.Services.Model
{
    public class TeamsTabExtended : TeamsTab
    {
        [JsonProperty("teamsAppId", NullValueHandling = NullValueHandling.Ignore)]
        public string TeamsAppId { get; set; }
    }
}
