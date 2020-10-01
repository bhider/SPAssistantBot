using Microsoft.Graph;
using Newtonsoft.Json;

namespace SPAssistantBot.Services.Model
{
    public class GroupExtended : Group
    {
        [JsonProperty("owners@odata.bind", NullValueHandling = NullValueHandling.Ignore)]
        public string[] OwnersODataBind { get; set; }

        [JsonProperty("members@odata.bind", NullValueHandling = NullValueHandling.Ignore)]
        public string[] MembersODataBind { get; set; }
    }
}
