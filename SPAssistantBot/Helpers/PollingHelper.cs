using Microsoft.Bot.Builder.Dialogs;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace SPAssistantBot.Helpers
{
    public static class PollingHelper
    {
        

        public static async Task<T> ExecuteLongPollingOperation<T>(HttpClient client, Uri statusQueryUri,  TimeSpan? retryInterval = null, int maxRetries = 50)
        {
            var count = maxRetries;
            
            var result  = await GetOperationResult(client, statusQueryUri);
            
            while (count > 0 && !IsOperationComplete(result))
            {
                await Task.Delay(10000);
                count--;
                result = await GetOperationResult(client, statusQueryUri);
            }
            if ((IsOperationComplete(result)) && (result.status == "Completed"))
            {
                return (T)result.output ;
            }
            else
            {
                return default(T);
            }
        }

        private static bool IsOperationComplete(dynamic result)
        {
            return result.status == "Completed" || result.status == "Failed";
        }

        private static async Task<dynamic> GetOperationResult (HttpClient client, Uri statusQueryUri)
        {
            var queryResponse = await client.GetAsync(statusQueryUri);
            var queryResult = await queryResponse.Content.ReadAsStringAsync();
            var result = JObject.Parse(queryResult);
            var status = result.Value<string>("runtimeStatus");
            var output = result.Value<string>("output");
            return new { status = status, output = output};
        }
    }
}
