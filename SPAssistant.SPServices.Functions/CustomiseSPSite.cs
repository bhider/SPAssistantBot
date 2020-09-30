using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using SPAssistant.SPServices.Functions.Helper;

namespace SPAssistant.SPServices.Functions
{
    public  class CustomiseSPSite
    {
        static CustomiseSPSite()
        {
            ApplicationHelper.Startup();
        }
        [FunctionName("CustomiseSPSite")]
        public  static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, [DurableClient] IDurableOrchestrationClient starter,  TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            // parse query parameter

            //var applicationId = Environment.GetEnvironmentVariable("AADClientId");
            //var applicationSecret = Environment.GetEnvironmentVariable("AADClientSecret");
            //var tenantId = Environment.GetEnvironmentVariable("TenantId");
            //var tenant = Environment.GetEnvironmentVariable("Tenant");
            //var tenantUrl = Environment.GetEnvironmentVariable("TenantUrl");
            //var keyVaultCertificateIdentifier = Environment.GetEnvironmentVariable("KeyVaultCertificateIdentifier");

            // Get request body
            //dynamic data = await req.Content.ReadAsAsync<object>();
            string requestBody = await new StreamReader(await req.Content.ReadAsStreamAsync()).ReadToEndAsync();
            var instanceId = await starter.StartNewAsync<string>("O_CustomiseSPSite", requestBody);
            log.Info($"Started orchestration with Id {instanceId}");

            return starter.CreateCheckStatusResponse(req, instanceId);
            //var createRequest = JsonConvert.DeserializeObject<CreateSiteRequest>(requestBody);

            //var siteUrl = string.Empty;

            //try
            //{
            //    using (var certificate509 = await KeyVaultService.GetCertificateAsync(keyVaultCertificateIdentifier))
            //    {
            //        var service = new SPService(tenantId, tenantUrl, applicationId, certificate509, log);
            //        siteUrl = await service.ProcessCreateSiteRequest(createRequest);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    var message = ex.Message;
            //    throw;
            //}
            //return req.CreateResponse<string>(HttpStatusCode.OK, siteUrl);
            //return name == null
            //    ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
            //    : req.CreateResponse(HttpStatusCode.OK, "Hello " + name);
        }
    }
}
