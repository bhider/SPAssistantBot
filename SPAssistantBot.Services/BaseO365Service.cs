using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using System.Threading.Tasks;

namespace SPAssistantBot.Services
{
    public abstract class BaseO365Service
    {
        //private readonly string aadApplicationId;
        //private readonly string aadClientSecret;
        //private readonly string spTenant;
        //private readonly string keyVaultSecretIdentifier;
        //private readonly Microsoft.Extensions.Logging.ILogger log;
        //private readonly KeyVaultService keyVaultService;
        public BaseO365Service(IConfiguration configuration, KeyVaultService keyVaultService, ILogger log)
        {
            AADApplicationId = configuration["AADClientId"];
            AADApplicationSecret = configuration["AADClientSecret"];
            SPTenant = configuration["Tenant"];
            KVServiceCertIdentifier = configuration["KeyVaultSecretIdentifier"];
            KVService = keyVaultService;
            Configuration = configuration;
            Log = log;
        }

        protected async Task<string> GetAccessToken(string[] scopes)
        {
            using (var certificate509 = KVService.GetCertificateAsync())
            {
                IConfidentialClientApplication clientApp = ConfidentialClientApplicationBuilder.Create(AADApplicationId).WithCertificate(certificate509).WithTenantId(SPTenant).Build();
                var authResult = await clientApp.AcquireTokenForClient(scopes).ExecuteAsync();
                var accessToken = authResult.AccessToken;
                return accessToken;
            }
        }

        protected ClientContext GetClientContextWithAccessToken(string url, string accessToken)
        {
            Log.LogInformation($"Attempting to get SPO context for site {url}");

            var context = new ClientContext(url);

            context.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs webRequestEventArgs)
            {
                webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] = $"Bearer {accessToken}";
            };

            return context;
        }

        protected IConfiguration Configuration { get; private set; }

        protected string AADApplicationId{get; private set;}

        protected string AADApplicationSecret { get; private set; }

        protected string SPTenant { get; private set; }

        protected ILogger Log { get; private set; }

        protected string KVServiceCertIdentifier { get; private set; }

        protected KeyVaultService KVService{ get; private set; }
    }
}
