using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace SPAssistantBot.Services
{
    public class SPService : BaseO365Service
    {
        //private readonly string aadApplicationId;
        //private readonly string aadClientSecret;
        //private readonly string spTenant;
        //private readonly string keyVaultSecretIdentifier;
        //private readonly Microsoft.Extensions.Logging.ILogger log;
        //private readonly KeyVaultService keyVaultService;
        public SPService(IConfiguration configuration, KeyVaultService keyVaultService,
            ILogger<SPService> log) : base(configuration, keyVaultService, log)
        {
            //this.aadApplicationId = configuration["AADClientId"];
            //this.aadClientSecret = configuration["AADClientSecret"];
            //this.spTenant = configuration["Tenant"];
            //this.keyVaultSecretIdentifier = configuration["KeyVaultSecretIdentifier"];
            //this.keyVaultService = keyVaultService;
            ////this.certificate509 = certificate509;
            //this.log = log;
        }

        public string CreateSite(string siteTitle, string description, string owners, string members)
        {
            var teamsiteUrl = string.Empty;
            
            using (var certificate509 = KVService.GetCertificateAsync())
            {
                var repo = new CsomSPRepository(AADApplicationId, AADApplicationSecret, SPTenant, certificate509, Log);
                var groupId  = repo.CreateSite(siteTitle, description,  owners, members);
                teamsiteUrl = repo.GetSiteUrlFromGroupId(groupId);
            }

            return teamsiteUrl;
        }

         

        //private async Task<string> GetAccessToken(string[] scopes)
        //{
        //    using (var certificate509 = KVService.GetCertificateAsync())
        //    {
        //        IConfidentialClientApplication clientApp = ConfidentialClientApplicationBuilder.Create(AADApplicationId).WithCertificate(certificate509).WithTenantId(SPTenant).Build();
        //        var authResult = await clientApp.AcquireTokenForClient(scopes).ExecuteAsync();
        //        var accessToken = authResult.AccessToken;
        //        return accessToken;
        //    }
        //}

        //private ClientContext GetClientContextWithAccessToken(string url, string accessToken)
        //{
        //    Log.LogInformation($"Attempting to get SPO context for site {url}");

        //    var context = new ClientContext(url);

        //    context.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs webRequestEventArgs)
        //    {
        //        webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] = $"Bearer {accessToken}";
        //    };

        //    return context;
        //}
    }
}
