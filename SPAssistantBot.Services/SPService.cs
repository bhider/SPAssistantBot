using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using System;
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
        private readonly SPCustomisationService _spCustomisationService;
        public SPService(IConfiguration configuration, KeyVaultService keyVaultService, SPCustomisationService spCustomisationService,
            ILogger<SPService> log) : base(configuration, keyVaultService, log)
        {
            //this.aadApplicationId = configuration["AADClientId"];
            //this.aadClientSecret = configuration["AADClientSecret"];
            //this.spTenant = configuration["Tenant"];
            //this.keyVaultSecretIdentifier = configuration["KeyVaultSecretIdentifier"];
            //this.keyVaultService = keyVaultService;
            ////this.certificate509 = certificate509;
            //this.log = log;
            _spCustomisationService = spCustomisationService;
        }

        public async Task<string> CreateSite(string siteTitle, string description, string templateSiteUrl, string owners, string members)
        {
            var teamSiteUrl = string.Empty;

            try
            {
                using (var certificate509 = await KVService.GetCertificateAsync(Log))
                {
                    var repo = new CsomSPRepository(AADApplicationId, AADApplicationSecret, SPTenant, certificate509, Log);
                    var groupId = await repo.CreateSite(siteTitle, description, owners, members);

                    teamSiteUrl = await repo.GetSiteUrlFromGroupId(groupId);

                    if (!string.IsNullOrWhiteSpace(templateSiteUrl))
                    {
                        if (!string.IsNullOrWhiteSpace(teamSiteUrl))
                        {
                            var success = await _spCustomisationService.CustomiseAsync(templateSiteUrl, teamSiteUrl);
                        }
                        else
                        {
                            Log.LogWarning($"Target team Url could not be determined. Target team '{siteTitle}' will not be customised");
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                Log.LogError($"Create Site exception: {ex.Message}");
                throw;
            }

            return teamSiteUrl;
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
