using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Threading.Tasks;

namespace SPAssistantBot.Services
{
    public class SPService : BaseO365Service
    {
       
        private readonly SPCustomisationService _spCustomisationService;

        public SPService(IConfiguration configuration, KeyVaultService keyVaultService, SPCustomisationService spCustomisationService,
            ILogger<SPService> log) : base(configuration, keyVaultService, log)
        {
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

    }
}
