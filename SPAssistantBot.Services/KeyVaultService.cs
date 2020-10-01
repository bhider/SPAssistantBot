using Microsoft.Azure.KeyVault;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace SPAssistantBot.Services
{
    public class KeyVaultService
    {
        private readonly string keyVaultSecretIdentifier;
        public KeyVaultService(IConfiguration configuration)
        {
            keyVaultSecretIdentifier = configuration["KeyVaultSecretIdentifier"];
        }

        public async Task<X509Certificate2> GetCertificateAsync(ILogger log)
        {
            try
            {
                var keyVaultClient = GetKeyVaultClient();
                var secret = await keyVaultClient.GetSecretAsync(keyVaultSecretIdentifier);//.GetAwaiter().GetResult();
                var pfxBytes = Convert.FromBase64String(secret.Value);
                return new X509Certificate2(pfxBytes);
            }
            catch (Exception ex)
            {
                log.LogError($"Exception occured retrieving certificate : {ex.Message}");
                throw;
            }
        }

        public string GetSecret(string secretIdentifier, ILogger log)
        {
            try
            {
                var keyVaultClient = GetKeyVaultClient();
                var secret = keyVaultClient.GetSecretAsync(secretIdentifier).GetAwaiter().GetResult();
                return secret.Value;
            }
            catch (Exception ex)
            {
                log.LogError($"Exception occured retrieving secret : {ex.Message}");
                throw;
            }

        }

        private KeyVaultClient GetKeyVaultClient()
        {
            var azureServiceTokenProvider = new AzureServiceTokenProvider();

            return new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(azureServiceTokenProvider.KeyVaultTokenCallback));
        }
    }

}
