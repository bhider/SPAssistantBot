using Microsoft.Azure.KeyVault;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Extensions.Configuration;
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

        private KeyVaultClient GetKeyVaultClient()
        {
            var azureServiceTokenProvider = new AzureServiceTokenProvider();

            return new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(azureServiceTokenProvider.KeyVaultTokenCallback));
        }

        public async Task<X509Certificate2> GetCertificateAsync()
        {
            var keyVaultClient = GetKeyVaultClient();
            //var certificateBundle = keyVaultClient.GetCertificateAsync(certificateIdentifier).GetAwaiter().GetResult();
            //var secretIdentifier = certificateBundle.SecretIdentifier.Identifier;
            var secret = await keyVaultClient.GetSecretAsync(keyVaultSecretIdentifier);//.GetAwaiter().GetResult();
            var pfxBytes = Convert.FromBase64String(secret.Value);
            return new X509Certificate2(pfxBytes);
        }

        public string GetSecret(string secretIdentifier)
        {
            var keyVaultClient = GetKeyVaultClient();
            var secret = keyVaultClient.GetSecretAsync(secretIdentifier).GetAwaiter().GetResult();
            return secret.Value;

        }
    }

}
