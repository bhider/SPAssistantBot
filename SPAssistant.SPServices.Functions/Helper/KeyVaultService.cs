using Microsoft.Azure.KeyVault;
using Microsoft.Azure.Services.AppAuthentication;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace SPAssistant.SPServices.Functions.Helper
{
    public static class  KeyVaultService
    {

        private static KeyVaultClient GetKeyVaultClient()
        {
            var azureServiceTokenProvider = new AzureServiceTokenProvider();

            return new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(azureServiceTokenProvider.KeyVaultTokenCallback));
        }

        public static async Task<X509Certificate2> GetCertificateAsync(string certificateIdentifier)
        {
            var client = GetKeyVaultClient();

            var secret = await client.GetSecretAsync(certificateIdentifier);
            var pfxBytes = Convert.FromBase64String(secret.Value);

            return new X509Certificate2(pfxBytes);
        }
    }
}
