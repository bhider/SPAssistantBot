using Microsoft.Azure.KeyVault;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Extensions.Logging;
using System;
using System.Security.Cryptography.X509Certificates;
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

        //Note: While debugging locally make sure that the account that is used to sign to Visual Studio has been
        //granted access to the KeyVault.
        //When deployed, enable Managed Identity for the Function App and ensure that the Managed identity has been granted appropriate
        //access to the KeyVault.
        public static async Task<X509Certificate2> GetCertificateAsync(string certificateIdentifier, ILogger log)
        {
            log.LogInformation("Retrieving certificate...");
            try
            {
                var client = GetKeyVaultClient();

                var secret = await client.GetSecretAsync(certificateIdentifier);
                var pfxBytes = Convert.FromBase64String(secret.Value);

                return new X509Certificate2(pfxBytes);
            }
            catch (Exception ex)
            {
                log.LogError($"Retrieve certificate exception: {ex.Message}");
                throw;
            }
        }
    }
}
