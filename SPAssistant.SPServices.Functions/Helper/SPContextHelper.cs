using Microsoft.SharePoint.Client;
using System.Security.Cryptography.X509Certificates;

namespace SPAssistant.SPServices.Functions.Helper
{
    public class SPContextHelper
    {
        public static ClientContext GetAuthenticatedAppOnlyContext(string tenant, string applicationId, X509Certificate2 certificate509, string targetUrl )
        {
            ClientContext context = null;

            context = new OfficeDevPnP.Core.AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(targetUrl, applicationId, tenant, certificate509);
            
            return context;
        }
    }
}
