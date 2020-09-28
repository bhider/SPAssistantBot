using Microsoft.Azure.WebJobs.Host;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using SPAssistant.SPServices.Functions.Helper;
using SPAssistant.SPServices.Functions.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace SPAssistant.SPServices.Functions.Services
{
    public interface ISPSiteCustomisationService
    {
        void Customize(string templateUrl, string targetUrl);
    }
    public class SiteCustomisationService : ISPSiteCustomisationService
    {
        private readonly string _applicationId;
        private readonly string _tenant;
        //private readonly string _tenantUrl;
        private readonly X509Certificate2 _certificate509;

        
        private readonly TraceWriter _log;

        private List<CustomDocumentTemplate> _templateContent = new List<CustomDocumentTemplate>();

        public SiteCustomisationService(string tenantId, string applicationId, X509Certificate2 x509Certificate2, TraceWriter log)
        {
            _tenant = tenantId;
            //_tenantUrl = tenantUrl;
            _applicationId = applicationId;
            _certificate509 = x509Certificate2;
            _log = log;
        }

        

        

     
        public void  Customize(string templateUrl, string targetUrl)
        {
             bool isValid = ValidateUrls(templateUrl, targetUrl);
            if (isValid)
            {
                try
                {
                    var template = ExportPnPWebTemplate(templateUrl);
                    ImportPnPWebTemplate(targetUrl, template);
                }
                catch (Exception ex)
                {

                }
            }
            
        }

        private ClientContext CreateContext(string targetUrl)
        {
            return SPContextHelper.GetAuthenticatedAppOnlyContext(_tenant, _applicationId, _certificate509, targetUrl);
        }

        private void ImportPnPWebTemplate(string targetUrl, ProvisioningTemplate template)
        {
            using(var context = CreateContext(targetUrl))
            {
                var web = context.Web;
                context.Load(web);
                context.ExecuteQuery();
                
                SetNoScriptSiteProp(targetUrl);

                var serverRelativeUrl = web.ServerRelativeUrl;

                foreach(var ct in template.ContentTypes)
                {
                    ct.DocumentTemplate = $"{serverRelativeUrl}{ct.DocumentTemplate}";
                }

                var ptai = new ProvisioningTemplateApplyingInformation();

                ptai.ProgressDelegate = delegate (string message, int progress, int total)
                {
                    _log.Info(string.Format("{0} - {1:00}/{2:00} - {3}", targetUrl, progress, total, message));
                };

                web.ApplyProvisioningTemplate(template, ptai);

                

                if (_templateContent.Count > 0)
                {
                    foreach(var documentTemplate in _templateContent)
                    {
                        var connector = new SharePointConnector(context, targetUrl, documentTemplate.ContainerName);

                        using(var ms = new MemoryStream(documentTemplate.Content))
                        {
                            connector.SaveFileStream(documentTemplate.Name, documentTemplate.ContainerName, ms);
                        }
                    }
                }

            }
        }

        private void SetNoScriptSiteProp(string targetUrl)
        {
            var targetUrlComponents = targetUrl.Split(new string[] { ".sharepoint.com" }, StringSplitOptions.RemoveEmptyEntries);

            if (targetUrlComponents.Length > 0)
            {
                var tenantAdminUrl = $"{targetUrlComponents[0]}-admin.sharepoint.com";

                using(var context = CreateContext(tenantAdminUrl))
                {
                    Tenant tenant = new Tenant(context);
                    context.Load(tenant);
                    context.ExecuteQuery();
                    tenant.SetSiteProperties(targetUrl, noScriptSite: false);
                    context.ExecuteQuery();
                }
            }
        }

        private ProvisioningTemplate ExportPnPWebTemplate(string templateUrl)
        {
            using (var context = CreateContext(templateUrl))
            {
                var web = context.Web;
                context.Load(web);
                context.ExecuteQuery();

                var site = context.Site;
                context.Load(site);
                context.ExecuteQuery();

                var ptci = new ProvisioningTemplateCreationInformation(web);
                ptci.HandlersToProcess = Handlers.RegionalSettings | Handlers.ContentTypes | Handlers.Fields | Handlers.Features | Handlers.PageContents;
                ptci.ProgressDelegate = delegate (string message, int progress, int total)
                {
                    _log.Info(string.Format("{0} - {1:00}/{2:00} - {3}", templateUrl, progress, total, message));
                };

                var template = web.GetProvisioningTemplate(ptci);

                //var documentLibraries = GetAllDocumentLibraries(templateUrl);

                var serverRelativeUrl = web.ServerRelativeUrl.ToLower();

                var customDocumenTemplates = new List<string>();

                foreach (var ct in template.ContentTypes)
                {
                    if (!string.IsNullOrWhiteSpace(ct.DocumentTemplate) &&
                        ct.DocumentTemplate.ToLower().Contains(serverRelativeUrl))
                    {
                        var documentPath = ct.DocumentTemplate.ToLower().Replace(serverRelativeUrl, string.Empty);
                        if (!string.IsNullOrEmpty(documentPath))
                        {
                            var indexStartFilename = documentPath.LastIndexOf(@"/");

                            var filename = documentPath.Substring(documentPath.LastIndexOf(@"/") + 1);
                            var containerName = documentPath.Substring(0, indexStartFilename).TrimStart('/');
                            using (var sr =
                                new SharePointConnector(context, templateUrl, containerName).GetFileStream(filename))
                            {
                                if (sr != null)
                                {
                                    using (var ms = new MemoryStream())
                                    {
                                        sr.CopyTo(ms);
                                        _templateContent.Add(new CustomDocumentTemplate(filename, ms.ToArray(),
                                            containerName));
                                    }
                                }
                                else
                                {
                                    throw new ApplicationException($"Document Template {filename} not found.");
                                }
                            }
                        }

                        ct.DocumentTemplate = documentPath;

                    }
                }

                return template;
            }
        }

        private bool ValidateUrls(string templateUrl, string targetUrl)
        {
            var isValid = false;
            try
            {

                isValid = IsValidUrl(templateUrl) && IsValidUrl(targetUrl);
                
            }
            catch(Exception ex)
            {
                
            }
            return isValid;
        }

        private bool IsValidUrl(string targetUrl)
        {
            try
            {
                using (var context = SPContextHelper.GetAuthenticatedAppOnlyContext(_tenant, _applicationId, _certificate509, targetUrl))
                {
                    var web = context.Web;
                    context.Load(web);
                    context.ExecuteQuery();
                    return !string.IsNullOrWhiteSpace(web.Title);
                }
            }
            catch (Exception)
            {

                
            }

            return false;
        }
    }
}
