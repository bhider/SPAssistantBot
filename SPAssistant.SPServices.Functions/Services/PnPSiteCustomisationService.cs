using Microsoft.Extensions.Logging;
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
using System.Security.Cryptography.X509Certificates;
using PnPFolder = OfficeDevPnP.Core.Framework.Provisioning.Model.Folder;

namespace SPAssistant.SPServices.Functions.Services
{
    public interface ISPSiteCustomisationService
    {
        bool Customize(string templateUrl, string targetUrl);
    }

    //Note Uses SharePointPnPCoreOnline classes to export, import and apply templates to
    //SP sites
    public class PnPSiteCustomisationService : ISPSiteCustomisationService
    {
        private readonly string _applicationId;
        private readonly string _tenant;
        private readonly X509Certificate2 _certificate509;

        
        private readonly Microsoft.Extensions.Logging.ILogger _log;

        private List<CustomDocumentTemplate> _templateContent = new List<CustomDocumentTemplate>();

        public PnPSiteCustomisationService(string tenantId, string applicationId, X509Certificate2 x509Certificate2, Microsoft.Extensions.Logging.ILogger log)
        {
            _tenant = tenantId;
            _applicationId = applicationId;
            _certificate509 = x509Certificate2;
            _log = log;
        }

        public bool  Customize(string templateUrl, string targetUrl)
        {
            var success = false;

            bool isValid = ValidateUrls(templateUrl, targetUrl);

            if (isValid)
            {
                try
                {
                    var template = ExportPnPWebTemplate(templateUrl);
                    ImportPnPWebTemplate(targetUrl, template);
                    success = true;
                }
                catch (Exception ex)
                {
                    _log.LogError(ex.Message);
                }
            }
            else
            {
                _log.LogError($"Customise Site Validation failed:  Template Site {templateUrl}; Target Team: {targetUrl}");
            }

            return success;
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
                ptci.HandlersToProcess = Handlers.RegionalSettings | Handlers.ContentTypes | Handlers.Fields | Handlers.Features | Handlers.PageContents | Handlers.Lists;
                ptci.ProgressDelegate = delegate (string message, int progress, int total)
                {
                    _log.LogInformation(string.Format("{0} - {1:00}/{2:00} - {3}", templateUrl, progress, total, message));
                };

                var template = web.GetProvisioningTemplate(ptci);

                var documentLibraries = GetAllDocumentLibraries(context);

                foreach (var docLib in documentLibraries)
                {
                    var listInstance = template.Lists.Find(l => l.Title.ToLower() == docLib.Title.ToLower());

                    if (listInstance != null)
                    {
                        var listFolders = GetDocumentLibrarySubFolders(context, docLib.Title);
                        template.Lists.Remove(listInstance);
                        listInstance.Folders.AddRange(listFolders);
                        template.Lists.Add(listInstance);
                    }
                }

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
                    _log.LogInformation(string.Format("{0} - {1:00}/{2:00} - {3}", targetUrl, progress, total, message));
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

        private ClientContext CreateContext(string targetUrl)
        {
            return SPContextHelper.GetAuthenticatedAppOnlyContext(_tenant, _applicationId, _certificate509, targetUrl, _log);
        }

        private void SetNoScriptSiteProp(string targetUrl)
        {
            _log.LogInformation($"Disabling No Script for site {targetUrl}");

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
                
        private List<PnPFolder> GetDocumentLibrarySubFolders(ClientContext context, string documentLibTitle)
        {
            _log.LogInformation($"Getting sub-folders for library {documentLibTitle}");

            var foldersList = new List<PnPFolder>();

            var docLib = context.Web.Lists.GetByTitle(documentLibTitle);
            context.Load(docLib, l => l.RootFolder);
            context.ExecuteQueryRetry();

            var rootFolder = docLib.RootFolder;
            context.Load(rootFolder.Folders, fldr=> fldr.Include(f=> f.Name, f=> f.ServerRelativeUrl, f=>f.ListItemAllFields));
            context.ExecuteQueryRetry();

            var folders = rootFolder.Folders;

            foreach(var folder in folders)
            {
                if (folder.ListItemAllFields.ServerObjectIsNull != null && folder.ListItemAllFields.ServerObjectIsNull.HasValue)
                {
                    var nestedFolder = GetFolder(context, folder);
                    foldersList.Add(nestedFolder);
                }
            }

            _log.LogInformation($"Feteched  sub-folders for library {documentLibTitle}");

            return foldersList;
        }

        private PnPFolder GetFolder(ClientContext context, Microsoft.SharePoint.Client.Folder folder)
        {
            context.Load(folder, f => f.Name, f => f.Folders);
            context.ExecuteQueryRetry();

            var topFolder = new PnPFolder();
            topFolder.Name = folder.Name;

            foreach (var nestedFolder in folder.Folders)
            {
                var childFolder = GetFolder(context, nestedFolder);
                topFolder.Folders.Add(childFolder);
            }

            return topFolder;
        }

        private List<Microsoft.SharePoint.Client.List> GetAllDocumentLibraries(ClientContext context)
        {
            _log.LogInformation("Getting all document libraries");

            var web = context.Web;
            context.Load(web.Lists);
            context.ExecuteQuery();
            var documentLibraries = new List<Microsoft.SharePoint.Client.List>();
            
            foreach(var list in web.Lists)
            {
                if (list.BaseType.ToString().Equals("DocumentLibrary", StringComparison.OrdinalIgnoreCase) && !list.Hidden)
                {
                    documentLibraries.Add(list);
                }
            }

            return documentLibraries;
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
            _log.LogInformation($"Validating url {targetUrl}");

            try
            {
                using (var context = SPContextHelper.GetAuthenticatedAppOnlyContext(_tenant, _applicationId, _certificate509, targetUrl, _log))
                {
                    var web = context.Web;
                    context.Load(web);
                    context.ExecuteQuery();
                    return !string.IsNullOrWhiteSpace(web.Title);
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"Exception occured when validating url {targetUrl} : {ex.Message}");
            }

            return false;
        }
    }
}
