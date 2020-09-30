using AngleSharp.Dom.Css;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Sites;
using SPAssistant.SPServices.Functions.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.PeerToPeer.Collaboration;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SPAssistant.SPServices.Functions.Services
{
    //class SPService
    //{
    //    private readonly string _applicationId;
    //    private readonly string _tenant;
    //    private readonly string _tenantUrl;
    //    private readonly X509Certificate2 _certificate509;
    //    private readonly ILogger _log;

    //    public SPService(string tenantId, string tenantUrl, string applicationId, X509Certificate2 x509Certificate2, ILogger log)
    //    {
    //        _tenant = tenantId;
    //        _tenantUrl = tenantUrl;
    //        _applicationId = applicationId;
    //        _certificate509 = x509Certificate2;
    //        _log = log;
    //    }

    //    private ClientContext CreateContext(string url)
    //    {
    //        ClientContext context = null;

    //        try
    //        {
    //            context = new OfficeDevPnP.Core.AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(url, _applicationId, _tenant, _certificate509);
    //        }
    //        catch(Exception ex)
    //        {
    //            _log.LogError(ex.Message);
    //            throw;
    //        }

    //        return context;
    //    }

    //    public void CustomiseSiteFromTemplate(CustomizeSiteFromTemplateInfo customizeSiteFromTemplateInfo)
    //    {
    //        //var teamSiteUrl = await CreateSPSite(createSiteRequest);

    //        if (!string.IsNullOrWhiteSpace(customizeSiteFromTemplateInfo.TemplateSiteUrl) && !string.IsNullOrWhiteSpace(customizeSiteFromTemplateInfo.TargetSiteUrl))
    //        {
    //            var customizationService = new SiteCustomisationService(_tenant, _applicationId, _certificate509, _log);
    //            customizationService.Customize(customizeSiteFromTemplateInfo.TemplateSiteUrl, customizeSiteFromTemplateInfo.TargetSiteUrl);
    //        }

    //        //return teamSiteUrl;
    //    }



    //    //private async Task<string> CreateSPSite(CustomizeSiteFromTemplateInfo createSiteRequest)
    //    //{
    //    //    var teamSiteUrl = string.Empty;

    //    //    var Url = Environment.GetEnvironmentVariable("CreateSiteUrl");

    //    //    using (var client = new HttpClient())
    //    //    using (var request = new HttpRequestMessage(HttpMethod.Post, Url))
    //    //    using (var httpContent = CreateHttpContent(createSiteRequest))
    //    //    {
    //    //        request.Content = httpContent;

    //    //        var responseMessage = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
    //    //            .ConfigureAwait(false);
    //    //        teamSiteUrl = await responseMessage.Content.ReadAsStringAsync();
    //    //    }

    //    //    return teamSiteUrl;
    //    //}

    //    public static void SerializeJsonIntoStream(object value, Stream stream)
    //    {
    //        using (var sw = new StreamWriter(stream, new UTF8Encoding(false), 1024, true))
    //        using (var jtw = new JsonTextWriter(sw) { Formatting = Formatting.None })
    //        {
    //            var js = new JsonSerializer();
    //            js.Serialize(jtw, value);
    //            jtw.Flush();
    //        }
    //    }

    //    private static HttpContent CreateHttpContent(object content)
    //    {
    //        HttpContent httpContent = null;

    //        if (content != null)
    //        {
    //            var ms = new MemoryStream();
    //            SerializeJsonIntoStream(content, ms);
    //            ms.Seek(0, SeekOrigin.Begin);
    //            httpContent = new StreamContent(ms);
    //            httpContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
    //        }

    //        return httpContent;
    //    }

    //    //public async Task<string> CreateSPSite(CreateSiteRequest createSiteRequest)
    //    //{
    //    //    string siteUrl = null;

    //    //    using(var context = CreateContext(_tenantUrl))
    //    //    {
    //    //        var teamSiteCollectionInfo = new TeamSiteCollectionCreationInformation
    //    //        {
    //    //            Alias = GetSiteAliasFromTitle(createSiteRequest.SiteTitle),
    //    //            DisplayName = createSiteRequest.SiteTitle,
    //    //            Description = createSiteRequest.Description,
    //    //            IsPublic = true
    //    //        };

    //    //        var teamContext = await context.CreateSiteAsync(teamSiteCollectionInfo);

    //    //        teamContext.Load(teamContext.Web, w => w.Url);
    //    //        teamContext.ExecuteQueryRetry();
    //    //        siteUrl = teamContext.Web.Url;
    //    //    }

    //    //    return siteUrl;
    //    //}

    //    private string GetSiteAliasFromTitle(string siteTitle)
    //    {
    //        var regex = new Regex("[^a-zA-Z0-9]");
    //        var siteAlias = regex.Replace(siteTitle, "").Trim();
    //        return siteAlias;
    //    }
    //}
}
