﻿using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SPAssistant.SPServices.Functions.Helper;
using SPAssistant.SPServices.Functions.Models;
using SPAssistant.SPServices.Functions.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPAssistant.SPServices.Functions
{
    public static class CustomiseSPSiteOrchestrator
    {
        [FunctionName("O_CustomiseSPSite")]
        public static async Task<bool> RunOrchestrator([OrchestrationTrigger]IDurableOrchestrationContext context, ILogger log)
        {
            var success = false;

            try
            {
                var requestBody = context.GetInput<string>();
                success = await context.CallActivityAsync<bool>("A_CustomiseSPSite", requestBody);
            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
                throw;
            }

            return success;
        }
    }

    public static class CustomiseSPSiteRequestProcessor
    {
        [FunctionName("A_CustomiseSPSite")]
        public static async Task<bool> CustomiseSPSite([ActivityTrigger]string input, ExecutionContext ec, ILogger log)
        {
            var success = false;

            var applicationId = Environment.GetEnvironmentVariable("AADClientId");
            var tenantId = Environment.GetEnvironmentVariable("TenantId");
             var keyVaultCertificateIdentifier = Environment.GetEnvironmentVariable("KeyVaultCertificateIdentifier");

            var createRequest = JsonConvert.DeserializeObject<CustomizeSiteFromTemplateInfo>(input);

            try
            {
                using (var certificate509 = await KeyVaultService.GetCertificateAsync(keyVaultCertificateIdentifier))
                {
                    var service = new PnPSiteCustomisationService(tenantId,  applicationId, certificate509, log);
                    success = service.Customize(createRequest.TemplateSiteUrl, createRequest.TargetSiteUrl);
                }
            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
                throw;
            }
            return success;
        }
    }
}
