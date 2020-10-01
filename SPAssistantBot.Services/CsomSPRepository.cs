using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using SPAssistantBot.Services.Helpers;
using SPAssistantBot.Services.Model;
using System;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SPAssistantBot.Services
{
    public class CsomSPRepository
    {
        private readonly string aadApplicationId;
        private readonly string aadClientSecret;
        private readonly string spTenant;
        private readonly X509Certificate2 certificate509;
        private readonly Microsoft.Extensions.Logging.ILogger log;

        public CsomSPRepository(string aadApplicationId, string aadClientSecret, string spTenant, X509Certificate2 certificate509,
            Microsoft.Extensions.Logging.ILogger log)
        {
            this.aadApplicationId = aadApplicationId;
            this.aadClientSecret = aadClientSecret;
            this.spTenant = spTenant;
            this.certificate509 = certificate509;
            this.log = log;
        }

        public async Task<string> CreateSite(string siteTitle, string description, string owners, string members)
        {
            GraphServiceClient graphClient = GetGraphServiceClient();

            var mailNickName = GetMailNickNameFromSiteTitle(siteTitle);
            var ownerList = await GetUserList(owners);
            var memberList = await GetUserList(members);

            var group = new GroupExtended
            {
                Description = string.IsNullOrWhiteSpace(description) ? siteTitle : description,
                DisplayName = siteTitle,
                GroupTypes = new List<string>(){"Unified"},
                MailEnabled = true,
                MailNickname =mailNickName,
                SecurityEnabled = false
            };

            if (ownerList != null && ownerList.Length > 0)
            {
                group.OwnersODataBind = ownerList;
            }

            if (memberList != null && memberList.Length > 0)
            {
                group.MembersODataBind = memberList;
            }

            var scopes = new string[] { "https://graph.microsoft.com/Group.ReadWrite.All" };
            var newGroup = await graphClient.Groups.Request().WithScopes(scopes).AddAsync(group);

            return newGroup.Id;
        }

        public async Task<string> GetSiteUrlFromGroupId(string groupId)
        {
            log.LogInformation("Getting site url for group");

            var teamSiteUrl = string.Empty;

            var teamSite = await GetGroupTeamSite(groupId);

            if (teamSite != null)
            {
                teamSiteUrl = teamSite.WebUrl;
            }

            return teamSiteUrl;
        }

        public async Task<Microsoft.Graph.User> GetUserFromEmail(string email)
        {
            log.LogInformation($"User: {email}");

            var graphCLient = GetGraphServiceClient();
           
            Microsoft.Graph.User user = null;
                        
            try
            {
                user = await graphCLient.Users[email].Request().GetAsync();
                
            }
            catch (System.Exception ex)
            {
                log.LogError($"Get User from email exception: {ex.Message}");
             
            }

            return user;
        }

        private async Task<Site> GetGroupTeamSite(string groupId)
        {
            try
            {
                GraphServiceClient graphClient = GetGraphServiceClient();
                
                var site = await GetGroupTeamSiteWithRetry(graphClient, groupId);
                
                return site;
            }
            catch (Exception ex)
            {
                log.LogError($"Get site from group exception: {ex.Message}");
                throw;
            }
        }

        private async Task<Site> GetGroupTeamSiteWithRetry(GraphServiceClient graphClient, string groupId,  int retryInterval = 3, int maxTries = 5)
        {
            log.LogInformation("Getting group by groupId");

            Site site = null;

            int count = maxTries;
            var ts = new TimeSpan(0, 0, retryInterval);

            //Note rety logic necessary as the Group may not be available immediately after it has been newly setup.
            while(site == null && count > 0)
            {
                try
                {
                    var httpRequestMsg = graphClient.Groups[groupId].Sites["root"].Request().GetHttpRequestMessage();
                    site = await graphClient.Groups[groupId].Sites["root"].Request().GetAsync();
                }
                catch(Exception ex)
                {
                    if (count > 1 && site == null)
                    {
                        log.LogInformation($"Site not found. Trying again in {retryInterval}s. Attempts left {count}");
                        await Task.Delay(ts);
                    }
                }
                count--;
            }

            return site;
        }

        private async Task<string[]> GetUserList(string userEmailList)
        {
            log.LogInformation($"Getting details for users {userEmailList}");

            var usersList = new List<string>();

            var userEmails = userEmailList.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach(var userEmail in userEmails)
            {
                var user = await GetUserFromEmail(userEmail);

                if (user != null)
                {
                    var userStr = $"https://graph.microsoft.com/v1.0/users/{user.Id}";
                    usersList.Add(userStr);
                }
            }

            return usersList.ToArray();
        }

        private string GetMailNickNameFromSiteTitle(string siteTitle)
        {
            log.LogInformation("Getting nickname from site title");

            var regex = new Regex("[^a-zA-Z0-9]"); 
            var mailNickName = regex.Replace(siteTitle, "").Trim();
            return mailNickName;
        }

        private GraphServiceClient GetGraphServiceClient()
        {
            log.LogInformation("Getting authenticated graph service client");
            try
            {
                return GraphClientHelper.GetGraphServiceClient(aadApplicationId, aadClientSecret, spTenant);
            }
            catch (Exception ex)
            {
                log.LogError($"Get graph service client exception: {ex.Message}");
                throw;
            }
            //IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder.Create(aadApplicationId)
            //                                                                                                                                                              .WithTenantId(spTenant)
            //                                                                                                                                                              .WithClientSecret(aadClientSecret)
            //                                                                                                                                                              .Build();
            //ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);

            //GraphServiceClient graphClient = new GraphServiceClient(authProvider);
            //return graphClient;
        }
    }
}
