using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
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

        public string CreateSite(string siteTitle, string description, string owners, string members)
        {
            GraphServiceClient graphClient = GetGraphServiceClient();

            var mailNickName = GetMailNickNameFromSiteTitle(siteTitle);
            var ownerList = GetUserList(owners);
            var memberList = GetUserList(members);

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
            var awaiter = graphClient.Groups.Request().WithScopes(scopes).AddAsync(group).GetAwaiter();

            var result = awaiter.GetResult();

            //try
            //{
            //    var siteawaiter = GetGroupTeamSite(result.Id).GetAwaiter();
            //    var teamSite = siteawaiter.GetResult();
            //    teamSiteUrl = teamSite.WebUrl;
            //}
            //catch(Exception ex)
            //{
            //    log.LogError(ex.Message);
            //}
            return result.Id;
        }

        public string GetSiteUrlFromGroupId(string groupId)
        {
            var siteawaiter = GetGroupTeamSite(groupId).GetAwaiter();
            var teamSite = siteawaiter.GetResult();
            var teamSiteUrl = teamSite.WebUrl;

            return teamSiteUrl;
        }


        public Microsoft.Graph.User GetUserFromEmail(string email)
        {
            var graphCLient = GetGraphServiceClient();
           
            Microsoft.Graph.User user = null;
                        
            try
            {
                var awaiter1 = graphCLient.Users[email].Request().GetAsync().GetAwaiter();
                user = awaiter1.GetResult();
            }
            catch (System.Exception ex)
            {
                var mesg = ex.Message;
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
                var message = ex.Message;
                throw;
            }
        }

        private async Task<Site> GetGroupTeamSiteWithRetry(GraphServiceClient graphClient, string groupId,  int retryInterval = 3, int maxTries = 5)
        {
            Site site = null;

            int count = maxTries;

            while(site == null && count > 0)
            {
                try
                {
                    var httpRequestMsg = graphClient.Groups[groupId].Sites["root"].Request().GetHttpRequestMessage();
                    site = await graphClient.Groups[groupId].Sites["root"].Request().GetAsync();
                }
                catch(Exception ex)
                {
                    if (count > 1)
                    {
                        await Task.Delay(new TimeSpan(0,0,retryInterval));
                    }
                }
                count--;
            }

            return site;
        }


        private string[] GetUserList(string userEmailList)
        {
            var usersList = new List<string>();

            var userEmails = userEmailList.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach(var userEmail in userEmails)
            {
                var user = GetUserFromEmail(userEmail);

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
            var regex = new Regex("[^a-zA-Z0-9]"); 
            var mailNickName = regex.Replace(siteTitle, "").Trim();
            return mailNickName;
        }

        private GraphServiceClient GetGraphServiceClient()
        {
            return GraphClientHelper.GetGraphServiceClient(aadApplicationId, aadClientSecret, spTenant);
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
