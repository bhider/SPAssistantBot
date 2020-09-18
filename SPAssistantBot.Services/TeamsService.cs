using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SPAssistantBot.Services.Helpers;
using SPAssistantBot.Services.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace SPAssistantBot.Services
{
    public class TeamsService : BaseO365Service
    {
        private readonly GraphServiceHttpClient _graphServiceHttpClient;
        //private readonly string aadApplicationId;
        //private readonly string aadClientSecret;
        //private readonly string spTenant;
        //private readonly string keyVaultSecretIdentifier;
        //private readonly Microsoft.Extensions.Logging.ILogger log;
        //private readonly KeyVaultService keyVaultService;
        public TeamsService(IConfiguration configuration, KeyVaultService keyVaultService, ILogger<TeamsService> log) : base(configuration, keyVaultService, log)
        {
            //AADApplicationId = configuration["AADClientId"];
            //AADApplicationSecret = configuration["AADClientSecret"];
            //SPTenant = configuration["Tenant"];
            //KVServiceCertIdentifier = configuration["KeyVaultSecretIdentifier"];
            //KVService = keyVaultService;
            //log = log;
            _graphServiceHttpClient = new GraphServiceHttpClient(configuration, log);
        }

        public async Task<Microsoft.Graph.Group> CreateGroup(string groupName, string description, string siteType, string owners, string members)
        {
            var mailNickName = GetMailNickNameFromGroupName(groupName);
            var ownerList = await GetUserList(owners);
            var memberList = await GetUserList(members);

            var group = new GroupExtended
            {
                Description = string.IsNullOrWhiteSpace(description) ? groupName : description,
                DisplayName = groupName,
                GroupTypes = new List<string>() { "Unified" },
                MailEnabled = true,
                MailNickname = mailNickName,
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

            var graphBaseUrl = Configuration["GraphAPIBaseUrl"];
            var resourceUrl = $"{graphBaseUrl}v1.0/groups";

            await _graphServiceHttpClient.Init();
            var result = await _graphServiceHttpClient.ExecutePostAsync(resourceUrl, JObject.FromObject(group));
            var groupStr = result.ToString();
            var newGroup = JsonConvert.DeserializeObject<Microsoft.Graph.Group>(groupStr);
            return newGroup;
        }

        private async Task<string[]> GetUserList(string userEmailList)
        {
            var usersList = new List<string>();

            var userEmails = userEmailList.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var userEmail in userEmails)
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

        public async Task<string> CreateTeam(string teamName, string description, string siteType, string owners, string members)
        {
            //using (var certificate509 = KVService.GetCertificateAsync())
            //{
            //var repo = new CsomSPRepository(AADApplicationId, AADApplicationSecret, SPTenant, certificate509, Log);
            //var groupId = repo.CreateSite(teamName, description, "Team Site", owners, members);
            //var graphClient = GraphClientHelper.GetGraphServiceClient(AADApplicationId, AADApplicationSecret, SPTenant);
            //var team = new Team
            //{
            //    MemberSettings = new TeamMemberSettings
            //    {
            //        //AllowCreatePrivateChannels = true,
            //        AllowCreateUpdateChannels = true,
            //        ODataType = null
            //    },
            //    MessagingSettings = new TeamMessagingSettings
            //    {
            //        AllowUserEditMessages = true,
            //        AllowUserDeleteMessages = true,
            //        ODataType = null
            //    },
            //    FunSettings = new TeamFunSettings
            //    {
            //        AllowGiphy = true,
            //        GiphyContentRating = GiphyRatingType.Strict,
            //        ODataType = null
            //    },
            //    ODataType = null

            //};

            //try
            //{
            //    var ownerList = owners.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            //    await GetUserFromEmail(ownerList[0]);
            //}
            //catch(Exception ex)
            //{
            //    var message = ex.Message;
            //}

            //var graphHttpClient = new GraphServiceHttpClient(Configuration, Log);
            var group = await CreateGroup(teamName, description, "Team Site", owners, members);

            if (group != null)
            {
                var groupId = group.Id;
                var team = new Team
                {
                    MemberSettings = new TeamMemberSettings
                    {
                        //AllowCreatePrivateChannels = true,
                        AllowCreateUpdateChannels = true,
                        ODataType = null
                    },
                    MessagingSettings = new TeamMessagingSettings
                    {
                        AllowUserEditMessages = true,
                        AllowUserDeleteMessages = true,
                        ODataType = null
                    },
                    FunSettings = new TeamFunSettings
                    {
                        AllowGiphy = true,
                        GiphyContentRating = GiphyRatingType.Strict,
                        ODataType = null
                    },
                    ODataType = null

                };
                await _graphServiceHttpClient.Init();
                var graphBaseUrl = Configuration["GraphAPIBaseUrl"];
                var resourceUrl = $"{graphBaseUrl}v1.0/groups/{groupId}/team";
                var result = await _graphServiceHttpClient.ExecutePutAsync(resourceUrl, JObject.FromObject(team));
            }

                //var awaiter =  graphClient.Groups[groupId].Team.Request().PutAsync(team).GetAwaiter();
                //var result = awaiter.GetResult();

                return string.Empty;//result.Id;
            //}
        }

        public void CloneTeam(string templateTeamId, string teamName, string description, string siteType, string owners, string members)
        {
            templateTeamId = "92568ef0-8a32-4029-a847-c0c1add8103d";

            var graphClient = GraphClientHelper.GetGraphServiceClient(AADApplicationId, AADApplicationSecret, SPTenant);

            //var displayName = "Library Assist";

            //var description = "Self help community for library";

            var mailNickname = GetMailNickNameFromGroupName(teamName);

            var partsToClone = ClonableTeamParts.Apps | ClonableTeamParts.Tabs | ClonableTeamParts.Settings | ClonableTeamParts.Channels | ClonableTeamParts.Members;

            var visibility = TeamVisibilityType.Public;

            try
            {
                var awaiter = graphClient.Teams[templateTeamId]
                        .Clone(visibility, partsToClone, teamName, description, mailNickname, null)
                        .Request()
                        .PostAsync().GetAwaiter();
                int retryCount = 1;
                while (!awaiter.IsCompleted && retryCount < 3)
                {
                    Thread.Sleep(new TimeSpan(0, 0, 5));
                    retryCount++;
                }
            }
            catch (AggregateException aex)
            {
                var msg = aex.Message;
                throw;
            }
            
        }

        public async Task<Microsoft.Graph.User> GetUserFromEmail(string email)
        {
            var graphClient = new GraphServiceHttpClient(Configuration, Log);
            await graphClient.Init();
            var graphBaseUrl = Configuration["GraphAPIBaseUrl"];
            Microsoft.Graph.User user = null;

            var resourceUrl  =  $"{graphBaseUrl}v1.0/users/{email}";

            var result = await graphClient.ExecuteGet(resourceUrl);
            user = JsonConvert.DeserializeObject<User>(result.ToString());
            //try
            //{
            //    var awaiter1 = graphCLient.Users["63e8948c-35ef-4dc9-b305-1538edabd841"].Request().GetAsync().GetAwaiter();
            //    user = awaiter1.GetResult();
            //}
            //catch (System.Exception ex)
            //{

            //    var mesg = ex.Message;
            //}

            try
            {
                
                //var awaiter1 = graphCLient.Users[email].Request().GetAsync().GetAwaiter();
                //user = awaiter1.GetResult();
            }
            catch (System.Exception ex)
            {

                var mesg = ex.Message;
            }

            return user;
        }

        private string GetMailNickNameFromGroupName(string siteTitle)
        {
            var regex = new Regex("[^a-zA-Z0-9]");
            var mailNickName = regex.Replace(siteTitle, "").Trim();
            return mailNickName;
        }

        //private string GetMailNickNameFromName(string siteTitle)
        //{
        //    var regex = new Regex("[^a-zA-Z0-9]");
        //    var mailNickName = regex.Replace(siteTitle, "").Trim();
        //    return mailNickName;
        //}
    }
}
