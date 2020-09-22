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
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SPAssistantBot.Services
{
    public class TeamsService : BaseO365Service
    {
        private readonly GraphServiceHttpClient _graphServiceHttpClient;

        private string _graphAPIBaseUrl;


        public TeamsService(IConfiguration configuration, KeyVaultService keyVaultService, ILogger<TeamsService> log) : base(configuration, keyVaultService, log)
        {
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

            //var graphBaseUrl = Configuration["GraphAPIBaseUrl"];
            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/groups";

            //await _graphServiceHttpClient.Init();
            var result = await _graphServiceHttpClient.ExecutePostAsync(resourceUrl, JObject.FromObject(group));
            var groupStr = result.ToString();
            var newGroup = JsonConvert.DeserializeObject<Microsoft.Graph.Group>(groupStr);
            return newGroup;
        }

        protected string GraphAPIBaseUrl
        {
            get
            {
                if (string.IsNullOrWhiteSpace(_graphAPIBaseUrl))
                {
                    _graphAPIBaseUrl = Configuration["GraphAPIBaseUrl"];
                }

                return _graphAPIBaseUrl;
            }
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
            string newTeamId = string.Empty;

            var group = await CreateGroup(teamName, description, "Team Site", owners, members);

            if (group != null)
            {
                var groupId = group.Id;
                var team = new Team
                {
                    MemberSettings = new TeamMemberSettings
                    {
                        //AllowCreatePrivateChannels = true, not supported anymore
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
                //await _graphServiceHttpClient.Init();
                //var graphBaseUrl = Configuration["GraphAPIBaseUrl"];
                var resourceUrl = $"{GraphAPIBaseUrl}v1.0/groups/{groupId}/team";
                var result = await _graphServiceHttpClient.ExecutePutAsync(resourceUrl, JObject.FromObject(team));
                newTeamId = groupId;
            }

            return newTeamId;
        }

        public async Task<string> CloneTeam(string templateTeamId, string teamName, string description, string siteType, string owners, string members)
        {
            var newTeamId = string.Empty;

            templateTeamId = "92568ef0-8a32-4029-a847-c0c1add8103d";
            
            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/teams/{templateTeamId}/clone";
            
            var mailNickname = GetMailNickNameFromGroupName(teamName);

            var partsToClone = ClonableTeamParts.Apps | ClonableTeamParts.Tabs | ClonableTeamParts.Settings | ClonableTeamParts.Channels | ClonableTeamParts.Members;

            var visibility = TeamVisibilityType.Public;

            dynamic team = new { displayName = teamName, description = description, mailNickName = mailNickname, partsToClone = partsToClone, visibility = visibility };
            

            try
            {
                var result = await _graphServiceHttpClient.ExecuteLongPollingPostAsync(resourceUrl, JObject.FromObject(team));

                if (result != null)
                {
                    var teamsAysncOperation = JsonConvert.DeserializeObject<TeamsAsyncOperation>(((JObject)result).ToString());
                    newTeamId = teamsAysncOperation.TargetResourceId;
                    
                    if (!string.IsNullOrWhiteSpace(newTeamId))
                    {
                        //Iterate through all the channels and delete the wiki tab
                        await ConfigureChannelTabs(templateTeamId, newTeamId);
                    }
                }
                //var awaiter = graphClient.Teams[templateTeamId]
                //                                        .Clone(visibility, partsToClone, teamName, description, mailNickname, null)
                //                                        .Request()
                //                                        .PostAsync().GetAwaiter();
                //int retryCount = 1;
                //while (!awaiter.IsCompleted && retryCount < 3)
                //{
                //    Thread.Sleep(new TimeSpan(0, 0, 5));
                //    retryCount++;
                //}
            }
            catch (AggregateException aex)
            {
                var msg = aex.Message;
                throw;
            }

            return newTeamId;
        }

        private async Task ConfigureChannelTabs(string templateTeamId, string teamId)
        {
            if (!string.IsNullOrWhiteSpace(templateTeamId) && !string.IsNullOrWhiteSpace(teamId))
            {
                var teamChannels = await GetTeamChannels(teamId);
                //If any channel has a wiki tab delete it; add a one note tab
                await DeleteWikiTabs(teamId, teamChannels);

                var templateTeamChannels = await GetTeamChannels(templateTeamId);
                
                foreach(var templateTeamChannel in templateTeamChannels)
                {
                    //Get corresponding channel in the new Team
                    var newTeamChannel = await GetChannelByName(teamId, templateTeamChannel.DisplayName);

                    if (newTeamChannel != null && !string.IsNullOrWhiteSpace(newTeamChannel.Id))
                    {
                        var templateTabs = await GetChannelTabs(templateTeamId, templateTeamChannel.Id);

                        foreach (var templateTab in templateTabs)
                        {
                            var teamsAppId = templateTab.TeamsAppId;
                            var newTeamTab = await GetTeamsChannelTabByAppIdAndName(teamId, newTeamChannel.Id, teamsAppId, templateTab.DisplayName);

                            if (newTeamTab != null)
                            {
                                if (newTeamTab.TeamsApp.Id == TeamsAppId.OneNote)
                                {
                                    Notebook notebook;
                                    if (await IsDefaultNoteBook(templateTeamId, templateTab.DisplayName))
                                    {
                                        notebook = await GetDefaultNotebook(teamId);
                                    }
                                    else
                                    {
                                        notebook = await CreateNotebook(teamId, WebUtility.UrlEncode(newTeamTab.DisplayName));
                                    }

                                    if (notebook != null)
                                    {
                                        var body = ConfigureOneNoteTab(teamId, notebook);
                                        if (!string.IsNullOrWhiteSpace(body))
                                        {
                                            var result = await UpdateTeamChannelTab(teamId, newTeamChannel.Id, newTeamTab.Id, body);
                                        }
                                    }
                                }
                            }
                        } 
                    }
                    else
                    {
                        Log.LogWarning($"Setting team channel {templateTeamChannel.DisplayName} not complete");
                    }
                }
            }
           
        }

        public async Task<JObject> UpdateTeamChannelTab(string teamId, string channelId, string tabId, string body)
        {
            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/teams/{teamId}/channels/{channelId}/tabs/{tabId}";
            return await _graphServiceHttpClient.ExecutePatchAsync(resourceUrl, JObject.FromObject(body));
        }

        private string ConfigureOneNoteTab(string teamId, Notebook notebook)
        {
            var siteUrl = string.Join("/", notebook.Links.OneNoteWebUrl.Href.Split('/').Take(5));
            return $"{{ 'displayName': '{notebook.DisplayName}', 'configuration': {{ 'contentUrl': 'https://www.onenote.com/teams/TabContent?entityid=%7BentityId%7D&subentityid=%7BsubEntityId%7D&auth_upn=%7Bupn%7D&notebookSource=Pick&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F{teamId}%2Fnotes%2Fnotebooks%2F{notebook.Id}&oneNoteWebUrl={notebook.Links.OneNoteWebUrl.Href}&notebookName={notebook.DisplayName}&siteUrl={siteUrl}&ui={{locale}}&tenantId={{tid}}' }} }}";
        }

        private async Task<Notebook> CreateNotebook(string teamId, string displayName)
        {
            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/groups/{teamId}/onenote/notebooks";
            var body = $"{{ 'displayName': '{displayName}' }}";

            var response = await _graphServiceHttpClient.ExecutePostAsync(resourceUrl, JObject.FromObject(body));

            var notebook = JsonConvert.DeserializeObject<Notebook>(response.ToString());

            return notebook;
        }

        private async Task<Notebook> GetDefaultNotebook(string teamId)
        {
            var resourceUrl = $"{GraphAPIBaseUrl}/v1.0/groups/{teamId}/onenote/notebooks?$orderby=createdDateTime";
            var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);

            if (result != null)
            {
                var valueToken = result.FindTokens("value");

                if (valueToken != null)
                {
                    var notebooksCollection = JsonConvert.DeserializeObject<List<Notebook>>(valueToken.ToString());

                    if (notebooksCollection.Count > 0)
                    {
                        return notebooksCollection[0];
                    }
                }
            }

            return null;
        }

        private async Task<bool> IsDefaultNoteBook(string teamId, string tabDisplayName)
        {
            var team = await GetGroupById(teamId);

            if (team != null)
            {
                return tabDisplayName.StartsWith(team.DisplayName);
            }

            return false;
        }

        public async Task<Microsoft.Graph.Group> GetGroupById(string groupId)
        {
            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/groups/{groupId}";
            var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);

            if(result != null)
            {
                var group = JsonConvert.DeserializeObject<Microsoft.Graph.Group>(result.ToString());
                return group;
            }

            return null;
        }

        private async Task<TeamsTab> GetTeamsChannelTabByAppIdAndName(string teamId, string channelId, string teamsAppId, string displayName)
        {
            var resourceUrl = $"{GraphAPIBaseUrl}beta/teams/{teamId}/channels/{channelId}/tabs?$filter=teamsAppId eq '{teamsAppId}' and displayName eq '{WebUtility.UrlEncode(displayName)}'";

            var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);

            if (result != null)
            {
                var valueToken = result.FindTokens("value");
                if (valueToken != null)
                {
                    var teamsTab = JsonConvert.DeserializeObject<List<TeamsTab>>(valueToken.ToString());

                    if (teamsTab.Count > 0)
                    {
                        return teamsTab[0];
                    }
                }
            }

            return null;
        }

        private async Task<Channel> GetChannelByName(string teamId,string channelName)
        {
            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/teams/{teamId}/channels?$filter=displayName eq '{channelName}'";
            var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);
            var valueToken = result.FindTokens("value").FirstOrDefault();
            if (valueToken != null)
            {
                try
                {
                    var channel = JsonConvert.DeserializeObject<List<Channel>>(valueToken.ToString());
                    if (channel.Count > 0)
                    {
                        return channel[0];
                    }
                }
                catch (Exception ex)
                {
                    var message = ex.Message;
                    
                }
            }

            return null;
        }

        private async Task DeleteWikiTabs(string teamId, List<Channel> teamChannels)
        {
            foreach (var channel in teamChannels)
            {
                var wikiTabsList = await GetChannelWikiTabs(teamId, channel.Id, "wiki");
                foreach (var wikiTab in wikiTabsList)
                {
                    var tabDeleteOpUrl = $"{GraphAPIBaseUrl}beta/teams/{teamId}/channels/{channel.Id}/tabs/{wikiTab.Id}";
                    await _graphServiceHttpClient.ExecuteDeleteAsync(tabDeleteOpUrl);
                }
            }
        }

        public async Task<List<Team>> ListTeams()
        {
            var teamsList = new List<Team>();
            var resourceUrl = $"{GraphAPIBaseUrl}beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')";

            var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);

            if (result != null)
            {
                var valueToken = result.FindTokens("value").FirstOrDefault();
                
                if (valueToken != null)
                {
                    var teams = JsonConvert.DeserializeObject<List<Team>>(valueToken.ToString());
                    teamsList.AddRange(teams);
                }
            }

            return teamsList;
        }

        public async Task<string> DeleteTeamsAsync(string teamsList)
        {
            var deletedTeams = new StringBuilder();
            var teams = await ListTeams();
            var teamsToDeleteArray = teamsList.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach(var teamtoDelete in teamsToDeleteArray)
            {
                var targetTeams = teams.Where(t => t.DisplayName.Trim().Equals(teamtoDelete.Trim(), StringComparison.OrdinalIgnoreCase));
                foreach (var team in targetTeams)
                {
                    if (team != null)
                    {
                        var resourceUrl = $"{GraphAPIBaseUrl}v1.0/groups/{team.Id}";
                        var deleted = await _graphServiceHttpClient.ExecuteDeleteAsync(resourceUrl);
                        if (deleted)
                        {
                            deletedTeams.Append(deletedTeams.Length > 0 ? $", {teamtoDelete}" : $"{teamtoDelete}");
                        }
                    }
                }
            }

            return deletedTeams.ToString();
        }

        public async Task<List<TeamsTabExtended>> GetChannelTabs(string teamId, string channelId)
        {
            var tabsList = new List<TeamsTabExtended>();

            var resourceUrl = $"{GraphAPIBaseUrl}beta/teams/{teamId}/channels/{channelId}/tabs";
            var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);
            var valueToken = result.FindTokens("value").FirstOrDefault();
            if (valueToken != null)
            {
                var tabList = JsonConvert.DeserializeObject<List<TeamsTabExtended>>(valueToken.ToString());
                if (tabList.Count > 0)
                {
                    tabsList.AddRange(tabList);
                }
            }

            return tabsList;
        }

        public async Task<List<TeamsTab>> GetChannelWikiTabs(string groupId, string channelId, string tabName)
        {
            var wikiTabsList = new List<TeamsTab>();
            var resourceUrl = $"{GraphAPIBaseUrl}beta/teams/{groupId}/channels/{channelId}/tabs?$filter=teamsAppId eq 'com.microsoft.teamspace.tab.wiki'";
            var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);
            var valueToken = result.FindTokens("value").FirstOrDefault();
            if (valueToken != null)
            {
                var tabList = JsonConvert.DeserializeObject<List<TeamsTab>>(valueToken.ToString());
                if (tabList.Count > 0)
                {
                    wikiTabsList.AddRange(tabList);
                }
            }

            return wikiTabsList;
        }

        public async Task<List<Channel>> GetTeamChannels(string teamId)
        {
            var teamChannels = new List<Channel>();

            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/teams/{teamId}/channels";

            var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);
            var valueToken = result.FindTokens("value").FirstOrDefault();

            if (valueToken != null)
            {
                List<Channel> channelList = JsonConvert.DeserializeObject<List<Channel>>(valueToken.ToString());
                teamChannels.AddRange(channelList);
            }

            return teamChannels;
        }

        public async Task<Microsoft.Graph.User> GetUserFromEmail(string email)
        {
            var graphClient = new GraphServiceHttpClient(Configuration, Log);
            //await graphClient.Init();
            //var graphBaseUrl = Configuration["GraphAPIBaseUrl"];
            Microsoft.Graph.User user = null;

            var resourceUrl  =  $"{GraphAPIBaseUrl}v1.0/users/{email}";

            var result = await graphClient.ExecuteGetAsync(resourceUrl);
            user = JsonConvert.DeserializeObject<User>(result.ToString());
            
            return user;
        }

        private string GetMailNickNameFromGroupName(string siteTitle)
        {
            var regex = new Regex("[^a-zA-Z0-9]");
            var mailNickName = regex.Replace(siteTitle, "").Trim();
            return mailNickName;
        }

    }
}
