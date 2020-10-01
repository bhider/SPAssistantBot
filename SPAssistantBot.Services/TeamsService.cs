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

        public async Task<Microsoft.Graph.Group> CreateGroup(string groupName, string description, string owners, string members)
        {
            Log.LogInformation("Creating Group");

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

        public async Task<string> CreateTeam(string teamName, string description, string owners, string members)
        {
            Log.LogInformation("Creating Team....");

            string newTeamId = string.Empty;

            var group = await CreateGroup(teamName, description, owners, members);

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

        public async Task<string> CloneTeam(string templateTeamId, string teamName, string description)
        {
            Log.LogInformation($"Cloning team {templateTeamId}");

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
                        await ConfigureChannelTabs(templateTeamId, newTeamId, teamName);
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

        public async Task<JObject> UpdateTeamChannelTab(string teamId, string channelId, string tabId, dynamic body)
        {
            Log.LogInformation("Updating channel tab");
            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/teams/{teamId}/channels/{channelId}/tabs/{tabId}";
            return await _graphServiceHttpClient.ExecutePatchAsync(resourceUrl, JObject.FromObject(body));
        }

        public async Task<Microsoft.Graph.Group> GetGroupById(string groupId)
        {
            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/groups/{groupId}";
            var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);

            if (result != null)
            {
                var group = JsonConvert.DeserializeObject<Microsoft.Graph.Group>(result.ToString());
                return group;
            }

            return null;
        }

        public async Task<List<Team>> ListTeams()
        {
            Log.LogInformation("Getting a list of teams....");

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

        public async Task<string> GetGroupUrlFromTeamId(string teamId)
        {
            Log.LogInformation("Getting Site url for team...");

            var teamSiteUrl = string.Empty;

            var resourceUrl = $"https://graph.microsoft.com/v1.0/groups/{teamId}/sites/root/weburl";

            var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);

            if (result != null)
            {
                teamSiteUrl = (string)result["value"];
            }

            return teamSiteUrl;
        }

        public async Task<string> DeleteTeamsAsync(string teamsList)
        {
            Log.LogInformation($"Deleting team(s): {teamsList}");

            var deletedTeams = new StringBuilder();
            var teams = await ListTeams();
            var teamsToDeleteArray = teamsList.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var teamtoDelete in teamsToDeleteArray)
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
            Log.LogInformation("Getting channel tabs...");

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

        public async Task<List<TeamsTab>> GetChannelWikiTabs(string groupId, string channelId)
        {
            Log.LogInformation("Getting Wiki tabs...");

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
            Log.LogInformation("Getting Team Channels");

            var teamChannels = new List<Channel>();

            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/teams/{teamId}/channels";

            var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);
            var valueToken = result.FindTokens("value").FirstOrDefault();

            if (valueToken != null)
            {
                List<Channel> channelList = JsonConvert.DeserializeObject<List<Channel>>(valueToken.ToString());
                teamChannels.AddRange(channelList);
            }

            Log.LogInformation($"Found {teamChannels.Count} channels..");
            return teamChannels;
        }

        public async Task<Microsoft.Graph.User> GetUserFromEmail(string email)
        {
            Log.LogInformation($"User: {email}");

            var graphClient = new GraphServiceHttpClient(Configuration, Log);
            //await graphClient.Init();
            //var graphBaseUrl = Configuration["GraphAPIBaseUrl"];
            Microsoft.Graph.User user = null;

            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/users/{email}";

            var result = await graphClient.ExecuteGetAsync(resourceUrl);
            user = JsonConvert.DeserializeObject<User>(result.ToString());

            return user;
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
            Log.LogInformation($"Getting details for uses: {userEmailList}");

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

        private async Task ConfigureChannelTabs(string templateTeamId, string teamId, string teamName)
        {
            Log.LogInformation("Configuring Tabs for all channels...");

            if (!string.IsNullOrWhiteSpace(templateTeamId) && !string.IsNullOrWhiteSpace(teamId))
            {
                var teamChannels = await GetTeamChannels(teamId);
                //If any channel has a wiki tab delete it; add a one note tab
                await DeleteWikiTabs(teamId, teamChannels);

                var templateTeamChannels = await GetTeamChannels(templateTeamId);

                foreach (var templateTeamChannel in templateTeamChannels)
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
                                if (newTeamTab.TeamsAppId == TeamsAppId.OneNote)
                                {
                                    Log.LogInformation("Configuring OneNote Tab...");
                                    Notebook notebook = null;
                                    if (await IsDefaultNoteBook(templateTeamId, templateTab.DisplayName))
                                    {
                                        //    notebook = await GetDefaultNotebook(teamId);
                                        //}
                                        //else
                                        //{
                                        var notebookName = WebUtility.UrlEncode($"{teamName}").Replace("+", " ");
                                        notebook = await CreateNotebook(teamId, $"{notebookName} Notebook");
                                    }

                                    if (notebook != null)
                                    {
                                        var body = ConfigureOneNoteTab(teamId, notebook);
                                        //if (!string.IsNullOrWhiteSpace(body))
                                        //{
                                        var result = await UpdateTeamChannelTab(teamId, newTeamChannel.Id, newTeamTab.Id, body);
                                        //}
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
                Log.LogInformation("Configuring tabs for channels finished");
            }

        }

        private dynamic ConfigureOneNoteTab(string teamId, Notebook notebook)
        {
            var siteUrl = string.Join("/", notebook.Links.OneNoteWebUrl.Href.Split('/').Take(5));
            //return $"{{ 'displayName': '{notebook.DisplayName}', 'configuration': {{ 'contentUrl': 'https://www.onenote.com/teams/TabContent?entityid=%7BentityId%7D&subentityid=%7BsubEntityId%7D&auth_upn=%7Bupn%7D&notebookSource=Pick&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F{teamId}%2Fnotes%2Fnotebooks%2F{notebook.Id}&oneNoteWebUrl={notebook.Links.OneNoteWebUrl.Href}&notebookName={notebook.DisplayName}&siteUrl={siteUrl}&ui={{locale}}&tenantId={{tid}}' }} }}";
            return new
            {
                displayName = notebook.DisplayName,
                configuration = new
                {
                    contentUrl = $"https://www.onenote.com/teams/TabContent?entityid=%7BentityId%7D&subentityid=%7BsubEntityId%7D&auth_upn=%7Bupn%7D&notebookSource=Pick&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F{teamId}%2Fnotes%2Fnotebooks%2F{notebook.Id}&oneNoteWebUrl={notebook.Links.OneNoteWebUrl.Href}&notebookName={notebook.DisplayName}&siteUrl={siteUrl}&ui={{locale}}&tenantId={{tid}}"
                }
            };
        }

        private async Task<Notebook> CreateNotebook(string teamId, string displayName)
        {
            Log.LogInformation($"Creating OneNote notebook {displayName}");

            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/groups/{teamId}/onenote/notebooks";
            try
            {
                var body = new { displayName = displayName };

                var response = await _graphServiceHttpClient.ExecutePostAsync(resourceUrl, JObject.FromObject(body));

                var notebook = JsonConvert.DeserializeObject<Notebook>(response.ToString());
                return notebook;
            }
            catch (Exception ex)
            {
                var message = ex.Message;
                throw;
            }


        }

        private async Task<Notebook> GetDefaultNotebook(string teamId)
        {
            Log.LogInformation("Get default OneNote notebook for team");

            var resourceUrl = $"{GraphAPIBaseUrl}v1.0/groups/{teamId}/onenote/notebooks?$orderby=createdDateTime";
            //var tempHttpClient = new GraphServiceHttpClient2(Configuration, Log);
            var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);// await tempHttpClient.ExecuteGetAsync(resourceUrl);

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

        private async Task<TeamsTabExtended> GetTeamsChannelTabByAppIdAndName(string teamId, string channelId, string teamsAppId, string displayName)
        {
            var resourceUrl = $"{GraphAPIBaseUrl}beta/teams/{teamId}/channels/{channelId}/tabs?$filter=teamsAppId eq '{teamsAppId}' and displayName eq '{WebUtility.UrlEncode(displayName)}'";
            
            var retryCount = 10;

            List<TeamsTabExtended> teamsTab = new List<TeamsTabExtended>();

            var retryInterval = new TimeSpan(0, 0, 5);


            //Note: Retry was found to be necessary - especially when channels and tabs are accessed right after a team has been newly created
            //The dreaded Microsoft Teams related delays - not quite extreme as the delay in site owner showing up but enough to trip up
            //processing logic.
            while (retryCount > 0 && teamsTab.Count == 0)
            {
                Log.LogInformation($"Attempts left {retryCount}");

                var result = await _graphServiceHttpClient.ExecuteGetAsync(resourceUrl);

                if (result != null)
                {
                    var valueToken = result.FindTokens("value").FirstOrDefault();
                    if (valueToken != null)
                    {
                        
                        try
                        {
                            teamsTab = JsonConvert.DeserializeObject<List<TeamsTabExtended>>(valueToken.ToString());
                        }
                        catch (Exception ex)
                        {
                            Log.LogError(ex.Message);
                            //var message = ex.Message;
                        }

                        //if (teamsTab.Count > 0)
                        //{
                        //    return teamsTab[0];
                        //}
                        //else
                        //{
                        //    Log.LogInformation($"Did not find Channel Tab with TeamsAppId {teamsAppId} for tab {displayName}");
                        //}
                    }
                }
                if (teamsTab.Count == 0)
                {
                    retryCount--;
                    Log.LogInformation($"Will attempt to retrieve Tab with TeamsAppId in {retryInterval.TotalSeconds}s. Attempts left: {retryCount}");
                    await Task.Delay(retryInterval);
                }
            }

            if (teamsTab.Count > 0)
            {
                return teamsTab[0];
            }
            else
            {
                Log.LogInformation($"Did not find Channel Tab with TeamsAppId {teamsAppId} for tab {displayName}");
            }

            Log.LogInformation($"Did not find Channel Tab with TeamsAppId {teamsAppId} for tab {displayName}");
            return null;
        }

        private async Task<Channel> GetChannelByName(string teamId,string channelName)
        {
            Log.LogInformation($"Getting channel details (by name) for channel {channelName}");

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
            Log.LogInformation("Deleting Wiki Tabs....");

            foreach (var channel in teamChannels)
            {
                var wikiTabsList = await GetChannelWikiTabs(teamId, channel.Id);
                foreach (var wikiTab in wikiTabsList)
                {
                    var tabDeleteOpUrl = $"{GraphAPIBaseUrl}beta/teams/{teamId}/channels/{channel.Id}/tabs/{wikiTab.Id}";
                    await _graphServiceHttpClient.ExecuteDeleteAsync(tabDeleteOpUrl);
                }
            }
        }

        private string GetMailNickNameFromGroupName(string siteTitle)
        {
            Log.LogInformation("Getting nick name from group name");

            var regex = new Regex("[^a-zA-Z0-9]");
            var mailNickName = regex.Replace(siteTitle, "").Trim();
            return mailNickName;
        }

    }
}
