using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Dialogs.Choices;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SPAssistantBot.Services;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SPAssistantBot.Dialogs
{
    public class TeamsAssistantDialog : ComponentDialog
    {
        private readonly TeamsService _teamsService;

        private readonly IConfiguration _configuration;
        public TeamsAssistantDialog(string dialogId, TeamsService teamsService, IConfiguration configuration) : base(dialogId)
        {
            _teamsService = teamsService;
            _configuration = configuration;

            InitializeWaterfallDialog();
        }

        private void InitializeWaterfallDialog()
        {
            var waterfallSteps = new WaterfallStep[]
            {
                InitialStepAsync,
                DescriptionStepAsync,
                ConfirmTemplateUseStepAsync,
                TemplateNameStepAsync,
                OwnersListStepAsync,
                MembersListStepAsync,
                FinalStepAsync
            };

            AddDialog(new TextPrompt($"{nameof(TeamsAssistantDialog)}.teamname"));
            AddDialog(new TextPrompt($"{nameof(TeamsAssistantDialog)}.description"));
            AddDialog(new ConfirmPrompt($"{nameof(TeamsAssistantDialog)}.confirm"));
            AddDialog(new ChoicePrompt ($"{nameof(TeamsAssistantDialog)}.templateName"));
            //AddDialog(new ChoicePrompt($"{nameof(TeamsAssistantDialog)}.siteType"));
            AddDialog(new TextPrompt($"{nameof(TeamsAssistantDialog)}.owners"));
            AddDialog(new TextPrompt($"{nameof(TeamsAssistantDialog)}.members"));
            AddDialog(new WaterfallDialog($"{nameof(TeamsAssistantDialog)}.mainflow", waterfallSteps));

            InitialDialogId = $"{nameof(TeamsAssistantDialog)}.mainflow";
        }

        private async Task<DialogTurnResult> InitialStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            return await stepContext.PromptAsync($"{nameof(TeamsAssistantDialog)}.teamname", new PromptOptions
            {
                Prompt = MessageFactory.Text("What is the Team Name?")
            }, cancellationToken);
        }

        private async Task<DialogTurnResult> DescriptionStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var teamName = (string)stepContext.Result;
            stepContext.Values["teamName"] = teamName;
            return await stepContext.BeginDialogAsync($"{nameof(TeamsAssistantDialog)}.description", new PromptOptions { Prompt = MessageFactory.Text("Please enter a description for your team") }, cancellationToken);
        }


        private async Task<DialogTurnResult> ConfirmTemplateUseStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var description = (string)stepContext.Result;
            stepContext.Values["teamDescription"] = description;

            return await stepContext.BeginDialogAsync($"{ nameof(TeamsAssistantDialog)}.confirm", new PromptOptions { Prompt = MessageFactory.Text("Do you want to clone a team?") }, cancellationToken);
        }

        private async Task<DialogTurnResult> TemplateNameStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var useTemplate = (bool)stepContext.Result;
            stepContext.Values["useTemplate"] = useTemplate;

            if (useTemplate)
            {
                return await stepContext.BeginDialogAsync($"{nameof(TeamsAssistantDialog)}.templateName", new PromptOptions { Prompt = MessageFactory.Text("Please choose the team that you want to clone"), Choices = ChoiceFactory.ToChoices(new List<string> { "Project Team Template 241" , "Another template"}) }, cancellationToken);
               //return await stepContext.BeginDialogAsync($"{nameof(SPAssistantDialog)}.siteType", new PromptOptions { Prompt = MessageFactory.Text("What type of site do you want to create?"), Choices = ChoiceFactory.ToChoices(new List<string> { "Team Site", "Communication Site" }) }, cancellationToken);
            }

            return await stepContext.NextAsync(null, cancellationToken);// ", new PromptOptions { Prompt = MessageFactory.Text("Do you want to clone a team?") }, cancellationToken);
        }


        private async Task<DialogTurnResult> OwnersListStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var templateName = (FoundChoice)stepContext.Result;
            stepContext.Values["templateName"] = templateName == null ? string.Empty : templateName.Value;

            return await stepContext.BeginDialogAsync($"{ nameof(TeamsAssistantDialog)}.owners", new PromptOptions { Prompt = MessageFactory.Text("Enter emails of team owners") }, cancellationToken);
        }

        
        private async Task<DialogTurnResult> MembersListStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var owners = (string)stepContext.Result;
            stepContext.Values["teamOwners"] = owners;

            return await stepContext.BeginDialogAsync($"{ nameof(TeamsAssistantDialog)}.members", new PromptOptions { Prompt = MessageFactory.Text("Enter emails of team members") }, cancellationToken);
        }

        private async Task<DialogTurnResult> FinalStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var members = (string)stepContext.Result;
            stepContext.Values["teamMembers"] = members;


            var teamName = (string)stepContext.Values["teamName"];
            var description = (string)stepContext.Values["teamDescription"];
            var useTemplate = (bool)stepContext.Values["useTemplate"];
            var templateName = (string)stepContext.Values["templateName"];
            var owners = (string)stepContext.Values["teamOwners"];

            if (!string.IsNullOrWhiteSpace(teamName))
            {
               
                await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Creating a team could take few seconds. Please wait...."), cancellationToken);
                var newTeam = await CreateTeam(teamName, description, "Team Site", owners, members, useTemplate, templateName); //spService.CreateSite(siteTitle, description, owners, members);
                await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Team ({newTeam}) creation complete"), cancellationToken);
                
            }

            return await stepContext.EndDialogAsync(null, cancellationToken);
        }

        private async Task<string> CreateTeam(string teamName, string description, string teamType, string owners, string members, bool useTemplate, string templateName)
        {
            var newTeam = string.Empty;
            var Url = _configuration["CreateTeamUrl"];

            var createTeamRequest = new { TeamName = teamName, Description = description, TeamType = teamType, OwnersUserEmailListAsString = owners, MembersUserEmailListAsString = members, UseTemplate = useTemplate, TemplateName = templateName };

            CancellationToken cancellationToken;


            using (var client = new HttpClient())
            using (var request = new HttpRequestMessage(HttpMethod.Post, Url))
            using (var httpContent = CreateHttpContent(createTeamRequest))
            {
                request.Content = httpContent;
                
                var responseMessage = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
                    .ConfigureAwait(false);
                if (responseMessage.IsSuccessStatusCode && responseMessage.StatusCode == System.Net.HttpStatusCode.Accepted)
                {
                    var statusCheckUri = responseMessage.Headers.Location;

                    var count = 50;
                    var queryResponse = await client.GetAsync(statusCheckUri);
                    var queryResult = await queryResponse.Content.ReadAsStringAsync();
                    var result = JObject.Parse(queryResult);
                    var status = result.Value<string>("runtimeStatus");
                    var isComplete = status == "Completed" || status == "Failed";
                    while (count > 0 && !isComplete)
                    {
                        await Task.Delay(10000);
                        count--;
                        queryResponse = await client.GetAsync(statusCheckUri);
                        queryResult = await queryResponse.Content.ReadAsStringAsync();
                        result = JObject.Parse(queryResult);
                        status = result.Value<string>("runtimeStatus");
                        isComplete = status == "Completed" || status == "Failed";// result.Value<string>("runtimeStatus") == "Completed";
                    }
                    if ((isComplete) && (status == "Completed"))
                    {
                        newTeam = result.Value<string>("output");
                    }
                    else
                    {
                        newTeam = "Failed";
                    }
                }
                else
                {
                    newTeam = await responseMessage.Content.ReadAsStringAsync();
                }
            }

            return newTeam;
        }

        public static void SerializeJsonIntoStream(object value, Stream stream)
        {
            using (var sw = new StreamWriter(stream, new UTF8Encoding(false), 1024, true))
            using (var jtw = new JsonTextWriter(sw) { Formatting = Formatting.None })
            {
                var js = new JsonSerializer();
                js.Serialize(jtw, value);
                jtw.Flush();
            }
        }

        private static HttpContent CreateHttpContent(object content)
        {
            HttpContent httpContent = null;

            if (content != null)
            {
                var ms = new MemoryStream();
                SerializeJsonIntoStream(content, ms);
                ms.Seek(0, SeekOrigin.Begin);
                httpContent = new StreamContent(ms);
                httpContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            }

            return httpContent;
        }
    }
}
