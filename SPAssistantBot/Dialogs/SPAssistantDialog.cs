using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Dialogs.Choices;
using Microsoft.Extensions.Configuration;
using Microsoft.Recognizers.Text.Choice;
using Newtonsoft.Json;
using SPAssistantBot.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SPAssistantBot.Dialogs
{
    public class SPAssistantDialog : ComponentDialog
    {
        private readonly SPService spService;

        private readonly IConfiguration configuration;

        public SPAssistantDialog(string dialogId, SPService spService, IConfiguration configuration) : base(dialogId)
        {
            this.spService = spService;
            this.configuration = configuration;

            InitializeWaterfallDialog();
        }

        private void InitializeWaterfallDialog()
        {
            WaterfallStep[] waterfallSteps = new WaterfallStep[]
            {
                InitialStepAsync,
                DescriptionStepAsync,
                SiteTypeStepAsync,
                OwnersListStepAsync,
                MembersListStepAsync,
                FinalStepAsync
            };

            AddDialog(new TextPrompt($"${nameof(SPAssistantDialog)}.siteTitle"));
            AddDialog(new TextPrompt($"{nameof(SPAssistantDialog)}.description"));
            AddDialog(new ChoicePrompt($"{nameof(SPAssistantDialog)}.siteType"));
            AddDialog(new TextPrompt($"{nameof(SPAssistantDialog)}.owners"));
            AddDialog(new TextPrompt($"{nameof(SPAssistantDialog)}.members"));
            AddDialog(new WaterfallDialog($"{nameof(SPAssistantDialog)}.mainflow", waterfallSteps));

            InitialDialogId = $"{nameof(SPAssistantDialog)}.mainflow";
        }

        private async Task<DialogTurnResult> InitialStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            return await stepContext.PromptAsync($"${nameof(SPAssistantDialog)}.siteTitle", new PromptOptions
            {
                Prompt = MessageFactory.Text("What is the title?")
            }, cancellationToken);
        }

        private async Task<DialogTurnResult> DescriptionStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var siteTitle = (string)stepContext.Result;
            stepContext.Values["siteTitle"] = siteTitle;
            return await stepContext.BeginDialogAsync($"{nameof(SPAssistantDialog)}.description", new PromptOptions { Prompt = MessageFactory.Text("Please enter a description for your site") }, cancellationToken);
        }

        public async Task<DialogTurnResult> SiteTypeStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var description = (string)stepContext.Result;
            stepContext.Values["siteDescription"] = description;

            return await stepContext.BeginDialogAsync($"{nameof(SPAssistantDialog)}.siteType", new PromptOptions { Prompt = MessageFactory.Text("What type of site do you want to create?"), Choices = ChoiceFactory.ToChoices(new List<string> { "Team Site", "Communication Site" }) }, cancellationToken);
        }

        private async Task<DialogTurnResult> OwnersListStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var siteType = (FoundChoice)stepContext.Result;
            stepContext.Values["siteType"] = siteType.Value;

            return await stepContext.BeginDialogAsync($"{ nameof(SPAssistantDialog)}.owners", new PromptOptions { Prompt = MessageFactory.Text("Enter emails of site owners") }, cancellationToken);
        }

        private async Task<DialogTurnResult> MembersListStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var owners = (string)stepContext.Result;
            stepContext.Values["siteOwners"] = owners;

            return await stepContext.BeginDialogAsync($"{ nameof(SPAssistantDialog)}.owners", new PromptOptions { Prompt = MessageFactory.Text("Enter emails of site members") }, cancellationToken);
        }

        private async Task<DialogTurnResult> FinalStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var members = (string)stepContext.Result;
            stepContext.Values["siteMembers"] = members;


            var siteTitle = (string)stepContext.Values["siteTitle"];
            var siteType = (string)stepContext.Values["siteType"];
            var description = (string)stepContext.Values["siteDescription"];
            var owners = (string)stepContext.Values["siteOwners"];

            if (!string.IsNullOrWhiteSpace(siteTitle))
            {
                await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Creating site could take a few seconds. Please wait...."), cancellationToken);
                var teamSiteUrl = await CreateSite(siteTitle, description, siteType, owners, members); //spService.CreateSite(siteTitle, description, owners, members);
                await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Site ({teamSiteUrl}) creation complete"), cancellationToken);
            }

            return await stepContext.EndDialogAsync(null, cancellationToken);
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

        private async Task<string> CreateSite(string siteTitle, string description, string siteType, string owners, string members)
        {
            var teamSiteUrl = string.Empty;
            var Url = configuration["CreateSiteUrl"];

            var createSiteRequest = new { SiteTitle = siteTitle, Description = description, SiteType = siteType, OwnersUserEmailListAsString = owners, MembersUserEmailListAsString = members };

            CancellationToken cancellationToken;


            using (var client = new HttpClient())
            using (var request = new HttpRequestMessage(HttpMethod.Post, Url))
            using (var httpContent = CreateHttpContent(createSiteRequest))
            {
                request.Content = httpContent;

                var responseMessage = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
                    .ConfigureAwait(false);
                teamSiteUrl = await responseMessage.Content.ReadAsStringAsync();
            }

            return teamSiteUrl;
        }
    }
}
