using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Extensions.Configuration;
using SPAssistantBot.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace SPAssistantBot.Dialogs
{
    public class DeleteTeamDialog : ComponentDialog
    {
        private readonly TeamsService _teamsService;

        private readonly IConfiguration _configuration;
        public DeleteTeamDialog(string dialogId, TeamsService teamsService, IConfiguration configuration) : base(dialogId)
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
                FinalStepAsync
            };

            AddDialog(new TextPrompt($"{nameof(DeleteTeamDialog)}.teamsToDelete"));
            AddDialog(new WaterfallDialog($"{nameof(DeleteTeamDialog)}.mainflow", waterfallSteps) );

            InitialDialogId = $"{nameof(DeleteTeamDialog)}.mainflow";
        }

        private async Task<DialogTurnResult> InitialStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            return await stepContext.BeginDialogAsync($"{nameof(DeleteTeamDialog)}.teamsToDelete", new PromptOptions { Prompt = MessageFactory.Text("Please enter the name(s) of Teams you want to delete. If mutiple teams are to be deleted, use commas to separate team names") }, cancellationToken);
        }

        private async Task<string> DeleteTeams(string teamsList)
        {
            var response = string.Empty;

            var deleteUrl = _configuration["DeleteTeamUrl"];
            var opUrl = $"{deleteUrl}?teamsList={teamsList}";
            
            using (var client = new HttpClient())
            using (var request = new HttpRequestMessage(HttpMethod.Post, opUrl))
            {
                var responseMessage = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                                                                          .ConfigureAwait(false);
                response = await responseMessage.Content.ReadAsStringAsync();
            }

            return response;
        }

        private async Task<DialogTurnResult> FinalStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var teamsList = (string)stepContext.Result;
            var response = await DeleteTeams(teamsList);
            await stepContext.Context.SendActivityAsync(MessageFactory.Text(response));
            return await stepContext.EndDialogAsync(null, cancellationToken);
        }
    }
}
