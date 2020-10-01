using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Extensions.Configuration;
using SPAssistantBot.Services;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace SPAssistantBot.Dialogs
{
    public class MainDialog : ComponentDialog
    {
        private readonly StateService _stateService;
        private readonly BotServices _botServices;
        private readonly SPService _spServices;
        private readonly TeamsService _teamsServices;
        private readonly IConfiguration _configuration;

        public MainDialog(StateService stateService, BotServices botservices, SPService spService,  TeamsService teamsServices, IConfiguration configuration)
        {
            _stateService = stateService ?? throw new ArgumentNullException(nameof(stateService));
            _botServices = botservices ?? throw new ArgumentNullException(nameof(botservices));
            _spServices = spService ?? throw new ArgumentNullException(nameof(spService));
            _teamsServices = teamsServices ?? throw new ArgumentNullException(nameof(teamsServices));
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));

            InitializeWaterfallDialog();
        }

        private void InitializeWaterfallDialog()
        {
            List<WaterfallStep> waterfallSteps = new List<WaterfallStep>
            {
                InitialStepAsync,
                FinalStepAsync
            };

            AddDialog(new GreetingDialog($"{nameof(MainDialog)}.greeting", _stateService));
            AddDialog(new DeleteTeamDialog($"{nameof(MainDialog)}.deleteMSTeam", _teamsServices, _configuration));
            AddDialog(new SPAssistantDialog($"{nameof(MainDialog)}.createspsite", _spServices, _configuration));
            AddDialog(new TeamsAssistantDialog($"{nameof(MainDialog)}.createMSTeam",  _teamsServices, _configuration));
            AddDialog(new WaterfallDialog($"{nameof(MainDialog)}.mainflow", waterfallSteps));

            InitialDialogId = $"{nameof(MainDialog)}.mainflow";

        }

        private async Task<DialogTurnResult> InitialStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var recognizerResult = await _botServices.Dispatch.RecognizeAsync(stepContext.Context, cancellationToken);

            var topIntent = recognizerResult.GetTopScoringIntent();

            switch (topIntent.intent)
            {
                case "Greeting":
                    return await stepContext.BeginDialogAsync($"{nameof(MainDialog)}.greeting", null, cancellationToken);
                case "CreateSiteIntent":
                    return await stepContext.BeginDialogAsync($"{nameof(MainDialog)}.createspsite", null, cancellationToken);
                case "CreateTeamIntent":
                    return await stepContext.BeginDialogAsync($"{nameof(MainDialog)}.createMSTeam", null, cancellationToken);
                case "DeleteTeamIntent":
                    return await stepContext.BeginDialogAsync($"{nameof(MainDialog)}.deleteMSTeam", null, cancellationToken);
                default:
                    await stepContext.Context.SendActivityAsync(MessageFactory.Text("I am sorry, I don't know what you mean."), cancellationToken);
                    break;
            }

            return await stepContext.NextAsync(null, cancellationToken);
        }

        private async Task<DialogTurnResult> FinalStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            return await stepContext.EndDialogAsync(null, cancellationToken);
        }
    }
}
