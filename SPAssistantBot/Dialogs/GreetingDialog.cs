using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using SPAssistantBot.Services;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace SPAssistantBot.Dialogs
{
    public class GreetingDialog : ComponentDialog
    {
        private readonly StateService _stateService;

        public GreetingDialog(string dialogId, StateService stateService) : base(dialogId)
        {
            _stateService = stateService ?? throw new ArgumentNullException(nameof(stateService));

            InitializeWaterfallDialog();
        }

        private void InitializeWaterfallDialog()
        {
            var waterfallSteps = new WaterfallStep[]
            {
                InitialStepAsync,
                FinalStepAsync
            };
            AddDialog(new TextPrompt($"{nameof(GreetingDialog)}.name"));

            AddDialog(new WaterfallDialog($"{nameof(GreetingDialog)}.mainflow", waterfallSteps));

            InitialDialogId = $"{nameof(GreetingDialog)}.mainflow";
        }

        private async Task<DialogTurnResult> InitialStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var userProfile = await _stateService.UserProfileAccessor.GetAsync(stepContext.Context, () => new Model.UserProfile());

            if (string.IsNullOrWhiteSpace(userProfile.Name))
            {
                return await  stepContext.PromptAsync($"{nameof(GreetingDialog)}.name", new PromptOptions
                {
                    Prompt = MessageFactory.Text("What is your Name?")
                }, cancellationToken) ;
            }

            return await stepContext.NextAsync(null, cancellationToken);
        }

        private async Task<DialogTurnResult> FinalStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var userProfile = await _stateService.UserProfileAccessor.GetAsync(stepContext.Context, () => new Model.UserProfile());

            if (string.IsNullOrWhiteSpace(userProfile.Name))
            {
                userProfile.Name = (string)stepContext.Result;

                await _stateService.UserProfileAccessor.SetAsync(stepContext.Context, userProfile, cancellationToken);

            }
            
            await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Hi {userProfile.Name}. How can I help you today?"), cancellationToken);
            
            return await stepContext.EndDialogAsync(null, cancellationToken);
        }

    }
}
