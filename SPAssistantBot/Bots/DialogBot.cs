using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Logging;
using SPAssistantBot.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SPAssistantBot.Bots
{
    public class DialogBot<T> : ActivityHandler where T : Dialog
    {
        private readonly StateService _stateService;
        private readonly T _dialog;
        private readonly ILogger _logger;
        public DialogBot(StateService stateService, T dialog, ILogger<DialogBot<T>> logger)
        {
            _stateService = stateService;
            _dialog = dialog;
            _logger = logger;
        }

        public override  async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            await base.OnTurnAsync(turnContext, cancellationToken);

            await _stateService.UserState.SaveChangesAsync(turnContext, true, cancellationToken);
            await _stateService.ConversationState.SaveChangesAsync(turnContext, true, cancellationToken);
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            await _dialog.RunAsync(turnContext, _stateService.DialogStateAccessor, cancellationToken);
        }
    }
}
