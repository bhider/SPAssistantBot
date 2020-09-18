using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SPAssistantBot.Helpers
{
    public static class DialogExtensions
    {

        public static async Task Run(this Dialog dialog, ITurnContext turnContext, IStatePropertyAccessor<DialogState> dialogStateAccessor, CancellationToken cancellationToken)
        {
            var dialogSet = new DialogSet(dialogStateAccessor);
            dialogSet.Add(dialog);

            var dialogContext = await dialogSet.CreateContextAsync(turnContext, cancellationToken);
            

            var dialogStatus = await dialogContext.ContinueDialogAsync(cancellationToken);

            if (dialogStatus.Status == DialogTurnStatus.Empty)
            {
                await dialogContext.BeginDialogAsync(dialog.Id, null, cancellationToken);
            }
        }
    }
}
