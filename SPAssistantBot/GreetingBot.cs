using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using SPAssistantBot.Model;
using SPAssistantBot.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SPAssistantBot
{
    public class GreetingBot : ActivityHandler
    {
        private readonly StateService stateService;
        public GreetingBot(StateService stateService)
        {
            this.stateService = stateService;
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            await GetName(turnContext, cancellationToken);
        }

        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            foreach(var member in membersAdded)
            {
                if (member.Id != turnContext.Activity.Recipient.Id)
                {
                    await GetName(turnContext, cancellationToken);
                }
            }
        }

        private async Task GetName(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            var userProfile = await stateService.UserProfileAccessor.GetAsync(turnContext, () => new Model.UserProfile());
            var conversationData = await stateService.ConversationDataAccessor.GetAsync(turnContext, () => new ConversationData());

            if (!string.IsNullOrEmpty(userProfile.Name))
            {
                await turnContext.SendActivityAsync(MessageFactory.Text($"Hi {userProfile.Name}. How can I help you today?"), cancellationToken);
            }
            else if (conversationData.PromptedUserForName)
            {
                userProfile.Name = turnContext.Activity.Text?.Trim();
                conversationData.PromptedUserForName = false;
                await turnContext.SendActivityAsync(MessageFactory.Text($"Hi {userProfile.Name}. How can I help you?"), cancellationToken);
            }
            else
            {
                await turnContext.SendActivityAsync(MessageFactory.Text("What is your name?"), cancellationToken);
                conversationData.PromptedUserForName = true;
            }

            await stateService.UserProfileAccessor.SetAsync(turnContext, userProfile, cancellationToken);
            await stateService.ConversationDataAccessor.SetAsync(turnContext, conversationData, cancellationToken);

            await stateService.UserState.SaveChangesAsync(turnContext, true, cancellationToken);
            await stateService.ConversationState.SaveChangesAsync(turnContext, true, cancellationToken);
        }
    }
}
