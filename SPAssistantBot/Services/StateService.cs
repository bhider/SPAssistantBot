using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using SPAssistantBot.Model;
using System;

namespace SPAssistantBot.Services
{
    public class StateService
    {
        public UserState    UserState { get; set; }

        public ConversationState ConversationState { get; set; }

        public static string UserProfileId { get; } = $"{nameof(StateService)}.UserProfile";

        public static string ConversationDataId { get; } = $"{nameof(StateService)}.ConversationData";

        public static string DialogStateId { get; } = $"{nameof(StateService)}.DialogState";

        public IStatePropertyAccessor<UserProfile> UserProfileAccessor { get; set; }

        public IStatePropertyAccessor<ConversationData> ConversationDataAccessor { get; set; }

        public IStatePropertyAccessor<DialogState> DialogStateAccessor { get; set; }

        public StateService(UserState userState, ConversationState conversationState)
        {
            UserState = userState ?? throw new ArgumentNullException(nameof(userState));
            ConversationState = conversationState ?? throw new ArgumentNullException(nameof(conversationState));
            InitializeAccessors();
        }

        private void InitializeAccessors()
        {
            UserProfileAccessor = UserState.CreateProperty<UserProfile>(UserProfileId);
            ConversationDataAccessor = ConversationState.CreateProperty<ConversationData>(ConversationDataId);
            DialogStateAccessor = ConversationState.CreateProperty<DialogState>(DialogStateId);
        }
    }
}
