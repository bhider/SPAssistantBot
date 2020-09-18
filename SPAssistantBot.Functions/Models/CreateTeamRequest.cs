namespace SPAssistantBot.Functions.Models
{
    public class CreateTeamRequest
    {
        public string TeamName { get; set; }

        public string Description { get; set; }

        public string TeamType { get; set; }

        public string OwnersUserEmailListAsString { get; set; }

        public string MembersUserEmailListAsString { get; set; }

        public bool UseTemplate { get; set; }

        public string TemplateName { get; set; }
    }
}
