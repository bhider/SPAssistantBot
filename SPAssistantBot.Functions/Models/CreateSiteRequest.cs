namespace SPAssistantBot.Functions.Models
{
    class CreateSiteRequest
    {
        public string SiteTitle { get; set; }

        public string Description { get; set; }

        public string SiteType { get; set; }

        public string OwnersUserEmailListAsString { get; set; }

        public string MembersUserEmailListAsString { get; set; }
    }
}
