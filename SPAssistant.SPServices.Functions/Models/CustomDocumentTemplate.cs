namespace SPAssistant.SPServices.Functions.Models
{
    public class CustomDocumentTemplate
    {
        public CustomDocumentTemplate(string name, byte[] content, string containerName)
        {
            Name = name;
            Content = content;
            ContainerName = containerName;
        }

        public string Name { get; }

        public byte[] Content { get; }

        public string ContainerName { get; }
    }
}
