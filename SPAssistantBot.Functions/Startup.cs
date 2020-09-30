using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using SPAssistantBot.Services;
using System;
using System.Collections.Generic;
using System.Text;

[assembly: FunctionsStartup(typeof(SPAssistantBot.Functions.Startup))]
namespace SPAssistantBot.Functions
{
    //https://docs.microsoft.com/en-us/azure/azure-functions/functions-dotnet-dependency-injection
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            builder.Services.AddSingleton<KeyVaultService>();
            builder.Services.AddSingleton<SPService>();
            builder.Services.AddSingleton<TeamsService>();
            builder.Services.AddSingleton<SPCustomisationService>();
        }
    }
}
