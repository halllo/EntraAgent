#:package Microsoft.Identity.Web@4.11.0
#:package Microsoft.Identity.Web.GraphServiceClient@4.11.0
#:package Microsoft.Identity.Web.AgentIdentities@4.11.0
#:package Microsoft.Extensions.Hosting@10.0.9

using System.Security.Claims;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Abstractions;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.TokenCacheProviders.InMemory;
using Microsoft.Kiota.Abstractions.Authentication;

var host = Host.CreateDefaultBuilder(args)
    .ConfigureAppConfiguration(config =>
    {
        config.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
        config.AddJsonFile("appsettings.local.json", optional: true, reloadOnChange: true);
    })
    .ConfigureServices((context, services) =>
    {
        services.AddHostedService<AgentHostedService>();

        services.AddTokenAcquisition(isTokenAcquisitionSingleton: true)
            .Configure<MicrosoftIdentityApplicationOptions>(context.Configuration.GetSection("AzureAd"))
            .AddInMemoryTokenCaches()
            .AddHttpClient()
            .AddMicrosoftGraph()
            .AddAgentIdentities();

        services.AddKeyedSingleton<GraphServiceClient>("ex1", (sp, _) => CreateGraphClient(
            serviceProvider: sp,
            agentIdentityId: context.Configuration["AgentIdentityId"]!,
            agentUserId: context.Configuration["AgentUserId"]!));
    })
    .ConfigureLogging(logging =>
    {
        logging.ClearProviders();
        logging.AddConsole();
    })
    .Build();

await host.RunAsync();

static GraphServiceClient CreateGraphClient(IServiceProvider serviceProvider, string agentIdentityId, string agentUserId)
{
    var authProvider = serviceProvider.GetRequiredService<IAuthorizationHeaderProvider>();
    var tokenProvider = new BaseBearerTokenAuthenticationProvider(new AgentIdentityAccessTokenProvider(authProvider, agentIdentityId, agentUserId));
    return new GraphServiceClient(tokenProvider);
}

public class AgentIdentityAccessTokenProvider(IAuthorizationHeaderProvider authProvider, string agentIdentityId, string agentUserId) : IAccessTokenProvider
{
    public AllowedHostsValidator AllowedHostsValidator { get; } = new(["graph.microsoft.com"]);

    public async Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object>? additionalAuthenticationContext = null, CancellationToken cancellationToken = default)
    {
        var options = new AuthorizationHeaderProviderOptions().WithAgentUserIdentity(agentIdentityId, Guid.Parse(agentUserId));
        var header = await authProvider.CreateAuthorizationHeaderForUserAsync(["https://graph.microsoft.com/.default"], options, new ClaimsPrincipal(), cancellationToken);
        return header["Bearer ".Length..];
    }
}

public class AgentHostedService([FromKeyedServices("ex1")] GraphServiceClient graphClient, IHostApplicationLifetime lifetime) : BackgroundService
{
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        var users = await graphClient.Users.GetAsync();

        Console.WriteLine($"Found {users?.Value?.Count ?? 0} user(s) in the tenant.");
        foreach (var user in users?.Value ?? [])
        {
            Console.WriteLine($"- {user.DisplayName} ({user.UserPrincipalName})");
        }

        lifetime.StopApplication();
    }
}
