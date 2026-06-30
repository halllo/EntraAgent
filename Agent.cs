#:package Microsoft.Identity.Web@4.11.0
#:package Microsoft.Identity.Web.AgentIdentities@4.11.0
#:package Microsoft.Identity.Web.GraphServiceClient@4.11.0
#:package Microsoft.Extensions.Configuration.UserSecrets@10.0.9
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
        config.AddJsonFile("appsettings.local.json", optional: false, reloadOnChange: true);
        config.AddUserSecrets("Agent-dc989f5e0c3e979d1dacbcc4ad52364c3ec39878e669c6a02ca8694f37c2cbbf");
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

public class AgentHostedService([FromKeyedServices("ex1")] GraphServiceClient graphClient, IHostApplicationLifetime lifetime, IConfiguration config) : BackgroundService
{
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        // Print details about the account the agent is currently acting on behalf of.
        var me = await graphClient.Me.GetAsync(cancellationToken: stoppingToken);
        Console.WriteLine("Signed in as:");
        Console.WriteLine($"- Display name:    {me?.DisplayName ?? "(unknown)"}");
        Console.WriteLine($"- User principal:  {me?.UserPrincipalName ?? "(unknown)"}");
        Console.WriteLine($"- Mail:            {me?.Mail ?? "(none)"}");
        Console.WriteLine($"- Job title:       {me?.JobTitle ?? "(none)"}");
        Console.WriteLine($"- Object id:       {me?.Id ?? "(unknown)"}");
        Console.WriteLine();

        // List all group chats the user is part of (chatType 'group', excludes 1:1 and meeting chats).
        var groupChats = await graphClient.Me.Chats.GetAsync(requestConfig =>
        {
            requestConfig.QueryParameters.Filter = "chatType eq 'group'";
        }, stoppingToken);
        Console.WriteLine($"Found {groupChats?.Value?.Count ?? 0} group chat(s).");
        foreach (var chat in groupChats?.Value ?? [])
        {
            Console.WriteLine($"- {chat.Topic ?? "(no topic)"} [{chat.Id}]");
        }

        // List all teams the user has joined.
        var teams = await graphClient.Me.JoinedTeams.GetAsync(cancellationToken: stoppingToken);
        Console.WriteLine($"Found {teams?.Value?.Count ?? 0} team(s).");
        foreach (var team in teams?.Value ?? [])
        {
            Console.WriteLine($"- {team.DisplayName} [{team.Id}]");
        }

        lifetime.StopApplication();
    }
}
