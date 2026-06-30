# EntraAgent

Demonstrates Microsoft Entra Agent Identities — a blueprint app authenticates with client credentials, exchanges for an agent identity token via FIC, then calls Microsoft Graph with delegated permissions acting as the agent's user account.

## One-time setup

### 1. Create the agent user account

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Connect-MgGraph -Scopes "User.ReadWrite.All" -TenantId {tenantId}

$result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/users" -Body (@{
    "@odata.type"     = "microsoft.graph.agentUser"
    displayName       = "ex1 Agent"
    userPrincipalName = "agent-{agentIdentityId[..8]}@{tenantDomain}"
    mailNickname      = "agent-{agentIdentityId[..8]}"
    accountEnabled    = $true
    identityParentId  = "{agentIdentityId}"
} | ConvertTo-Json) -ContentType "application/json"

$result.id  # save as AgentUserId
```

### 2. Grant delegated Graph consent to the agent user

```powershell
Connect-MgGraph -Scopes "DelegatedPermissionGrant.ReadWrite.All","Application.Read.All" -TenantId {tenantId}

$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$agentSp  = Get-MgServicePrincipal -Filter "id eq '{agentIdentityId}'"

New-MgOauth2PermissionGrant -BodyParameter @{
    clientId    = $agentSp.Id
    consentType = "Principal"
    principalId = "{agentUserId}"
    resourceId  = $graphSp.Id
    scope       = "Chat.Read Team.ReadBasic.All ChatMessage.Send ChannelMessage.Send"
}
```

### 3. Configure appsettings.local.json

| Key | Description |
| --- | ----------- |
| `AzureAd:TenantId` | Tenant ID |
| `AzureAd:ClientId` | Blueprint app registration client ID |
| `AzureAd:ClientCredentials[0]:ClientSecret` | Blueprint client secret |
| `AgentIdentityId` | Agent identity service principal object ID |
| `TenantDomain` | e.g. `contoso.onmicrosoft.com` |
| `AgentUserId` | Agent user account object ID (from step 2) |

## Run

```sh
dotnet run ./Agent.cs
```

Calls `GET /me/chats` (filtered to group chats) and `GET /me/joinedTeams` delegated as the agent user, and prints the results.
