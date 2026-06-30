# EntraAgent

Demonstrates Microsoft Entra Agent Identities — a blueprint app authenticates with client credentials, exchanges for an agent identity token via FIC, then calls Microsoft Graph with delegated permissions acting as the agent's user account.

## One-time setup

### 1. Create the Agent Identity Blueprint and Agent Identity

In the [Microsoft Entra admin center](https://entra.microsoft.com):

1. Create the blueprint — go to Entra ID > Agents > Agent blueprints > New agent blueprint (Preview). On the blueprint's detail page add a client secret.
2. Create the agent identity — go to Entra ID > Agents > Agent identities > New agent identity (Preview). Select the blueprint from step 1 under Agent blueprint.

### 2. Create the agent user account

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

### 3. Grant delegated Graph consent to the agent user

```powershell
Connect-MgGraph -Scopes "DelegatedPermissionGrant.ReadWrite.All","Application.Read.All" -TenantId {tenantId}

$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$agentSp  = Get-MgServicePrincipal -Filter "id eq '{agentIdentityId}'"

New-MgOauth2PermissionGrant -BodyParameter @{
    clientId    = $agentSp.Id
    consentType = "Principal"
    principalId = "{agentUserId}"
    resourceId  = $graphSp.Id
    scope       = "User.Read Chat.Read Team.ReadBasic.All ChatMessage.Send ChannelMessage.Send"
}
```

### 4. Assign a Teams license to the agent user

Teams Graph endpoints (`/me/chats`, `/me/joinedTeams`) require the agent user to hold a valid Teams/Microsoft 365 license, or they fail with "Failed to get license information for the user.".

### 5. Configure appsettings.local.json

| Key | Description |
| --- | ----------- |
| `AzureAd:TenantId` | Tenant ID |
| `AzureAd:ClientId` | Blueprint app registration client ID |
| `AzureAd:ClientCredentials[0]:ClientSecret` | Blueprint client secret |
| `AgentIdentityId` | Agent identity service principal object ID |
| `AgentUserId` | Agent user account object ID (from step 2) |

## Run

```sh
dotnet run ./Agent.cs
```

Prints the agent user's profile (`GET /me`), then calls `GET /me/chats` (filtered to group chats) and `GET /me/joinedTeams` delegated as the agent user, and prints the results.
