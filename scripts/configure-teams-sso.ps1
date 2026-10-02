[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ClientId = "54a740da-a7cc-40f3-8951-247f8ff1c307",
    [string]$TenantId = "a05fd237-40b5-40b4-bae8-c577a96df96c",
    [uri]$TabUrl
)

$ErrorActionPreference = "Stop"

function Invoke-AzJson {
    param([string[]]$Arguments)

    $result = & az @Arguments --only-show-errors --output json
    if ($LASTEXITCODE -ne 0) {
        throw "Azure CLI failed. Resolve the reported error before running this script again."
    }
    return (($result -join "`n") | ConvertFrom-Json)
}

function Update-AppRegistration {
    param([string]$ObjectId, [hashtable]$Properties)

    # Passing JSON through a file avoids PowerShell/native-command quoting issues.
    $payloadPath = [System.IO.Path]::GetTempFileName()
    try {
        $payload = $Properties | ConvertTo-Json -Depth 50
        [System.IO.File]::WriteAllText(
            $payloadPath,
            $payload,
            [System.Text.UTF8Encoding]::new($false)
        )
        & az rest --method PATCH `
            --url "https://graph.microsoft.com/v1.0/applications/$ObjectId" `
            --headers "Content-Type=application/json" `
            --body "@$payloadPath" --only-show-errors --output none
        if ($LASTEXITCODE -ne 0) {
            throw "Could not update the application. Existing successful steps are preserved; the script can be rerun."
        }
    } finally {
        Remove-Item -LiteralPath $payloadPath -ErrorAction SilentlyContinue
    }
}

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    throw "Azure CLI is not available in PATH."
}

$account = Invoke-AzJson -Arguments @("account", "show")
if ($account.tenantId -ne $TenantId) {
    throw "Wrong active tenant. Sign in with: az login --tenant $TenantId"
}

$registration = Invoke-AzJson -Arguments @("ad", "app", "show", "--id", $ClientId)
$resourceUri = "api://$ClientId"
if ($TabUrl) {
    if (-not $TabUrl.IsAbsoluteUri -or $TabUrl.Scheme -ne "https" -or -not $TabUrl.IsDefaultPort -or $TabUrl.UserInfo) {
        throw "TabUrl must be a public HTTPS URL without credentials or a custom port."
    }
    $resourceUri = "api://$($TabUrl.Host)/$ClientId"
}
if ($registration.api.requestedAccessTokenVersion -ne 2) {
    throw "Version 2 access tokens must be configured first."
}
if (-not $TabUrl -and $registration.identifierUris -notcontains $resourceUri) {
    throw "Configure the application identifier first, or provide -TabUrl for the Teams tab."
}

$scopes = @($registration.api.oauth2PermissionScopes | Where-Object { $null -ne $_ })
$matchingScopes = @($scopes | Where-Object { $_.value -eq "access_as_user" })
if ($matchingScopes.Count -gt 1) {
    throw "Multiple access_as_user scopes found. Review the registration before continuing."
}

if ($matchingScopes.Count -eq 1) {
    $scope = $matchingScopes[0]
    if (-not $scope.isEnabled) {
        throw "The existing access_as_user scope is disabled. Review it before continuing."
    }
} else {
    $scope = [pscustomobject]@{
        id = [guid]::NewGuid().ToString()
        value = "access_as_user"
        type = "User"
        isEnabled = $true
        adminConsentDisplayName = "Access DCT Teams Assistant"
        adminConsentDescription = "Allow the application to access DCT Teams Assistant on behalf of the signed-in user."
        userConsentDisplayName = "Access DCT Teams Assistant"
        userConsentDescription = "Allow the application to access DCT Teams Assistant on your behalf."
    }
}

if (-not $PSCmdlet.ShouldProcess(
    "$($registration.displayName) ($ClientId)",
    "Ensure SSO resource $resourceUri, access_as_user scope, and Teams client pre-authorization"
)) {
    return
}

if ($registration.identifierUris -notcontains $resourceUri) {
    $identifierUris = @($registration.identifierUris) + @($resourceUri)
    Update-AppRegistration -ObjectId $registration.id -Properties @{ identifierUris = $identifierUris }
}

if ($matchingScopes.Count -eq 0) {
    $registration.api.oauth2PermissionScopes = @($scopes) + @($scope)
    Update-AppRegistration -ObjectId $registration.id -Properties @{ api = $registration.api }
}

# Read back before adding clients that reference the saved scope ID.
$registration = Invoke-AzJson -Arguments @("ad", "app", "show", "--id", $ClientId)
$savedScope = @($registration.api.oauth2PermissionScopes | Where-Object { $_.value -eq "access_as_user" })
if ($savedScope.Count -ne 1 -or $savedScope[0].id -ne $scope.id) {
    throw "The new scope is not visible yet. Wait briefly and rerun this script."
}

$clients = @($registration.api.preAuthorizedApplications | Where-Object { $null -ne $_ })
$teamsClientIds = @(
    "1fec8e78-bce4-4aaf-ab1b-5451cc387264", # Teams desktop/mobile
    "5e3ce6c0-2b1f-4285-8d4b-75ee78787346"  # Teams web
)

foreach ($teamsClientId in $teamsClientIds) {
    $existing = @($clients | Where-Object { $_.appId -eq $teamsClientId })
    if ($existing.Count -gt 1) {
        throw "Duplicate pre-authorization entries found for $teamsClientId."
    }
    if ($existing.Count -eq 1) {
        if ($existing[0].delegatedPermissionIds -notcontains $scope.id) {
            $existing[0].delegatedPermissionIds = @($existing[0].delegatedPermissionIds) + @($scope.id)
        }
    } else {
        $clients += [pscustomobject]@{
            appId = $teamsClientId
            delegatedPermissionIds = @($scope.id)
        }
    }
}

$registration.api.preAuthorizedApplications = $clients
Update-AppRegistration -ObjectId $registration.id -Properties @{ api = $registration.api }

$verified = Invoke-AzJson -Arguments @("ad", "app", "show", "--id", $ClientId)
if ($verified.identifierUris -notcontains $resourceUri) {
    throw "The expected SSO resource URI was not found."
}
foreach ($teamsClientId in $teamsClientIds) {
    $entry = @($verified.api.preAuthorizedApplications | Where-Object {
        $_.appId -eq $teamsClientId -and $_.delegatedPermissionIds -contains $scope.id
    })
    if ($entry.Count -ne 1) {
        throw "Pre-authorization is not visible for $teamsClientId yet. Wait briefly and rerun."
    }
}

Write-Host "SSO scope and Teams client pre-authorization verified."
Write-Host "Application ID: $ClientId"
Write-Host "SSO resource URI: $resourceUri"
Write-Host "Scope: $resourceUri/access_as_user"
if ($TabUrl) {
    Write-Host "Remaining: Fabric service principal access, server credential, Teams tab package, and deployment."
} else {
    Write-Host "Remaining: Fabric service principal access, Teams manifest, and backend deployment."
}
