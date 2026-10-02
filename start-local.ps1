[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$projectDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $projectDirectory

foreach ($command in @("node", "npm", "az")) {
    if (-not (Get-Command $command -ErrorAction SilentlyContinue)) {
        throw "'$command' is required but was not found in PATH."
    }
}

$dependenciesReady = Test-Path -LiteralPath "node_modules\.bin\vite.cmd"
if ($dependenciesReady) {
    node -e "import('vite').then(() => process.exit(0)).catch(() => process.exit(1))"
    $dependenciesReady = $LASTEXITCODE -eq 0
}

if (-not $dependenciesReady) {
    Write-Host "Installing application dependencies..."
    npm ci
    if ($LASTEXITCODE -ne 0) {
        throw "npm ci failed with exit code $LASTEXITCODE."
    }
}

az account show --only-show-errors --output none 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Sign in with the Microsoft account that can access the Fabric Data Agent."
    az login --output none
    if ($LASTEXITCODE -ne 0) {
        throw "Azure CLI sign-in failed."
    }
}

Write-Host "Getting a short-lived Fabric access token..."
$accessToken = az account get-access-token `
    --scope "https://api.fabric.microsoft.com/.default" `
    --query accessToken `
    --output tsv `
    --only-show-errors

if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($accessToken)) {
    throw "Could not obtain a Fabric access token. Confirm that your signed-in account can access the published Data Agent."
}

$env:FABRIC_ACCESS_TOKEN = $accessToken.Trim()
$env:FABRIC_MCP_URL = "https://api.fabric.microsoft.com/v1/mcp/workspaces/f67b0ccc-1f83-430c-93b4-0cbfc45c7f3a/dataagents/5071fab0-9cf5-40d0-9e0c-26dc4d9398a2/agent"
$env:VITE_LOCAL_AUTH_BYPASS = "true"

Write-Host "Starting DCT Contract Assistant..."
Write-Host "Open http://localhost:5173 in your browser. Press Ctrl+C to stop."

try {
    npm run dev
} finally {
    Remove-Item Env:FABRIC_ACCESS_TOKEN -ErrorAction SilentlyContinue
    Remove-Item Env:FABRIC_MCP_URL -ErrorAction SilentlyContinue
    Remove-Item Env:VITE_LOCAL_AUTH_BYPASS -ErrorAction SilentlyContinue
}
