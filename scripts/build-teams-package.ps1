[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$projectDirectory = Split-Path -Parent $PSScriptRoot
$manifestPath = Join-Path $projectDirectory "teams\manifest.json"
$manifest = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json
$artifactDirectory = Join-Path $projectDirectory "artifacts"
New-Item -ItemType Directory -Path $artifactDirectory -Force | Out-Null
$stagingDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ("dct-teams-" + [guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $stagingDirectory | Out-Null

Add-Type -AssemblyName System.Drawing
$generatedFiles = @()
try {
    $manifestCopy = Join-Path $stagingDirectory "manifest.json"
    Copy-Item -LiteralPath $manifestPath -Destination $manifestCopy
    $generatedFiles += $manifestCopy
    foreach ($icon in @(@{ Name = "color.png"; Size = 192 }, @{ Name = "outline.png"; Size = 32 })) {
        $bitmap = [System.Drawing.Bitmap]::new($icon.Size, $icon.Size)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.Clear([System.Drawing.Color]::Transparent)
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            if ($icon.Size -eq 192) {
                $graphics.Clear([System.Drawing.ColorTranslator]::FromHtml("#315E4D"))
            }
            # Reuse the three-bar mark from the website as a simple Teams icon.
            for ($bar = 0; $bar -lt 3; $bar++) {
                $x = [single]($icon.Size * (0.2 + $bar * 0.22))
                $y = [single]($icon.Size * (0.3 - $bar * 0.05))
                $graphics.FillRectangle([System.Drawing.Brushes]::White, $x, $y, [single]($icon.Size * 0.16), [single]($icon.Size * 0.5))
            }
            $iconPath = Join-Path $stagingDirectory $icon.Name
            $bitmap.Save($iconPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $generatedFiles += $iconPath
        } finally {
            $graphics.Dispose()
            $bitmap.Dispose()
        }
    }
    $packagePath = Join-Path $artifactDirectory ("dct-teams-tab-" + $manifest.version + ".zip")
    Compress-Archive -LiteralPath $generatedFiles -DestinationPath $packagePath -Force
    Write-Host "Teams tab package: $packagePath"
} finally {
    # Remove only the files created by this run; no recursive deletion.
    foreach ($generatedFile in $generatedFiles) {
        Remove-Item -LiteralPath $generatedFile -ErrorAction SilentlyContinue
    }
    Remove-Item -LiteralPath $stagingDirectory -ErrorAction SilentlyContinue
}
