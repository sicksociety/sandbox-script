#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs 7-Zip and Sublime Text 4, then disables Windows Smart App Control.
.NOTES
    Run as Administrator:
        Set-ExecutionPolicy Bypass -Scope Process -Force
        .\setup.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Helpers ───────────────────────────────────────────────────────────────────

function Write-Step { param($msg) Write-Host "`n==> $msg" -ForegroundColor Cyan   }
function Write-OK   { param($msg) Write-Host "    [OK] $msg" -ForegroundColor Green  }
function Write-Warn { param($msg) Write-Host "    [~]  $msg" -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host "    [!!] $msg" -ForegroundColor Red    }

# NOTE: Do NOT name the args parameter $Args — that is a reserved PowerShell variable.
function Install-Exe {
    param(
        [string]   $AppName,
        [string]   $Url,
        [string]   $Dest,
        [string[]] $SilentArgs
    )

    Write-Step "Downloading $AppName..."
    try {
        Invoke-WebRequest -Uri $Url -OutFile $Dest -UseBasicParsing
        Write-OK "Saved to $Dest"
    } catch {
        Write-Fail "Download failed: $_"
        return
    }

    Write-Step "Installing $AppName..."
    try {
        # Choose the right host process for MSI vs EXE
        if ($Dest -match '\.msi$') {
            $proc = Start-Process -FilePath "msiexec.exe" `
                                  -ArgumentList (@("/i", "`"$Dest`"") + $SilentArgs) `
                                  -Wait -PassThru
        } else {
            $proc = Start-Process -FilePath $Dest `
                                  -ArgumentList $SilentArgs `
                                  -Wait -PassThru
        }

        if ($proc.ExitCode -eq 0) {
            Write-OK "$AppName installed successfully."
        } else {
            Write-Fail "$AppName exited with code $($proc.ExitCode)"
        }
    } catch {
        Write-Fail "Install failed: $_"
    }

    Remove-Item -Path $Dest -Force -ErrorAction SilentlyContinue
}

# ── Temp folder ───────────────────────────────────────────────────────────────

$tmp = Join-Path $env:TEMP "AppInstalls_$(Get-Random)"
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

# ─────────────────────────────────────────────────────────────────────────────
# 1. 7-Zip  —  resolve via GitHub Releases API (always accurate, no scraping)
# ─────────────────────────────────────────────────────────────────────────────

Write-Step "Resolving latest 7-Zip version from GitHub..."

$sevenZipUrl = $null
try {
    $release = Invoke-RestMethod `
        -Uri "https://api.github.com/repos/ip7z/7zip/releases/latest" `
        -UseBasicParsing
    $asset = $release.assets | Where-Object { $_.name -match 'x64\.msi$' } | Select-Object -First 1
    if ($asset) {
        $sevenZipUrl = $asset.browser_download_url
        Write-OK "Latest release: $($release.tag_name)  ->  $sevenZipUrl"
    }
} catch {
    Write-Warn "GitHub API failed: $_"
}

if (-not $sevenZipUrl) {
    $sevenZipUrl = "https://github.com/ip7z/7zip/releases/download/26.00/7z2600-x64.msi"
    Write-Warn "Using pinned fallback: $sevenZipUrl"
}

Install-Exe -AppName "7-Zip" `
            -Url $sevenZipUrl `
            -Dest (Join-Path $tmp "7zip-x64.msi") `
            -SilentArgs @("/qn", "/norestart")

# ─────────────────────────────────────────────────────────────────────────────
# 2. Sublime Text 4  —  resolve via official update-check feed
# ─────────────────────────────────────────────────────────────────────────────

Write-Step "Resolving latest Sublime Text 4 build..."

$sublimeUrl = $null
try {
    $feed = Invoke-RestMethod `
        -Uri "https://www.sublimetext.com/updates/4/stable_update_check?platform=windows&version=1" `
        -UseBasicParsing
    if ($feed.url) {
        $sublimeUrl = $feed.url
        Write-OK "Found: $sublimeUrl"
    }
} catch {
    Write-Warn "Sublime update feed failed: $_"
}

if (-not $sublimeUrl) {
    $sublimeUrl = "https://download.sublimetext.com/sublime_text_build_4200_x64_setup.exe"
    Write-Warn "Using pinned fallback: $sublimeUrl"
}

Install-Exe -AppName "Sublime Text 4" `
            -Url $sublimeUrl `
            -Dest (Join-Path $tmp "SublimeTextSetup.exe") `
            -SilentArgs @("/VERYSILENT", "/NORESTART", "/SUPPRESSMSGBOXES")

# ─────────────────────────────────────────────────────────────────────────────
# 3. Disable Smart App Control
# ─────────────────────────────────────────────────────────────────────────────

Write-Step "Disabling Smart App Control..."
try {
    $sacKey = "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy"

    if (-not (Test-Path $sacKey)) {
        New-Item -Path $sacKey -Force | Out-Null
    }

    # 0 = Off | 1 = Evaluation | 2 = On
    Set-ItemProperty -Path $sacKey -Name "VerifiedAndReputablePolicyState"             -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $sacKey -Name "VerifiedAndReputablePolicyStateMinValueSeen" -Value 0 -Type DWord -Force
    Write-OK "Smart App Control set to OFF."
} catch {
    Write-Fail "Registry update failed: $_"
}

# ── Cleanup & summary ─────────────────────────────────────────────────────────

Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Done! Reboot to fully apply SAC change." -ForegroundColor Cyan
Write-Host "========================================`n"  -ForegroundColor Cyan
