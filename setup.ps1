#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs 7-Zip and Sublime Text, then disables Windows Smart App Control.

.DESCRIPTION
    - Downloads and installs 7-Zip (latest MSI)
    - Downloads and installs Sublime Text 4 (latest)
    - Disables Smart App Control via the registry

.NOTES
    Must be run as Administrator.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Helpers ──────────────────────────────────────────────────────────────────

function Write-Step  { param($msg) Write-Host "`n==> $msg" -ForegroundColor Cyan }
function Write-OK    { param($msg) Write-Host "    [OK] $msg" -ForegroundColor Green }
function Write-Fail  { param($msg) Write-Host "    [!!] $msg" -ForegroundColor Red }

function Install-FromUrl {
    param(
        [string]$Name,
        [string]$Url,
        [string]$OutFile,
        [string[]]$Args
    )

    Write-Step "Downloading $Name..."
    try {
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
        Write-OK "Downloaded to $OutFile"
    } catch {
        Write-Fail "Failed to download $Name`: $_"
        return $false
    }

    Write-Step "Installing $Name..."
    try {
        $proc = Start-Process -FilePath $OutFile -ArgumentList $Args -Wait -PassThru
        if ($proc.ExitCode -ne 0) {
            Write-Fail "$Name installer exited with code $($proc.ExitCode)"
            return $false
        }
        Write-OK "$Name installed successfully."
    } catch {
        Write-Fail "Failed to install $Name`: $_"
        return $false
    }

    Remove-Item -Path $OutFile -Force -ErrorAction SilentlyContinue
    return $true
}

# ── Temp folder ───────────────────────────────────────────────────────────────

$tmp = Join-Path $env:TEMP "AppInstalls_$(Get-Random)"
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

# ── 1. 7-Zip ──────────────────────────────────────────────────────────────────

Write-Step "Resolving latest 7-Zip MSI..."
try {
    # Scrape the official 7-zip download page to find the current x64 MSI link
    $page    = Invoke-WebRequest -Uri "https://www.7-zip.org/download.html" -UseBasicParsing
    $msiLink = ($page.Links | Where-Object { $_.href -match '7z\d+-x64\.msi$' } | Select-Object -First 1).href
    if (-not $msiLink) { throw "Could not locate MSI link on 7-zip.org" }
    $sevenZipUrl = "https://www.7-zip.org/$msiLink"
    Write-OK "Found: $sevenZipUrl"
} catch {
    # Fallback to a pinned recent version
    Write-Host "    [~] Falling back to pinned 7-Zip 24.09 URL" -ForegroundColor Yellow
    $sevenZipUrl = "https://www.7-zip.org/a/7z2409-x64.msi"
}

$sevenZipMsi = Join-Path $tmp "7zip-x64.msi"
Install-FromUrl -Name "7-Zip" `
                -Url $sevenZipUrl `
                -OutFile $sevenZipMsi `
                -Args @("/i", $sevenZipMsi, "/qn", "/norestart")

# ── 2. Sublime Text 4 ─────────────────────────────────────────────────────────

# Sublime Text's stable Windows x64 installer URL (always points to latest build)
$sublimeUrl = "https://download.sublimetext.com/sublime_text_build_4180_x64_setup.exe"
Write-Step "Resolving latest Sublime Text build..."
try {
    # Follow the redirect from the 'latest' alias
    $resp = Invoke-WebRequest -Uri "https://www.sublimetext.com/download_thanks?target=win-x64" `
                              -UseBasicParsing -MaximumRedirection 5
    $dlLink = ($resp.Links | Where-Object { $_.href -match 'sublime_text.*x64.*\.exe' } | Select-Object -First 1).href
    if ($dlLink) { $sublimeUrl = $dlLink; Write-OK "Found: $sublimeUrl" }
    else         { Write-Host "    [~] Using pinned Sublime Text URL" -ForegroundColor Yellow }
} catch {
    Write-Host "    [~] Using pinned Sublime Text URL" -ForegroundColor Yellow
}

$sublimeExe = Join-Path $tmp "SublimeTextSetup.exe"
Install-FromUrl -Name "Sublime Text 4" `
                -Url $sublimeUrl `
                -OutFile $sublimeExe `
                -Args @("/VERYSILENT", "/NORESTART", "/SUPPRESSMSGBOXES")

# ── 3. Disable Smart App Control ─────────────────────────────────────────────

Write-Step "Disabling Smart App Control..."
try {
    $sacKey = "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy"

    # VerifiedAndReputablePolicyState values:
    #   0 = Off  |  1 = Evaluation  |  2 = On
    if (-not (Test-Path $sacKey)) {
        New-Item -Path $sacKey -Force | Out-Null
    }

    Set-ItemProperty -Path $sacKey -Name "VerifiedAndReputablePolicyState" -Value 0 -Type DWord -Force
    Write-OK "Smart App Control set to OFF (registry key updated)."

    # Also stamp the 'was-turned-off-manually' flag so Windows doesn't re-enable it
    $sacFlag = "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy"
    Set-ItemProperty -Path $sacFlag -Name "VerifiedAndReputablePolicyStateMinValueSeen" -Value 0 -Type DWord -Force
    Write-OK "SAC minimum-seen state flag cleared."
} catch {
    Write-Fail "Could not update Smart App Control registry: $_"
}

# ── Done ──────────────────────────────────────────────────────────────────────

Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  All tasks completed." -ForegroundColor Cyan
Write-Host "  A reboot is recommended for SAC changes" -ForegroundColor Cyan
Write-Host "  to take full effect." -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan
