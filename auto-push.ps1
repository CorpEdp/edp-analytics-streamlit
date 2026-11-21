# ============================================
# EDP Analytics - Automatic GitHub Watcher
# ============================================

$ErrorActionPreference = "Stop"
$env:GIT_TERMINAL_PROMPT = "0"

# The folder containing this script
$RepoPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Git settings
$Branch = "main"

# How often to check for changes (seconds)
$CheckInterval = 5
$RetryInterval = 30
$LogDirectory = Join-Path $env:LOCALAPPDATA "EDP Analytics"
$LogPath = Join-Path $LogDirectory "auto-push.log"

New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null

function Write-Log {
    param([string]$Message)

    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

function Invoke-Git {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $output = & git @Arguments 2>&1
    return [PSCustomObject]@{
        ExitCode = $LASTEXITCODE
        Output = @($output | ForEach-Object { "$_" })
    }
}

# Prevent duplicate watcher processes after logon or an accidental manual launch.
$mutexCreated = $false
$mutex = New-Object System.Threading.Mutex($false, "Global\EDP-Analytics-Auto-Push", [ref]$mutexCreated)
if (-not $mutexCreated) {
    Write-Host "Another EDP Analytics watcher is already running."
    exit 0
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "     EDP ANALYTICS AUTO GITHUB WATCHER" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Repository: $RepoPath"
Write-Host "Branch:     $Branch"
Write-Host "Checking:   Every $CheckInterval seconds"
Write-Host ""
Write-Host "Press CTRL+C to stop."
Write-Host ""

# Move into repository
Set-Location $RepoPath

# Verify this is a Git repository
if (-not (Test-Path ".git")) {
    Write-Log "ERROR: $RepoPath is not a Git repository."
    exit 1
}

# Make sure we are on the correct branch
$branchResult = Invoke-Git @("branch", "--show-current")
if ($branchResult.ExitCode -ne 0) {
    Write-Log "ERROR: Could not determine the current Git branch."
    exit 1
}
$currentBranch = ($branchResult.Output -join "").Trim()

if ($currentBranch -ne $Branch) {
    Write-Host "Switching to branch: $Branch" -ForegroundColor Yellow
    $checkoutResult = Invoke-Git @("checkout", $Branch)

    if ($checkoutResult.ExitCode -ne 0) {
        $checkoutResult.Output | ForEach-Object { Write-Log "  $_" }
        Write-Log "ERROR: Could not switch to $Branch; resolve local changes and try again."
        exit 1
    }
}

Write-Log "Watcher started. Repository: $RepoPath; branch: $Branch."

while ($true) {
    $statusResult = Invoke-Git @("status", "--porcelain", "--untracked-files=normal")
    if ($statusResult.ExitCode -ne 0) {
        $statusResult.Output | ForEach-Object { Write-Log "  $_" }
        Write-Log "ERROR: git status failed; will retry."
        Start-Sleep -Seconds $RetryInterval
        continue
    }
    $status = $statusResult.Output

    if ($status) {
        Write-Log "Change detected:"
        $shortStatusResult = Invoke-Git @("status", "--short")
        $shortStatusResult.Output | ForEach-Object { Write-Log $_ }

        # Wait briefly so multiple files saved together
        Start-Sleep -Seconds 2

        $statusAfterWaitResult = Invoke-Git @("status", "--porcelain", "--untracked-files=normal")
        if ($statusAfterWaitResult.ExitCode -ne 0) {
            $statusAfterWaitResult.Output | ForEach-Object { Write-Log "  $_" }
            Write-Log "ERROR: Could not recheck Git status; will retry."
            Start-Sleep -Seconds $RetryInterval
            continue
        }
        $statusAfterWait = $statusAfterWaitResult.Output

        if ($statusAfterWait) {
            $addResult = Invoke-Git @("add", "-A")

            if ($addResult.ExitCode -ne 0) {
                $addResult.Output | ForEach-Object { Write-Log "  $_" }
                Write-Log "ERROR: git add failed; will retry."
                Start-Sleep -Seconds $RetryInterval
                continue
            }

            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $commitMessage = "Auto update EDP Analytics - $timestamp"
            $commitResult = Invoke-Git @("commit", "-m", $commitMessage)

            if ($commitResult.ExitCode -ne 0) {
                $commitResult.Output | ForEach-Object { Write-Log "  $_" }
                Write-Log "ERROR: git commit failed; staged changes will be retried."
            }
            else {
                $commitResult.Output | ForEach-Object { Write-Log "$_" }
                Write-Log "Committed local changes."
            }
        }
    }

    # Fetch and push independently of worktree changes so a failed push is retried.
    $fetchResult = Invoke-Git @("fetch", "--prune", "origin")
    if ($fetchResult.ExitCode -ne 0) {
        $fetchResult.Output | ForEach-Object { Write-Log "  $_" }
        Write-Log "ERROR: git fetch failed; will retry."
        Start-Sleep -Seconds $RetryInterval
        continue
    }

    $divergenceResult = Invoke-Git @("rev-list", "--left-right", "--count", "HEAD...origin/$Branch")
    if ($divergenceResult.ExitCode -ne 0) {
        $divergenceResult.Output | ForEach-Object { Write-Log "  $_" }
        Write-Log "ERROR: Could not compare local and remote commits; will retry."
        Start-Sleep -Seconds $RetryInterval
        continue
    }
    $divergence = ($divergenceResult.Output -join "").Trim()

    $commitCounts = "$divergence" -split '\s+'
    if ($commitCounts.Count -ne 2) {
        Write-Log "ERROR: Unexpected output while comparing local and remote commits: $divergence"
        Start-Sleep -Seconds $RetryInterval
        continue
    }

    $localOnly = [int]$commitCounts[0]
    $remoteOnly = [int]$commitCounts[1]

    if ($remoteOnly -gt 0 -and $localOnly -eq 0) {
        git merge --ff-only "origin/$Branch"
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Fast-forwarded local $Branch from origin/$Branch."
        }
        else {
            Write-Log "ERROR: Could not fast-forward local $Branch; manual sync is required."
            Start-Sleep -Seconds $RetryInterval
        }
        continue
    }

    if ($remoteOnly -gt 0 -and $localOnly -gt 0) {
        Write-Log "ERROR: Local and GitHub $Branch have diverged; manual sync is required."
        Start-Sleep -Seconds $RetryInterval
        continue
    }

    if ($localOnly -gt 0) {
        $pushOutput = git push origin $Branch 2>&1
        $pushExitCode = $LASTEXITCODE
        $pushOutput | ForEach-Object { Write-Log "$_" }

        if ($pushExitCode -eq 0) {
            Write-Log "SUCCESS: Pushed local commits to GitHub. Streamlit Cloud can now deploy the update."
        }
        else {
            Write-Log "ERROR: git push failed; pending commits will be retried."
            Start-Sleep -Seconds $RetryInterval
        }
    }

    # Wait before checking again
    Start-Sleep -Seconds $CheckInterval
}