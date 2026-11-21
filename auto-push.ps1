# ============================================
# EDP Analytics - Automatic GitHub Watcher
# ============================================

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
$currentBranch = git branch --show-current

if ($currentBranch -ne $Branch) {
    Write-Host "Switching to branch: $Branch" -ForegroundColor Yellow
    git checkout $Branch

    if ($LASTEXITCODE -ne 0) {
        Write-Log "ERROR: Could not switch to $Branch."
        exit 1
    }
}

Write-Log "Watcher started. Repository: $RepoPath; branch: $Branch."

while ($true) {
    $status = git status --porcelain --untracked-files=normal
    if ($LASTEXITCODE -ne 0) {
        Write-Log "ERROR: git status failed; will retry."
        Start-Sleep -Seconds $RetryInterval
        continue
    }

    if ($status) {
        Write-Log "Change detected:"
        git status --short | ForEach-Object { Write-Log $_ }

        # Wait briefly so multiple files saved together
        Start-Sleep -Seconds 2

        $statusAfterWait = git status --porcelain --untracked-files=normal
        if ($LASTEXITCODE -ne 0) {
            Write-Log "ERROR: Could not recheck Git status; will retry."
            Start-Sleep -Seconds $RetryInterval
            continue
        }

        if ($statusAfterWait) {
            git add -A

            if ($LASTEXITCODE -ne 0) {
                Write-Log "ERROR: git add failed; will retry."
                Start-Sleep -Seconds $RetryInterval
                continue
            }

            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $commitMessage = "Auto update EDP Analytics - $timestamp"
            git commit -m "$commitMessage"

            if ($LASTEXITCODE -ne 0) {
                Write-Log "ERROR: git commit failed; staged changes will be retried."
            }
            else {
                Write-Log "Committed local changes."
            }
        }
    }

    # Fetch and push independently of worktree changes so a failed push is retried.
    git fetch origin $Branch
    if ($LASTEXITCODE -ne 0) {
        Write-Log "ERROR: git fetch failed; will retry."
        Start-Sleep -Seconds $RetryInterval
        continue
    }

    $divergence = git rev-list --left-right --count "HEAD...origin/$Branch"
    if ($LASTEXITCODE -ne 0) {
        Write-Log "ERROR: Could not compare local and remote commits; will retry."
        Start-Sleep -Seconds $RetryInterval
        continue
    }

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