# ============================================
# EDP Analytics - Automatic GitHub Watcher
# ============================================

# The folder containing this script
$RepoPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Git settings
$Branch = "main"

# How often to check for changes (seconds)
$CheckInterval = 5

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
    Write-Host "ERROR: This folder is not a Git repository." -ForegroundColor Red
    Write-Host ""
    Write-Host "Run this first:"
    Write-Host "git clone https://github.com/CorpEdp/edp-analytics-streamlit.git"
    exit 1
}

# Make sure we are on the correct branch
$currentBranch = git branch --show-current

if ($currentBranch -ne $Branch) {
    Write-Host "Switching to branch: $Branch" -ForegroundColor Yellow
    git checkout $Branch

    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Could not switch to $Branch." -ForegroundColor Red
        exit 1
    }
}

Write-Host "Watcher started successfully." -ForegroundColor Green
Write-Host ""

while ($true) {

    # Check Git status
    $status = git status --porcelain

    if ($status) {

        Write-Host ""
        Write-Host "--------------------------------------------" -ForegroundColor DarkGray
        Write-Host "Change detected!" -ForegroundColor Yellow
        Write-Host (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Write-Host "--------------------------------------------" -ForegroundColor DarkGray

        # Show changed files
        Write-Host ""
        Write-Host "Changed files:" -ForegroundColor Cyan
        git status --short

        # Wait briefly so multiple files saved together
        Write-Host ""
        Write-Host "Waiting for file saves to finish..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2

        # Check again
        $statusAfterWait = git status --porcelain

        if ($statusAfterWait) {

            # Add all changes
            Write-Host ""
            Write-Host "Adding changes..." -ForegroundColor Cyan
            git add .

            if ($LASTEXITCODE -ne 0) {
                Write-Host "ERROR: git add failed." -ForegroundColor Red
                Start-Sleep -Seconds $CheckInterval
                continue
            }

            # Create commit message
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $commitMessage = "Auto update EDP Analytics - $timestamp"

            Write-Host "Creating commit..." -ForegroundColor Cyan
            git commit -m "$commitMessage"

            if ($LASTEXITCODE -ne 0) {
                Write-Host "ERROR: git commit failed." -ForegroundColor Red
                Start-Sleep -Seconds $CheckInterval
                continue
            }

            # Push to GitHub
            Write-Host ""
            Write-Host "Pushing to GitHub..." -ForegroundColor Cyan

            git push origin $Branch

            if ($LASTEXITCODE -eq 0) {
                Write-Host ""
                Write-Host "SUCCESS!" -ForegroundColor Green
                Write-Host "Changes pushed to GitHub." -ForegroundColor Green
                Write-Host "Streamlit can now deploy the update." -ForegroundColor Green
            }
            else {
                Write-Host ""
                Write-Host "ERROR: GitHub push failed." -ForegroundColor Red
                Write-Host "Check your GitHub authentication and permissions." -ForegroundColor Red
            }

            Write-Host ""
        }
    }

    # Wait before checking again
    Start-Sleep -Seconds $CheckInterval
}