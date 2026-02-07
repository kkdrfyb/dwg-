param(
    [string]$Message = "update"
)

Write-Host "Git status:"
git status
if ($LASTEXITCODE -ne 0) {
    Write-Error "git status failed. Is this a git repo?"
    exit 1
}

Write-Host "Adding all changes..."
git add .
if ($LASTEXITCODE -ne 0) {
    Write-Error "git add failed."
    exit 1
}

Write-Host "Committing..."
git commit -m $Message
if ($LASTEXITCODE -ne 0) {
    Write-Error "git commit failed. (Maybe nothing to commit)"
    exit 1
}

Write-Host "Pushing..."
git push
if ($LASTEXITCODE -ne 0) {
    Write-Error "git push failed."
    exit 1
}

Write-Host "Done."
