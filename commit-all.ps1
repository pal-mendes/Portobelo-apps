[CmdletBinding()]
param(
  [string]$Message = "comentários"
)

$ErrorActionPreference = "Stop"

function Show-RepoLineCounts {
  Write-Host "`n=== Lines per tracked file (working tree) ==="
  git ls-files | ForEach-Object {
    $file = $_
    $count = (Get-Content -LiteralPath $file -ReadCount 0).Count
    "{0,6}  {1}" -f $count, $file
  }
}

git add -A

Write-Host "`n=== Staged changes (added / removed) ==="
git diff --cached --numstat

# Commit only if there is something staged
git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
  Write-Host "`nNothing to commit."
} else {
  git commit -m $Message
  Write-Host "`n=== Last commit (name-status) ==="
  git show --name-status --oneline -1
  git push
}

Show-RepoLineCounts
