[CmdletBinding()]
param(
  [string]$Message = "comentários"
)

$ErrorActionPreference = "Stop"

function Show-RepoLineCounts {
  Write-Host "`n=== Lines per tracked .js / .html files (working tree) ==="
  git ls-files |
    Where-Object { $_ -match '\.(js|html)$' -and $_ -notmatch '(^|[\\/])\.' } |
    ForEach-Object {
      $file = $_
      $count = (Get-Content -LiteralPath $file -ReadCount 0).Count
      "{0,6}  {1}" -f $count, $file
    }
}

# Stage everything
git add -A

# Show staged change stats
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

  # Ensure remote refs are up to date for the lease check
  git fetch origin main

  # Force push (safe variant)
  git push --force-with-lease origin main
  # git push --force origin main

}

Show-RepoLineCounts
