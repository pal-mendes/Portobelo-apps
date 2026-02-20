[CmdletBinding()]
param(
  [string]$Message = "comentários",
  [string]$AnchorFile = "Anuncios/src/Main.html"   # ficheiro "exemplo" para o permalink único
)

$ErrorActionPreference = "Stop"

function Get-LineCount([string]$file) {
  # Conta linhas de forma rápida
  (Get-Content -LiteralPath $file -ReadCount 0).Count
}

function Get-TrackedJsHtmlFiles {
  git ls-files |
    Where-Object { $_ -match '\.(js|html)$' -and $_ -notmatch '(^|[\\/])\.' } |
    Sort-Object
}

function Show-RepoLineCounts {
  Write-Host "`n=== Lines per tracked .js / .html files (working tree) ==="
  Get-TrackedJsHtmlFiles | ForEach-Object {
    $file = $_
    $count = Get-LineCount $file
    "{0,6}  {1}" -f $count, $file
  }
}

function Show-ChatGPTBlock {
  $hash = (git rev-parse HEAD).Trim()

  Write-Host "`n===== TEXT TO COPY INTO CHATGPT =====`n"
  Write-Host "Não responder sem consultar o código no <HASH> do permalink seguinte, e via raw.githubusercontent.com (source integral) - Não usar open()/render:"
  Write-Host ("https://github.com/pal-mendes/Portobelo-apps/blob/{0}/{1}?raw=1" -f $hash, $AnchorFile)
  Write-Host "O comando ""git show"", aplicado a cada um dos ficheiros <FILE> do mesmo <HASH>, retorna os números de linhas seguintes:"
  Write-Host 'git show <HASH>:<FILE> | find /c /v ""'
  Get-TrackedJsHtmlFiles | ForEach-Object {
    $file = $_
    $count = Get-LineCount $file
    "{0,6}  {1}" -f $count, $file
  }
  Write-Host "`n===== END =====`n"
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

  git fetch origin main
  git push --force-with-lease origin main
}

Show-RepoLineCounts
Show-ChatGPTBlock