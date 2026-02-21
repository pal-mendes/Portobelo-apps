[CmdletBinding()]
param(
  [string]$Message = "comentários"
  # [string]$AnchorFile = "AuthCoreLib/src/Login.html"   # ficheiro "exemplo" para o permalink único
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


function Show-ChatGPTBlock {
  $hash = (git rev-parse HEAD).Trim()
  $repo = "pal-mendes/Portobelo-apps"
  
  $files = @(
  "Anuncios/src/Anuncios.js",
  "Anuncios/src/App.html",
  "Anuncios/src/Form.html",
  "Anuncios/src/Main.html",
  "Anuncios/src/OldAuthCore.js",
  "Anuncios/src/OldLogin.html",
  "Anuncios/src/Styles.html",
  "Associados/src/Associados.js",
  "Associados/src/Logger_js.html",
  "Associados/src/Main.html",
  "AuthCoreLib/src/AuthCore.js",
  "AuthCoreLib/src/Logger_js.html",
  "AuthCoreLib/src/Login.html",
  "AuthCoreLib/src/RGPD.html"
)

  Write-Host ""
  Write-Host "`n===== TEXT TO COPY INTO CHATGPT =====`n"
  Write-Host "Usar o conector GitHub (já ligado) para ler o código fonte integral neste repo e commit:"
  Write-Host "REPO=$repo"
  Write-Host "REF=$hash"
  Write-Host "FILES (paths no repo, a ler no REF):"
  $files | ForEach-Object { Write-Host "- $_" }
  Write-Host ""
  Write-Host "RULE: Não responder sem consultar estes ficheiros via conector GitHub no REF indicado (conteúdo integral). Não usar open()/render. Só usar RAW como fallback se o conector falhar."
  # Write-Host ("https://github.com/pal-mendes/Portobelo-apps/blob/{0}/{1}?raw=1" -f $hash, $AnchorFile)
  # Write-Host  ("RAW_BASE=https://raw.githubusercontent.com/pal-mendes/Portobelo-apps/{0}/" -f $hash)
  # Write-Host "O comando ""git show"", aplicado a cada um dos ficheiros FILE do mesmo HASH, retorna os seguintes números de linhas:"
  # Write-Host 'git show <HASH>:<FILE> | find /c /v ""'
  # Get-TrackedJsHtmlFiles | ForEach-Object {
  #   $file = $_
  #   $count = Get-LineCount $file
  #   "{0,6}  {1}" -f $count, $file
  # }
  Write-Host "(Checks locais opcionais: linhas por ficheiro via git show)"
  foreach ($f in $files) {
    $n = (git show "$hash`:$f" | find /c /v "")
    Write-Host ("{0,6}  {1}" -f $n, $f)
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

# [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# $OutputEncoding = [System.Text.Encoding]::UTF8
Show-ChatGPTBlock