# Força o terminal a interpretar caracteres UTF-8 corretamente
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
@('AuthCoreLib','Associados','Anuncios','Titulares','Anuncios-sheet') | ForEach-Object {
  $project = $_
  if (Test-Path $project) {
    Push-Location $project
    try {
      # Captura o output e substitui a linha final injetando o nome do projeto
      (clasp pull) -replace 'Pulled (\d+) files\.', "Pulled `$1 files from $project."
    } finally { Pop-Location }
  }
}