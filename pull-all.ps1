@('AuthCoreLib','Associados','Anuncios','Titulares') | ForEach-Object {
  Push-Location $_
  try { clasp pull } finally { Pop-Location }
}
