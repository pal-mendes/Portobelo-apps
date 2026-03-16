@('AuthCoreLib','Associados','Anuncios') | ForEach-Object {
  Push-Location $_
  try { clasp pull } finally { Pop-Location }
}
