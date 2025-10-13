“Ativar a Apps Script API” em https://script.google.com/home/usersettings
 → Enable Google Apps Script API.
Se o clasp clone já funcionou, isso já está ON e não tens de mexer mais.

Atualizar o repositório local a partir do Command Prompt:

cd Portobelo-apps/
C:\Users\Pal-m\My Drive\Imóveis\Portobelo\Portobelo-apps>powershell .\pull-all.ps1

Estrutura:
Portobelo-apps/
  AuthCoreLib/
    .clasp.json
    src/
      appsscript.json
      *.gs
      html/
  Associados/
    .clasp.json
    src/
      appsscript.json
      *.gs
      html/
  Anuncios/
    .clasp.json
    src/
      appsscript.json
      *.gs
      html/
  .gitignore
  README.md  (opcional, explica a estrutura)
