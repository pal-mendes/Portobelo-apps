“Ativar a Apps Script API” em https://script.google.com/home/usersettings
 → Enable Google Apps Script API.
Se o clasp clone já funcionou, isso já está ON e não tens de mexer mais.

É preciso fazer login, que expira ao fim de algumas horas:
clasp logout
clasp help login
clasp login


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


Quando se quiser mandar para o Git Hub, para o ChatGPT ver:
git commit -m "comentários"
git commit -a

Acesso online para o ChatGPT:
https://github.com/pal-mendes/Portobelo-apps/tree/main/AuthCoreLib/src
https://github.com/pal-mendes/Portobelo-apps/blob/main/AuthCoreLib/src/Login.html


Permalink:
O que é <SHA> no URL do GitHub?
<SHA> é o hash do commit (ex.: 2f552bd…). Um link com SHA é um permalink para essa versão exacta do ficheiro: as linhas #L42-L92 nunca se “mexem” porque o conteúdo não muda.
Para obteres o permalink na UI do GitHub:
Abre o ficheiro → clica no botão ⋯ (menu “More”) → Copy permalink;
ou carrega no número da linha e depois na tecla Y e copia o URL que mudou para …/blob/<SHA>/….
=> resultado = https://github.com/pal-mendes/Portobelo-apps/blob/2f552bd1206775522ea5f136801a2d4be6bddb17/AuthCoreLib/src/RGPD.html
=> resultado = https://github.com/pal-mendes/Portobelo-apps/blob/2f552bd1206775522ea5f136801a2d4be6bddb17/AuthCoreLib/src/Login.html

Usar main no URL (como fizeste) é totalmente válido e público — só que é móvel: ao fazeres novos commits, as linhas podem mudar.

Permalink por commit (não muda com futuros commits) com âncoras de linhas, ex.: .../blob/<SHA>/AuthCoreLib/src/RGPD.html#L12-L60
Raw (se quiseres que eu veja só o texto):
https://raw.githubusercontent.com/pal-mendes/Portobelo-apps/<branch ou SHA>/AuthCoreLib/src/RGPD.html
