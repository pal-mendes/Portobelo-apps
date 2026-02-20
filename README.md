“Ativar a Apps Script API” em https://script.google.com/home/usersettings
 → Enable Google Apps Script API.
Se o clasp clone já funcionou, isso já está ON e não tens de mexer mais.


************************************
Utilizar o Windows Command Prompt
************************************

É preciso fazer login (admin@titulares-portobelo.pt), que expira ao fim de algumas horas:
clasp logout
clasp help login
clasp login

Atualizar o repositório local a partir da Google (fonte):
cd \dev\Portobelo-apps
C:\dev\Portobelo-apps>powershell .\pull-all.ps1

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

Para adicionar novos ficheiros ao git:
git add -A

Quando se quiser mandar para o Git Hub, para o ChatGPT poder consultar e sugerir melhorias:
git add -A && git commit -m "comentários" && git show --name-status --oneline -1 && git push

O melhor é executar o script:
C:\dev\Portobelo-apps>powershell .\commit-all.ps1

# verificar o SHA no Windows Command Prompt: diz quantas linhas tem o ficheiro
git show 00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc:AuthCoreLib/src/RGPD.html | find /c /v ""
284


Acesso ao GitHub (não para o ChatGPT)
https://github.com/pal-mendes/Portobelo-apps/tree/main/AuthCoreLib/src
https://github.com/pal-mendes/Portobelo-apps/blob/main/AuthCoreLib/src/Login.html





O que é <SHA> ou <HASH> no URL do GitHub?
<SHA> é o hash do commit (ex.: 2f552bd…). Um link com SHA é um permalink para essa versão exacta do ficheiro: as linhas #L42-L92 nunca se “mexem” porque o conteúdo não muda.
Para obteres o permalink na UI do GitHub:
1 - navegar para um dos projetos (URL https://github.com/pal-mendes/Portobelo-apps/tree/main/Anuncios) e escolher "Copy permalink" (Ctrl Shift ,) do botão "..." no canto superior direito.
2 - Abre o ficheiro → clica no botão ⋯ (menu “More”) → Copy permalink; ou carrega no número da linha e depois na tecla Y e copia o URL que mudou para …/blob/<SHA>/… se for para um ficheiro ou …/tree/<SHA>/… se for para o diretório.
=> resultado = https://github.com/pal-mendes/Portobelo-apps/blob/2f552bd1206775522ea5f136801a2d4be6bddb17/AuthCoreLib/src/RGPD.html
=> resultado = https://github.com/pal-mendes/Portobelo-apps/blob/2f552bd1206775522ea5f136801a2d4be6bddb17/AuthCoreLib/src/Login.html
=> resultado = https://github.com/pal-mendes/Portobelo-apps/blob/2f552bd1206775522ea5f136801a2d4be6bddb17/Associados/src/Associados.js
=> resultado = https://github.com/pal-mendes/Portobelo-apps/tree/3378b69b4b1c5f2c3e559303fce4507fb7ee73cc/AuthCoreLib/src

Usar main no URL (como fizeste) é totalmente válido e público — só que é móvel: ao fazeres novos commits, as linhas podem mudar.

Permalink por commit (não muda com futuros commits) com âncoras de linhas, ex.: .../blob/<SHA>/AuthCoreLib/src/RGPD.html#L12-L60
Raw (se quiseres que eu veja só o texto):
https://raw.githubusercontent.com/pal-mendes/Portobelo-apps/<branch ou SHA>/AuthCoreLib/src/RGPD.html


########## Pedir ao ChatGPT para verificar o código no  - usar em cada mensagem:
Não responder sem consultar o código no <HASH> do permalink seguinte, e via raw.githubusercontent.com (source integral) - Não usar open()/render:
https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/Anuncios/src/Main.html?raw=1
O comando "git show", aplicado a cada um dos ficheiros <FILE> do mesmo <HASH>, retorna os números de linhas seguintes:
git show <HASH>:<FILE> | find /c /v ""
   774  Anuncios/src/Anuncios.js
   260  Anuncios/src/App.html
   238  Anuncios/src/AuthCore.gs.js
   238  Anuncios/src/AuthCore.js
    29  Anuncios/src/Form.html
   269  Anuncios/src/Login.html
   182  Anuncios/src/Main.html
   170  Anuncios/src/Styles.html
  1035  Associados/src/Associados.js
    51  Associados/src/Logger_js.html
   914  Associados/src/Main.html
   907  AuthCoreLib/src/AuthCore.js
    51  AuthCoreLib/src/Logger_js.html
   384  AuthCoreLib/src/Login.html
   286  AuthCoreLib/src/RGPD.html



Não responder sem consultar o código nos permalinks seguintes, e via raw.githubusercontent.com (source integral) - Não usar open()/render.
O comando "git show", aplicado a cada um dos ficheiros <FILE> do mesmo <HASH>, retorna os números de linhas seguintes:
git show <HASH>:<FILE> | find /c /v ""
   774  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/Anuncios/src/Anuncios.js?raw=1
   260  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/Anuncios/src/App.html?raw=1
   238  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/Anuncios/src/AuthCore.js?raw=1
    29  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/Anuncios/src/Form.html?raw=1
   269  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/Anuncios/src/Login.html?raw=1
   182  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/Anuncios/src/Main.html?raw=1
   170  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/Anuncios/src/Styles.html?raw=1
  1035  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/Associados/src/Associados.js?raw=1
    51  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/Associados/src/Logger_js.html?raw=1
   914  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/Associados/src/Main.html?raw=1
   907  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/AuthCoreLib/src/AuthCore.js?raw=1
    51  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/AuthCoreLib/src/Logger_js.html?raw=1
   384  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/AuthCoreLib/src/Login.html?raw=1
   286  https://github.com/pal-mendes/Portobelo-apps/blob/00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc/AuthCoreLib/src/RGPD.html?raw=1
