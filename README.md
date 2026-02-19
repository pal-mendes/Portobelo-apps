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
git add -A && git commit -am "comentários" && git show --name-status --oneline -1 && git push


Acesso online para o ChatGPT não pode ser direto, que seria assim:
https://github.com/pal-mendes/Portobelo-apps/tree/main/AuthCoreLib/src
https://github.com/pal-mendes/Portobelo-apps/blob/main/AuthCoreLib/src/Login.html


é preciso usar o Permalink:
O que é <SHA> no URL do GitHub?
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

Para ficheiros html, é preciso enviar ?raw=1 no fim do URL, e dizer ao ChatGPT que leia o source RAW para que o ChatGPT não faça o seu rendering. Ele diz para usar plain=1

https://github.com/pal-mendes/Portobelo-apps/tree/3378b69b4b1c5f2c3e559303fce4507fb7ee73cc/Associados/src?raw=1
https://github.com/pal-mendes/Portobelo-apps/tree/3378b69b4b1c5f2c3e559303fce4507fb7ee73cc/Associados/src?plain=1


Memorizar no ChatGPT que não deve usar web.run open() para consultar ficheiros HTML no Github usando os permalinks, pois reduz ficheiros de centenas de linhas a apenas 10 linhas, incluíndo “Proteção de dados… Guardar”. É preciso fazer o download direto do RAW (sem extração) de raw.githubusercontent.com, que devolve o HTML completo.


# verificar o SHA no Windows Command Prompt:
git show 00aa703f2fa6b9cec3ebc4a7aa73e3c878fe9ccc:AuthCoreLib/src/RGPD.html | find /c /v ""

c221a1258c07c427b7d0900238c63b97f95727fd


Ao analisar a consola do browser, as mensagens relevantes incluem quase todas "exec:33", MAS NEM TODAS:

Navigated to https://www.titulares-portobelo.pt/associados
...
exec:33 [LOGIN] [LOGIN] goWithTicket → https://script.google.com/a/titulares-portobelo.pt/macros/s/AKfycbznY5OWGf0uFbO7AFvYIzA-g_9Y0_r5pWBbu9i_OaSikRYKU5GLRacqDh64ZXKeSmge/exec?ticket=eyJlbWFpbCI6InBhbC5tZW5kZXMyM0BnbWFpbC5jb20iLCJleHAiOjE3NzI1NTEzNTk2MTQsInYiOjIsImlhdCI6MTc3MTM0MTc1OTYxNCwibmFtZSI6IlBlZHJvIE1lbmRlcyIsInBpY3R1cmUiOiJodHRwczovL2xoMy5nb29nbGV1c2VyY29udGVudC5jb20vYS9BQ2c4b2NMbU5FT3hCRklMbV9jbXFsZWRwa1pxa3pnbFQzZEdrMHROU2kzeGRmUmUtOXMtbnowWEtnPXM5Ni1jIn0%3D.tu1LgcxLkqKvJgYSBpmMdE5SeeonIZ3NXHMOC-l8jus%3D&ts=1771438281401
...
userCodeAppPanel?createOAuthDialog=true:34 [RGPD] RGPD page ready; CANON=https://script.google.com/a/titulares-portobelo.pt/macros/s/AKfycbznY5OWGf0uFbO7AFvYIzA-g_9Y0_r5pWBbu9i_OaSikRYKU5GLRacqDh64ZXKeSmge/exec; have TICKET=true
...
userCodeAppPanel?createOAuthDialog=true:34 [RGPD] [RGPD] replace → https://script.google.com/a/titulares-portobelo.pt/macros/s/AKfycbznY5OWGf0uFbO7AFvYIzA-g_9Y0_r5pWBbu9i_OaSikRYKU5GLRacqDh64ZXKeSmge/exec?action=postrgpd&ticket=eyJlbWFpbCI6InBhbC5tZW5kZXMyM0BnbWFpbC5jb20iLCJleHAiOjE3NzI1NTEzNTk2MTQsInYiOjIsImlhdCI6MTc3MTM0MTc1OTYxNCwibmFtZSI6IlBlZHJvIE1lbmRlcyIsInBpY3R1cmUiOiJodHRwczovL2xoMy5nb29nbGV1c2VyY29udGVudC5jb20vYS9BQ2c4b2NMbU5FT3hCRklMbV9jbXFsZWRwa1pxa3pnbFQzZEdrMHROU2kzeGRmUmUtOXMtbnowWEtnPXM5Ni1jIn0%3D.tu1LgcxLkqKvJgYSBpmMdE5SeeonIZ3NXHMOC-l8jus%3D&rows=137&from=rgpd-save&ts=1771438293545
...
VM470 exec:33 [LOGIN] [LOGIN] goWithTicket → https://script.google.com/a/titulares-portobelo.pt/macros/s/AKfycbznY5OWGf0uFbO7AFvYIzA-g_9Y0_r5pWBbu9i_OaSikRYKU5GLRacqDh64ZXKeSmge/exec?ticket=eyJlbWFpbCI6InBhbC5tZW5kZXMyM0BnbWFpbC5jb20iLCJleHAiOjE3NzI1NTEzNTk2MTQsInYiOjIsImlhdCI6MTc3MTM0MTc1OTYxNCwibmFtZSI6IlBlZHJvIE1lbmRlcyIsInBpY3R1cmUiOiJodHRwczovL2xoMy5nb29nbGV1c2VyY29udGVudC5jb20vYS9BQ2c4b2NMbU5FT3hCRklMbV9jbXFsZWRwa1pxa3pnbFQzZEdrMHROU2kzeGRmUmUtOXMtbnowWEtnPXM5Ni1jIn0%3D.tu1LgcxLkqKvJgYSBpmMdE5SeeonIZ3NXHMOC-l8jus%3D&ts=1771438353675
...
exec:33 [MAIN] [render] content updated
