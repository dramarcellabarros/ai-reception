# Setup — Criar agendamento direto pelo Dashboard

Isso conecta o Dashboard a um pequeno script que roda com a sua própria
permissão de dono da planilha/agenda, para poder **criar** agendamentos
(não só ler). Leva uns 5 minutos.

## Passo 1 — Abrir o editor de script

1. Abra a planilha real: https://docs.google.com/spreadsheets/d/1epXj_UweUIR8sq29ZWjZ_lGiFsHzyp4ULz9A8DhEKps/edit
2. Menu **Extensões** → **Apps Script**
3. Vai abrir um editor com um arquivo `Código.gs` vazio (ou com um `function myFunction(){}` padrão)

## Passo 2 — Colar o código

1. Apague todo o conteúdo padrão do editor
2. Cole o conteúdo inteiro do arquivo [`Code.gs`](./Code.gs) deste projeto
3. Clique no ícone de salvar (💾) ou `Ctrl+S`
4. Dê um nome ao projeto quando pedir, ex: "CRM Backend"

## Passo 3 — Publicar como aplicativo web

1. Clique no botão azul **"Implantar"** (canto superior direito) → **"Nova implantação"**
2. Clique no ícone de engrenagem ⚙️ ao lado de "Selecionar tipo" → escolha **"Aplicativo da web"**
3. Preencha:
   - **Descrição**: `CRM Backend v1` (o que quiser)
   - **Executar como**: **Eu** (sua conta — importante, é isso que dá a permissão)
   - **Quem tem acesso**: **Qualquer pessoa**
4. Clique **"Implantar"**
5. Na primeira vez, o Google vai pedir para você **autorizar o script**:
   - Clique em "Autorizar acesso"
   - Escolha sua conta Google
   - Vai aparecer um aviso "O Google não verificou este app" — isso é normal (é o seu próprio script). Clique em **"Configurações avançadas"** → **"Acessar CRM Backend (não seguro)"**
   - Confirme as permissões (acesso à planilha e à agenda)
6. Copie a **URL do aplicativo da web** que aparece (termina em `/exec`)

## Passo 4 — Colar a URL no Dashboard

1. Abra o Dashboard, clique em **⚙️ Config**
2. Cole a URL no campo **"URL do Apps Script (para criar agendamentos)"**
3. Salvar

Pronto — o botão **"+ Novo Agendamento"** no Dashboard já vai funcionar.

## Se precisar atualizar o script depois

Sempre que editar `Code.gs` (ou receber uma versão nova deste projeto):

1. Cole o novo código no editor do Apps Script
2. **Implantar** → **Gerenciar implantações** → clique no ícone de lápis ✏️ na implantação existente
3. Em "Versão", escolha **"Nova versão"** → **"Implantar"**

Isso atualiza o mesmo endpoint (mesma URL), sem precisar reconfigurar o Dashboard.

## Segurança

A URL gerada é como uma senha longa e aleatória — qualquer pessoa que a
tiver pode criar agendamentos através dela (o script roda com sua
permissão de dono). Não publique essa URL em lugar público (ex: não a
inclua no código do GitHub Pages, que é público — ela deve ficar salva
apenas no Config do navegador de cada dispositivo, exatamente como a API
Key e o ID da planilha já funcionam hoje).
