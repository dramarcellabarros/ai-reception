# 📊 OPÇÃO 3: Google Data Studio - Setup Completo

## ✅ O que você vai criar

Um relatório visual e interativo que:
- ✅ Conecta direto na planilha Google Sheets (dados sempre atualizados)
- ✅ Mostra 4 cards com estatísticas
- ✅ Tabela interativa com todos os leads
- ✅ Gráfico de pizza (distribuição por status)
- ✅ Gráfico de barras (leads por origem)
- ✅ Filtros interativos (por status, origem, data)
- ✅ Cores da marca Dra. Marcella (Dourado, Café)
- ✅ Responsivo (funciona em celular)

---

## 🚀 PASSO-A-PASSO (30 minutos)

### PASSO 1: Acessar Google Data Studio

1. Abra [datastudio.google.com](https://datastudio.google.com)
2. Faça login com sua conta Google
3. Clique em **"Criar"** no canto inferior esquerdo
4. Selecione **"Relatório em branco"**

---

### PASSO 2: Conectar a Planilha Google Sheets

1. No painel esquerdo, clique em **"Adicionar dados"**
2. Procure por **"Google Sheets"**
3. Clique em **"Google Sheets"**
4. Selecione sua planilha:
   - Nome: `[Procure pela planilha que contém os leads]`
   - Abas: Selecione `Sheet1` (ou o nome da aba com os dados)
5. Clique em **"Conectar"**

---

### PASSO 3: Adicionar Título e Descrição

1. Clique no **"Titulo"** padrão (texto no topo)
2. Mude para: **"CRM — Consultório Dra. Marcella Barros"**
3. Clique em outro lugar para confirmar
4. Adicione abaixo: **"Gestão de Leads e Agendamentos"** (menor, cinza)

---

### PASSO 4: Adicionar Score Cards (Estatísticas)

**Criar 4 cards com totais:**

#### Card 1: Total de Leads
1. Clique em **"Inserir"** → **"Score card"**
2. Arraste para o topo esquerdo
3. Clique no card para configurar
4. **Métrica**: `COUNTA(Telefone)` ou selecione "Telefone" como dimensão
5. **Texto**: "Leads Totais"
6. **Cor de fundo**: #E8DDD0 (Areia)

#### Card 2: Agendamentos Confirmados
1. Clique em **"Inserir"** → **"Score card"**
2. **Métrica**: Contar registros onde "Agendamento Confirmado (JSON)" não está vazio
   - Ou: `COUNTIF('Agendamento Confirmado (JSON)', '<>')`
3. **Texto**: "Agendamentos Confirmados"
4. **Cor**: #d4edda (Verde)

#### Card 3: Pendentes de Confirmação
1. Similarmente, contar onde "Estado" = "BOOKED" ou "READY_TO_BOOK"
2. **Texto**: "Pendentes de Confirmação"
3. **Cor**: #fff3cd (Amarelo)

#### Card 4: Não Compareceram
1. Contar onde "Lead Status" contém "não compareceu"
2. **Texto**: "Não Compareceram"
3. **Cor**: #f8d7da (Vermelho)

---

### PASSO 5: Adicionar Tabela de Dados

1. Clique em **"Inserir"** → **"Tabela"**
2. Arraste para ocupar metade inferior da página
3. **Colunas a adicionar** (nesta ordem):
   - Nome
   - Telefone
   - Origem
   - Estado
   - Lead Status
   - Horário Escolhido (JSON)
   - Agendamento Confirmado (JSON)

4. Configurar formatação:
   - **Cabeçalho**: Cor de fundo #6B4423 (Café), texto branco
   - **Linhas**: Alternar cores (branco e #E8DDD0)

5. **Adicionar filtro** (opcional):
   - Clique em **"Adicionar filtro"**
   - Selecione "Status" ou "Origem"
   - Isso permite filtrar os dados dinamicamente

---

### PASSO 6: Adicionar Gráficos

#### Gráfico 1: Distribuição por Status (Pizza)

1. Clique em **"Inserir"** → **"Gráfico de pizza"**
2. Arraste para o lado direito
3. **Dimensão**: "Lead Status" (ou "Estado")
4. **Métrica**: Contar registros
5. **Cores**:
   - Confirmado: #d4edda
   - Não compareceu: #f8d7da
   - Pendente: #fff3cd
   - Padrão: #E8DDD0

#### Gráfico 2: Leads por Origem (Barra)

1. Clique em **"Inserir"** → **"Gráfico de colunas"**
2. **Dimensão**: "Origem"
3. **Métrica**: Contar registros
4. **Cor**: #C9A961 (Dourado)

---

### PASSO 7: Formatação Geral

1. **Plano de fundo da página**:
   - Clique em **"Arquivo"** → **"Configurações de página"**
   - **Cor de fundo**: #f5f2ed (Bege claro)

2. **Fontes**:
   - Selecione um elemento
   - **Tipografia**: Arial ou Roboto
   - **Tamanho**: 14px para corpo, 20px para títulos

3. **Paleta de cores - Aplicar globalmente**:
   - Clique em **"Arquivo"** → **"Tema e layout"**
   - **Cor primária**: #C9A961 (Dourado)
   - **Cor secundária**: #6B4423 (Café)

---

### PASSO 8: Adicionar Filtros Interativos (Opcional)

1. Clique em **"Inserir"** → **"Controle de filtro"**
2. Configure para filtrar por:
   - Status
   - Origem
   - Data
3. Vincule ao painel de controle clicando no gráfico/tabela

---

### PASSO 9: Salvar e Compartilhar

1. Clique em **"Arquivo"** → **"Salvar"**
2. Nome: **"CRM - Dra. Marcella Barros"**
3. Para compartilhar:
   - Clique em **"Compartilhar"** (canto superior direito)
   - Adicione e-mails ou gere link público

---

## 🎨 Paleta de Cores - Copie e Cole

```css
/* Cores da Marca */
--primary: #C9A961;         /* Dourado */
--secondary: #6B4423;       /* Café escuro */
--tertiary: #8B5A3C;        /* Marrom */
--neutral-light: #E8DDD0;   /* Areia/Bege */
--status-confirmed: #d4edda;/* Verde (Confirmado) */
--status-pending: #fff3cd;  /* Amarelo (Pendente) */
--status-noshow: #f8d7da;   /* Vermelho (Não Compareceu) */
--background: #f5f2ed;      /* Fundo */
```

---

## 📱 Configurações de Responsividade

**No Data Studio:**
1. Clique em **"Arquivo"** → **"Configurações de página"**
2. **Tamanho**: Selecione "Automático" (responsivo)
3. Isso faz o relatório se adaptar automaticamente a celular/tablet

---

## 🔄 Manutenção e Atualização

- ✅ Os dados são atualizados **automaticamente** cada vez que você abre o relatório
- ✅ Se adicionar novas linhas à planilha, aparecem aqui em tempo real
- ✅ Não precisa fazer nada — a conexão é direta

---

## 📊 O que o relatório final vai parecer

```
┌────────────────────────────────────────┐
│  CRM — Consultório Dra. Marcella Barros │
│  Gestão de Leads e Agendamentos        │
└────────────────────────────────────────┘

┌──────────────┬──────────────┬──────────────┬──────────────┐
│     6        │      3       │      1       │      1       │
│ Leads Totais │ Confirmados  │  Pendentes   │ Não Comp.    │
└──────────────┴──────────────┴──────────────┴──────────────┘

┌──────────────────────────────┬──────────────────────────────┐
│  Tabela de Leads (interativa)│  Distribuição por Status      │
│  - Nome, Telefone, Origem    │  [Gráfico Pizza]              │
│  - Estado, Status            │  - Confirmado: 50%            │
│  - Agendamento              │  - Não comp.: 16%             │
│                             │  - Pendente: 34%              │
│  [Filtros: Status, Origem]   │                               │
└──────────────────────────────┴──────────────────────────────┘

┌──────────────────────────────────────────────────────────┐
│  Leads por Origem (Gráfico de Barras)                   │
│  - Instagram: 4                                          │
│  - WhatsApp: 1                                           │
│  - Site: 1                                               │
└──────────────────────────────────────────────────────────┘
```

---

## 🆘 Troubleshooting

| Problema | Solução |
|---|---|
| **"Dados não aparecem"** | Verifique se conectou a aba correta da planilha (Sheet1) |
| **"Filtro não funciona"** | Clique no gráfico/tabela → "Dados" → Verifique dimensão/métrica |
| **"Cores não aparecem"** | Data Studio pode sobrescrever — use tema global "Arquivo" → "Tema e layout" |
| **"Relatório lento"** | Remova gráficos desnecessários ou filtre por período |

---

## ✅ Próximos Passos

1. **Criar o Data Studio** seguindo este guia (30 min)
2. **Combinar com Opção 1** (formatação Sheets) — tira 5 min
3. **Usar Dashboard HTML** (Opção 2) como alternativa mobile-first

**Recomendação final:** Use as **3 opções em conjunto**:
- 📱 **Opção 2 (Dashboard HTML)**: Para celular, no app
- 📊 **Opção 3 (Data Studio)**: Para desktop, relatórios
- 🎨 **Opção 1 (Formatação Sheets)**: Como backup/fonte de verdade

---

## 📞 Suporte

Se tiver dúvidas:
- Google Data Studio Help: [support.google.com/datastudio](https://support.google.com/datastudio)
- Documentação Sheets API: [developers.google.com/sheets](https://developers.google.com/sheets)
