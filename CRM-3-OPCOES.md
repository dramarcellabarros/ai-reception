# CRM — 3 Opções de Visualização

## ✅ OPÇÃO 1: Formatação Google Sheets (Mais Rápido)

**O que fazer:**
1. Abra a planilha real no Google Sheets
2. Selecione todas as linhas de dados
3. Menu → Formato → Formatação condicional
4. Aplique cores baseadas no status:
   - **Cor Dourada (#C9A961)** → Agendamentos Confirmados
   - **Cor Café (#6B4423)** → Pendentes de Confirmação
   - **Cor Vermelha (#d9534f)** → Não Compareceram
   - **Cor Bege (#E8DDD0)** → Leads sem agendamento

**Passos específicos:**
```
Menu → Formato → Formatação condicional
Regra: Se 'Agendamento Confirmado (JSON)' não está vazio → Cor Dourada
Regra: Se 'Estado' = 'BOOKED' → Cor Café
Regra: Se 'Status' contém 'não compareceu' → Cor Vermelha
```

**Vantagens:**
- ✅ Rápido de fazer (5 minutos)
- ✅ Funciona direto na Google Sheets
- ✅ Acesso via celular/desktop

**Desvantagens:**
- ❌ Limitado em customização
- ❌ Ainda é uma planilha (não é "CRM")

---

## ✅ OPÇÃO 2: Dashboard Web HTML/CSS (Já Criado!)

**Arquivo:** `crm-dashboard.html` na raiz do projeto

**Como usar:**
1. Abra o arquivo direto no navegador (desktop ou celular)
2. Busque por nome, telefone ou origem
3. Veja estatísticas em tempo real
4. Botões de ação (View, Edit, etc.) prontos pra expandir

**Características:**
- 🎨 Design elegante com identidade visual da clínica
- 📊 Dashboard com estatísticas (Total, Confirmados, Pendentes, Não Compareceram)
- 🔍 Busca em tempo real
- 📱 Responsivo (funciona em celular)
- 💾 Dados simulados (em produção, conectar com API da planilha)

**Vantagens:**
- ✅ Visual profissional e CRM real
- ✅ Responsivo (celular, tablet, desktop)
- ✅ Rápido e leve (sem dependências)
- ✅ Customizável (mudar cores, adicionar campos fácil)

**Para conectar com dados reais:**
```javascript
// Substituir mockData com chamada à API
const response = await fetch('https://sheets.googleapis.com/v4/spreadsheets/{ID}/values/Sheet1?key={API_KEY}');
const data = await response.json();
```

**Desvantagens:**
- ⚠️ Dados ainda são simulados (precisa conectar API)
- ⚠️ Precisa de servidor pra alojar (ou rodar local)

---

## ✅ OPÇÃO 3: Google Data Studio (Meio-Termo)

**Como criar:**
1. Acesse [datastudio.google.com](https://datastudio.google.com)
2. "Criar" → "Relatório em branco"
3. "Adicionar dados" → selecione a planilha Google Sheets da clínica
4. Arraste para o canvas:
   - **Tabela de dados** (todos os leads com status)
   - **Score cards** (4 cards mostrando totais: Leads, Confirmados, Pendentes, Não Compareceram)
   - **Gráfico de pizza** (distribuição de status)
   - **Mapa de calor** (agendamentos por dia)

5. Formatação:
   - Cores primárias: Dourado (#C9A961) e Café (#6B4423)
   - Tipografia: Arial ou Roboto
   - Paleta neutra: Branco, Bege (#E8DDD0), Preto

**Vantagens:**
- ✅ Conecta direto na planilha (dados sempre atualizados)
- ✅ Sem código
- ✅ Pronto em 15 minutos
- ✅ Visual mais bonito que planilha
- ✅ Filtros e interatividade

**Desvantagens:**
- ⚠️ Menos controle visual que HTML customizado
- ⚠️ Precisa de conta Google com acesso pago (relatórios avançados)
- ⚠️ Mais lento em atualização que HTML puro

---

## 🎯 Recomendação

| Situação | Recomendação |
|----------|---|
| Quer agora, rápido | **Opção 1** (Formatação Sheets) |
| Quer profissional e responsivo | **Opção 2** (Dashboard HTML) |
| Quer fácil + bonito + automático | **Opção 3** (Google Data Studio) |
| **Melhor custo-benefício** | **Opção 2 + Opção 1** (HTML + formatação) |

---

## 🚀 Próximos Passos

1. **Teste o Dashboard HTML** (já está pronto em `crm-dashboard.html`)
2. **Escolha qual usar** (ou combine 1 + 2)
3. **Se escolher Opção 2**: Conectar com API real da planilha (trocar `mockData`)
4. **Se escolher Opção 3**: Posso criar o Data Studio pra você com 1 clique

---

## 📝 Paleta de Cores da Marca (Confirmada)

```css
--primary: #C9A961;      /* Dourado */
--secondary: #6B4423;    /* Café */
--tertiary: #8B5A3C;     /* Marrom */
--neutral-light: #E8DDD0;/* Areia/Bege */
--neutral-white: #FFFFFF;/* Branco */
--neutral-dark: #000000; /* Preto */
```

**Identidade:** Luxuosa, minimalista, elegante — perfeita pra clínica estética.
