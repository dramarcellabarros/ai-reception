/**
 * OPÇÃO 1: Script para formatar Google Sheets com cores da marca Dra. Marcella
 *
 * Como usar:
 * 1. Abra a planilha Google Sheets real
 * 2. Menu → Extensões → Apps Script
 * 3. Cole este código inteiro
 * 4. Execute a função: formatLeadsSheet()
 * 5. Pronto! Planilha formatada
 *
 * Paleta de cores:
 * - Dourado: #C9A961
 * - Café: #6B4423
 * - Marrom: #8B5A3C
 * - Areia: #E8DDD0
 * - Verde (confirmado): #d4edda
 * - Vermelho (não compareceu): #f8d7da
 * - Amarelo (pendente): #fff3cd
 */

function formatLeadsSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();

  // Cores
  const colors = {
    header: '#6B4423',        // Café escuro
    dourado: '#C9A961',       // Dourado
    confirmed: '#d4edda',     // Verde claro
    noshow: '#f8d7da',        // Vermelho claro
    pending: '#fff3cd',       // Amarelo claro
    areia: '#E8DDD0'          // Bege/Areia
  };

  // 1. Formatar header (primeira linha)
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setBackground(colors.header);
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setHeight(30);

  // 2. Formatar dados (aplicar cores por status)
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const backgrounds = [];
  const fontColors = [];

  // Encontrar índices das colunas
  const headers = values[0];
  const statusCol = headers.indexOf('Status');
  const confirmedCol = headers.indexOf('Agendamento Confirmado (JSON)');
  const leadStatusCol = headers.indexOf('Lead Status');

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    let bgColor = '#FFFFFF';
    let fontColor = '#333333';

    // Determinar cor baseada no status
    if (confirmedCol >= 0 && row[confirmedCol]) {
      // Tem agendamento confirmado
      bgColor = colors.confirmed;
      fontColor = '#155724';
    } else if (leadStatusCol >= 0 && row[leadStatusCol]?.includes('não compareceu')) {
      // Não compareceu
      bgColor = colors.noshow;
      fontColor = '#721c24';
    } else if (row[statusCol]?.includes('BOOKED') || row[statusCol]?.includes('READY_TO_BOOK')) {
      // Pendente de confirmação
      bgColor = colors.pending;
      fontColor = '#856404';
    } else {
      // Default
      bgColor = colors.areia;
      fontColor = '#333333';
    }

    backgrounds[i] = new Array(sheet.getLastColumn()).fill(bgColor);
    fontColors[i] = new Array(sheet.getLastColumn()).fill(fontColor);
  }

  // Aplicar cores de fundo
  for (let i = 1; i < backgrounds.length; i++) {
    const rowRange = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());
    rowRange.setBackgrounds([backgrounds[i]]);
    rowRange.setFontColors([fontColors[i]]);
  }

  // 3. Congelar primeira linha
  sheet.setFrozenRows(1);

  // 4. Ajustar largura das colunas
  const columnWidths = {
    'Telefone': 150,
    'Nome': 200,
    'Horários Oferecidos (JSON)': 250,
    'Horário Escolhido (JSON)': 250,
    'Agendamento Confirmado (JSON)': 300,
    'Status': 150,
    'Lead Status': 200
  };

  headers.forEach((header, index) => {
    if (columnWidths[header]) {
      sheet.setColumnWidth(index + 1, columnWidths[header]);
    }
  });

  // 5. Adicionar filtro automático
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).createFilter();

  // 6. Centralizar e ajustar altura das linhas
  sheet.setRowHeight(1, 35); // Header
  for (let i = 2; i <= sheet.getLastRow(); i++) {
    sheet.setRowHeight(i, 25);
  }

  SpreadsheetApp.getUi().alert('✅ Formatação concluída! Cores aplicadas com sucesso.\n\nPaleta da marca:\n🟨 Dourado: Ativo\n🟩 Verde: Confirmado\n🟥 Vermelho: Não Compareceu\n🟨 Amarelo: Pendente');
}

/**
 * Função auxiliar: colorir específicos leads por status
 * Útil se quiser re-aplicar cores depois de adicionar novos leads
 */
function recolorLeads() {
  formatLeadsSheet();
}

/**
 * Adicionar menu personalizado no Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 Formatação CRM')
    .addItem('🎨 Aplicar cores da marca', 'formatLeadsSheet')
    .addItem('🔄 Re-colorir leads', 'recolorLeads')
    .addDivider()
    .addItem('📖 Documentação', 'showDocumentation')
    .addToUi();
}

function showDocumentation() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; margin: 20px; }
      h2 { color: #6B4423; }
      .color-box { display: inline-block; width: 30px; height: 30px; margin-right: 10px; vertical-align: middle; border-radius: 4px; }
    </style>
    <h2>📖 Documentação - Formatação CRM</h2>
    <p><strong>Cores aplicadas:</strong></p>
    <ul>
      <li><span class="color-box" style="background: #d4edda;"></span> <strong>Verde</strong> = Agendamento Confirmado</li>
      <li><span class="color-box" style="background: #fff3cd;"></span> <strong>Amarelo</strong> = Pendente de Confirmação</li>
      <li><span class="color-box" style="background: #f8d7da;"></span> <strong>Vermelho</strong> = Não Compareceu</li>
      <li><span class="color-box" style="background: #E8DDD0;"></span> <strong>Bege</strong> = Sem Agendamento</li>
    </ul>
    <p><strong>O que foi feito:</strong></p>
    <ul>
      <li>✅ Header formatado em café (#6B4423) com texto branco</li>
      <li>✅ Primeira linha congelada</li>
      <li>✅ Filtro automático habilitado</li>
      <li>✅ Colunas JSON ajustadas de largura</li>
      <li>✅ Cores aplicadas por status</li>
    </ul>
  `);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Documentação CRM');
}
