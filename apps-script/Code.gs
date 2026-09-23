/**
 * Google Apps Script vinculado à planilha de Leads — expõe um endpoint web
 * seguro (rodando com a identidade de quem publica o script, dono da
 * planilha e da agenda) para o Dashboard HTML poder ESCREVER agendamentos:
 * criar o evento real no Google Calendar e gravar/atualizar a linha
 * correspondente na aba "Leads".
 *
 * Por que isso existe: a API Key usada pelo Dashboard só lê dados públicos
 * (Sheets/Calendar). O Google nunca permite escrita com API Key simples —
 * escrita exige OAuth de usuário ou, como aqui, um script rodando com
 * permissão de dono. Isso evita expor qualquer credencial sensível no
 * código-fonte público do Dashboard (GitHub Pages).
 *
 * Setup (ver apps-script/SETUP.md para o passo a passo com screenshots):
 * 1. Na planilha real, Extensões → Apps Script.
 * 2. Colar este arquivo inteiro no editor (substituindo o conteúdo padrão).
 * 3. Ajustar CALENDAR_ID abaixo se a agenda mudar.
 * 4. Implantar → Nova implantação → Tipo "Aplicativo da web" → Executar
 *    como "Eu" → Quem tem acesso "Qualquer pessoa" → Implantar.
 * 5. Copiar a URL gerada (termina em /exec) e colar no Dashboard, em
 *    Config → "URL do Apps Script".
 */

const SHEET_NAME = 'Leads';
const CALENDAR_ID = '2c135b41ae97d7bf8c4b10be19c21e59cc77ee5e08a16a9126e0f5dc24179285@group.calendar.google.com';
const TIME_ZONE = 'America/Sao_Paulo';

// Mesmas colunas/ordem de src/googleSheetsClient.js#DEFAULT_COLUMNS — mudar
// aqui exige mudar lá também (e vice-versa) para não dessincronizar.
const COLUMNS = [
  'Telefone',
  'Nome',
  'Atendimento Manual (motivo)',
  'Origem',
  'Interesse/Procedimento',
  'Objetivo',
  'Última Intenção',
  'Score',
  'Estado da Conversa',
  'Estágio do Funil',
  'Status',
  'Primeiro Contato',
  'Última Interação',
  'Horários Oferecidos (JSON)',
  'Horário Escolhido (JSON)',
  'Agendamento Confirmado (JSON)',
];

// Mesmo horário de funcionamento de config/business_rules.json#scheduling —
// mudar lá também exige mudar aqui (Apps Script não lê o .env do projeto).
const BUSINESS_HOURS = {
  0: null, // domingo
  1: { open: '09:00', close: '19:00' },
  2: { open: '09:00', close: '19:00' },
  3: { open: '09:00', close: '19:00' },
  4: { open: '09:00', close: '19:00' },
  5: { open: '09:00', close: '19:00' },
  6: { open: '09:00', close: '13:00' },
};

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.action === 'createAppointment') {
      return jsonResponse(createAppointment(payload));
    }
    if (payload.action === 'markNoShow') {
      return jsonResponse(markNoShow(payload));
    }

    return jsonResponse({ ok: false, error: 'Ação desconhecida: ' + payload.action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    if (e.parameter.action === 'availability') {
      return jsonResponse(getAvailability(e.parameter.fromDate, Number(e.parameter.daysAhead || 7)));
    }
    return jsonResponse({ ok: false, error: 'Ação desconhecida' });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

/**
 * Cria o evento real no Calendar e grava/atualiza a linha do lead na
 * planilha — equivalente ao que src/schedulingService.js#bookAppointment +
 * src/googleSheetsClient.js#persistLead fazem juntos, só que rodando aqui
 * (com permissão de dono) em vez de via Service Account.
 */
function createAppointment(payload) {
  const { phone, name, source, date, start, end } = payload;
  if (!phone || !name || !date || !start || !end) {
    return { ok: false, error: 'Campos obrigatórios: phone, name, date, start, end.' };
  }

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) {
    return { ok: false, error: 'Agenda não encontrada — verifique CALENDAR_ID no script.' };
  }

  const startDateTime = parseLocalDateTime(date, start);
  const endDateTime = parseLocalDateTime(date, end);

  if (endDateTime <= startDateTime) {
    return { ok: false, error: 'Horário de término precisa ser depois do início.' };
  }

  // Nunca cria em cima de um evento já existente — mesma garantia de
  // src/schedulingService.js#listAvailableSlots, checada aqui de novo pois
  // esta é a via de escrita real.
  const conflicting = calendar.getEvents(startDateTime, endDateTime);
  if (conflicting.length > 0) {
    return { ok: false, error: 'Esse horário já está ocupado na agenda. Escolha outro.' };
  }

  const event = calendar.createEvent(`Avaliação - ${name}`, startDateTime, endDateTime, {
    description: `Agendado via CRM. Telefone: ${phone}. Origem: ${source || '—'}.`,
  });

  const confirmedAppointment = {
    id: event.getId(),
    htmlLink: buildEventHtmlLink(event.getId()),
    start: { dateTime: formatIso(startDateTime), timeZone: TIME_ZONE },
    end: { dateTime: formatIso(endDateTime), timeZone: TIME_ZONE },
  };

  upsertLeadRow({
    phone,
    name,
    source: source || '',
    conversationState: 'CONFIRMED',
    chosenSlot: { date, start, end },
    confirmedAppointment,
  });

  return { ok: true, event: confirmedAppointment };
}

/** Marca leadStatus = "não compareceu" numa linha já existente, sem mexer no resto. */
function markNoShow(payload) {
  const { phone } = payload;
  if (!phone) return { ok: false, error: 'Campo obrigatório: phone.' };

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const phoneCol = headers.indexOf('Telefone');
  const statusCol = headers.indexOf('Status');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][phoneCol]) === String(phone)) {
      sheet.getRange(i + 1, statusCol + 1).setValue('não compareceu');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Lead não encontrado para esse telefone.' };
}

function upsertLeadRow(fields) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const phoneCol = headers.indexOf('Telefone');

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][phoneCol]) === String(fields.phone)) {
      rowIndex = i;
      break;
    }
  }

  const fieldToHeader = {
    phone: 'Telefone',
    name: 'Nome',
    source: 'Origem',
    conversationState: 'Estado da Conversa',
  };

  const existingRow = rowIndex >= 0 ? data[rowIndex] : COLUMNS.map(() => '');

  const newRow = COLUMNS.map((header, colIndex) => {
    for (const key in fieldToHeader) {
      if (fieldToHeader[key] === header && fields[key] !== undefined) return fields[key];
    }
    if (header === 'Horário Escolhido (JSON)' && fields.chosenSlot) return JSON.stringify(fields.chosenSlot);
    if (header === 'Agendamento Confirmado (JSON)' && fields.confirmedAppointment) return JSON.stringify(fields.confirmedAppointment);
    return existingRow[colIndex] || '';
  });

  if (rowIndex === -1) {
    sheet.appendRow(newRow);
  } else {
    sheet.getRange(rowIndex + 1, 1, 1, newRow.length).setValues([newRow]);
  }
}

/**
 * Mesma lógica de src/availability.js#computeAvailableSlots, reimplementada
 * aqui porque Apps Script não importa módulos do projeto Node — mudanças
 * numa precisam ser espelhadas na outra.
 */
function getAvailability(fromDateStr, daysAhead) {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) return { ok: false, error: 'Agenda não encontrada.' };

  const fromDate = fromDateStr ? new Date(fromDateStr + 'T00:00:00') : new Date();
  const results = [];

  for (let offset = 0; offset < daysAhead; offset++) {
    const day = new Date(fromDate);
    day.setDate(day.getDate() + offset);
    const weekday = day.getDay();
    const hours = BUSINESS_HOURS[weekday];
    if (!hours) continue;

    const dateStr = Utilities.formatDate(day, TIME_ZONE, 'yyyy-MM-dd');
    const dayStart = parseLocalDateTime(dateStr, hours.open);
    const dayEnd = parseLocalDateTime(dateStr, hours.close);
    const events = calendar.getEvents(dayStart, dayEnd);

    const slots = [];
    let cursor = new Date(dayStart);
    const now = new Date();

    while (cursor.getTime() + 60 * 60000 <= dayEnd.getTime()) {
      const slotEnd = new Date(cursor.getTime() + 60 * 60000);
      const overlaps = events.some((ev) => ev.getStartTime() < slotEnd && ev.getEndTime() > cursor);
      if (!overlaps && cursor > now) {
        slots.push({ start: Utilities.formatDate(cursor, TIME_ZONE, 'HH:mm'), end: Utilities.formatDate(slotEnd, TIME_ZONE, 'HH:mm') });
      }
      cursor = slotEnd;
    }

    if (slots.length > 0) results.push({ date: dateStr, slots });
  }

  return { ok: true, availability: results };
}

function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Aba "${SHEET_NAME}" não encontrada.`);
  return sheet;
}

function parseLocalDateTime(dateStr, timeStr) {
  return new Date(`${dateStr}T${timeStr}:00-03:00`);
}

function formatIso(date) {
  return Utilities.formatDate(date, TIME_ZONE, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function buildEventHtmlLink(eventId) {
  const base64 = Utilities.base64Encode(`${eventId} ${CALENDAR_ID}`).replace(/=+$/, '');
  return `https://www.google.com/calendar/event?eid=${base64}`;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
