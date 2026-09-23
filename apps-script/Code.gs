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

// Notificações por e-mail (pedido do usuário 2026-09-23): aviso instantâneo
// a cada agendamento novo + resumo diário do dia seguinte, enviado no
// horário definido em DAILY_SUMMARY_HOUR. Usa MailApp (cota gratuita do
// Apps Script) — não depende de nenhuma configuração do Google Calendar.
const NOTIFICATION_EMAIL = 'dramarcellabarros@gmail.com';
const DAILY_SUMMARY_HOUR = 20; // 20h — horário do resumo do dia seguinte

// Mesmas colunas/ordem de src/googleSheetsClient.js#DEFAULT_COLUMNS — mudar
// aqui exige mudar lá também (e vice-versa) para não dessincronizar.
// "Tipo de Atendimento" (2026-09-23, 2ª auditoria UX): antes só existia no
// título do evento do Calendar, impossível de contar/relatar a partir só
// da planilha. upsertLeadRow() migra a planilha real sozinha na primeira
// escrita depois desta mudança — não precisa editar a planilha na mão.
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
  'Tipo de Atendimento',
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
    if (payload.action === 'updateLead') {
      return jsonResponse(updateLead(payload));
    }
    if (payload.action === 'completeAppointment') {
      return jsonResponse(completeAppointment(payload));
    }
    if (payload.action === 'rescheduleAppointment') {
      return jsonResponse(rescheduleAppointmentAction(payload));
    }
    if (payload.action === 'cancelAppointment') {
      return jsonResponse(cancelAppointmentAction(payload));
    }
    if (payload.action === 'reopenAppointment') {
      return jsonResponse(reopenAppointment(payload));
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
// Tipos de atendimento identificáveis na agenda (seção nova, 2026-09-23,
// a pedido do usuário: já existem pacientes com procedimento agendado e
// retorno, não só avaliação inicial). O rótulo vira o prefixo do título do
// evento no Calendar — é esse prefixo que o Dashboard usa pra colorir/
// identificar o tipo na tela da Agenda (ver TYPE_KEYWORDS no HTML).
const APPOINTMENT_TYPE_LABELS = {
  AVALIACAO: 'Avaliação',
  PROCEDIMENTO: 'Procedimento',
  RETORNO: 'Retorno',
};

function createAppointment(payload) {
  const { phone, name, source, date, start, end, appointmentType } = payload;
  if (!phone || !name || !date || !start || !end) {
    return { ok: false, error: 'Campos obrigatórios: phone, name, date, start, end.' };
  }
  const typeLabel = APPOINTMENT_TYPE_LABELS[appointmentType] || APPOINTMENT_TYPE_LABELS.AVALIACAO;

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

  const event = calendar.createEvent(`${typeLabel} - ${name}`, startDateTime, endDateTime, {
    description: `Agendado via CRM. Telefone: ${phone}. Origem: ${source || '—'}. Tipo: ${typeLabel}.`,
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
    leadStatus: '',
    appointmentType: typeLabel,
    chosenSlot: { date, start, end },
    confirmedAppointment,
  });

  // Falha no e-mail nunca deve derrubar a criação do agendamento — o
  // evento e a linha na planilha já foram gravados com sucesso acima,
  // isso aqui é só um aviso extra.
  try {
    notifyNewAppointment({ name, phone, source, typeLabel, date, start, end, htmlLink: confirmedAppointment.htmlLink });
  } catch (err) {
    // Silencioso de propósito — ver nota acima.
  }

  return { ok: true, event: confirmedAppointment };
}

/**
 * E-mail instantâneo a cada agendamento novo — pedido do usuário
 * 2026-09-23: a notificação nativa "Novos eventos" do Google Calendar só
 * manda e-mail e depende de configuração manual na conta dela; isso aqui
 * dispara na hora, direto do momento em que o CRM cria o evento.
 */
function notifyNewAppointment({ name, phone, source, typeLabel, date, start, end, htmlLink }) {
  if (!NOTIFICATION_EMAIL) return;
  const dateLabel = Utilities.formatDate(parseLocalDateTime(date, start), TIME_ZONE, 'dd/MM/yyyy');
  const subject = `Novo agendamento — ${name}`;
  const body = [
    `Novo agendamento criado no CRM.`,
    ``,
    `Paciente: ${name}`,
    `Tipo: ${typeLabel}`,
    `Data: ${dateLabel}`,
    `Horário: ${start} às ${end}`,
    `Telefone: ${phone}`,
    `Origem: ${source || '—'}`,
    ``,
    `Ver na agenda: ${htmlLink}`,
  ].join('\n');
  MailApp.sendEmail(NOTIFICATION_EMAIL, subject, body);
}

/**
 * "Baixa" um atendimento — marca o evento do Calendar como concluído
 * prefixando o título com "✅ " (mantém o resto do título intacto, então
 * detectAppointmentType() no Dashboard continua reconhecendo o tipo).
 * Não apaga nem move o evento — só sinaliza visualmente que já aconteceu.
 * Pedido do usuário 2026-09-23: fluxo de "baixar" avaliação/procedimento
 * e, dali, encadear o próximo agendamento (procedimento ou retorno).
 */
// Os 3 prefixos que tiram um evento da Agenda ativa e mandam pro histórico
// do Dashboard — cada um com um significado diferente, todos reversíveis
// (ver reopenAppointment). Definidos num só lugar pra não dessincronizar
// com o Dashboard (que precisa reconhecer os mesmos símbolos).
const STATUS_PREFIXES = { DONE: '✅', CANCELED: '❌', NO_SHOW: '🚫' };

/** Prefixa (ou reprefixa) o título de um evento com um dos STATUS_PREFIXES. */
function markCalendarEvent(eventId, prefix) {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) return { ok: false, error: 'Agenda não encontrada — verifique CALENDAR_ID no script.' };

  const event = calendar.getEventById(eventId);
  if (!event) return { ok: false, error: 'Evento não encontrado na agenda (pode já ter sido apagado).' };

  const bareTitle = stripStatusPrefix(event.getTitle());
  event.setTitle(`${prefix} ${bareTitle}`);
  return { ok: true };
}

/** Remove qualquer um dos STATUS_PREFIXES do início do título, se houver. */
function stripStatusPrefix(title) {
  let t = (title || '').trim();
  Object.values(STATUS_PREFIXES).forEach((p) => {
    if (t.startsWith(p)) t = t.slice(p.length).trim();
  });
  return t;
}

function completeAppointment(payload) {
  const { eventId } = payload;
  if (!eventId) return { ok: false, error: 'Campo obrigatório: eventId.' };
  return markCalendarEvent(eventId, STATUS_PREFIXES.DONE);
}

/**
 * "Desfaz" uma baixa, desmarcação ou não-comparecimento — remove o prefixo
 * do título, o evento volta a aparecer na Agenda ativa como se nada
 * tivesse acontecido. Pedido da 2ª auditoria UX 2026-09-23: um toque
 * errado em "Desmarcar" (perto de "Remarcar" na lista) não tinha volta
 * pela interface antes disso.
 */
function reopenAppointment(payload) {
  const { eventId } = payload;
  if (!eventId) return { ok: false, error: 'Campo obrigatório: eventId.' };

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) return { ok: false, error: 'Agenda não encontrada — verifique CALENDAR_ID no script.' };

  const event = calendar.getEventById(eventId);
  if (!event) return { ok: false, error: 'Evento não encontrado na agenda (pode já ter sido apagado).' };

  event.setTitle(stripStatusPrefix(event.getTitle()));
  return { ok: true };
}

/**
 * Remarca um atendimento existente para nova data/horário — atualiza o
 * MESMO evento (mesmo id/link, event.setTime), nunca cria um novo. Pedido
 * do usuário 2026-09-23.
 */
function rescheduleAppointmentAction(payload) {
  const { eventId, date, start, end } = payload;
  if (!eventId || !date || !start || !end) {
    return { ok: false, error: 'Campos obrigatórios: eventId, date, start, end.' };
  }

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) return { ok: false, error: 'Agenda não encontrada — verifique CALENDAR_ID no script.' };

  const event = calendar.getEventById(eventId);
  if (!event) return { ok: false, error: 'Evento não encontrado na agenda (pode já ter sido apagado).' };

  const newStart = parseLocalDateTime(date, start);
  const newEnd = parseLocalDateTime(date, end);
  if (newEnd <= newStart) return { ok: false, error: 'Horário de término precisa ser depois do início.' };

  // Mesma checagem de conflito de createAppointment, excluindo o próprio
  // evento (senão ele sempre "colidiria" consigo mesmo).
  const conflicting = calendar.getEvents(newStart, newEnd).filter((e) => e.getId() !== event.getId());
  if (conflicting.length > 0) {
    return { ok: false, error: 'Esse horário já está ocupado na agenda. Escolha outro.' };
  }

  event.setTime(newStart, newEnd);

  // Atualiza o lead na planilha (se der pra identificar o telefone na
  // descrição do evento), pra tabela/painel de detalhe do Dashboard
  // ficarem coerentes com a nova data sem precisar editar a mão.
  const phoneMatch = (event.getDescription() || '').match(/Telefone:\s*(\+?\d+)/);
  if (phoneMatch) {
    upsertLeadRow({
      phone: phoneMatch[1],
      chosenSlot: { date, start, end },
      confirmedAppointment: {
        id: event.getId(),
        htmlLink: buildEventHtmlLink(event.getId()),
        start: { dateTime: formatIso(newStart), timeZone: TIME_ZONE },
        end: { dateTime: formatIso(newEnd), timeZone: TIME_ZONE },
      },
    });
  }

  return { ok: true };
}

/**
 * Desmarca um atendimento — NÃO apaga o evento (revisado 2026-09-23 a
 * pedido do usuário: "desmarcar deve ser baixado também e ir pro
 * histórico"). Prefixa o título com "❌ " (mesmo princípio de
 * completeAppointment, que usa "✅ ") — some da agenda ativa do Dashboard e
 * passa a aparecer na aba de histórico, mantendo o registro em vez de
 * perdê-lo.
 */
function cancelAppointmentAction(payload) {
  const { eventId } = payload;
  if (!eventId) return { ok: false, error: 'Campo obrigatório: eventId.' };

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) return { ok: false, error: 'Agenda não encontrada — verifique CALENDAR_ID no script.' };

  const event = calendar.getEventById(eventId);
  if (!event) return { ok: false, error: 'Evento não encontrado na agenda (pode já ter sido apagado).' };

  event.setTitle(`${STATUS_PREFIXES.CANCELED} ${stripStatusPrefix(event.getTitle())}`);

  // Sincroniza com a planilha (2ª auditoria UX 2026-09-23): sem isso, a
  // tabela do Dashboard continuava mostrando "Confirmado" pra um
  // atendimento já desmarcado na Agenda — dois sistemas de status que não
  // se falavam. Limpa o agendamento confirmado (não é mais verdade) e
  // marca o motivo; só roda se der pra achar o telefone na descrição do
  // evento (eventos muito antigos, de antes desse campo existir, não têm).
  const phoneMatch = (event.getDescription() || '').match(/Telefone:\s*(\+?\d+)/);
  if (phoneMatch) {
    upsertLeadRow({ phone: phoneMatch[1], leadStatus: 'desmarcado', confirmedAppointment: null });
  }

  return { ok: true };
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

/**
 * Edita campos de negócio de um lead já existente (nome, origem, status,
 * atendimento manual) — a partir do painel de detalhe do Dashboard. Nunca
 * cria linha nova (diferente de upsertLeadRow): se o telefone não existe
 * ainda, falha explicitamente em vez de inserir um registro incompleto.
 * Só atualiza os campos que vierem presentes no payload — os demais ficam
 * como estavam.
 */
function updateLead(payload) {
  const { phone, name, source, leadStatus, manualOnlyNote } = payload;
  if (!phone) return { ok: false, error: 'Campo obrigatório: phone.' };

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const phoneCol = headers.indexOf('Telefone');

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][phoneCol]) === String(phone)) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex === -1) return { ok: false, error: 'Lead não encontrado para esse telefone.' };

  const fieldToHeader = {
    name: 'Nome',
    source: 'Origem',
    leadStatus: 'Status',
    manualOnlyNote: 'Atendimento Manual (motivo)',
  };

  for (const key in fieldToHeader) {
    if (payload[key] !== undefined) {
      const colIndex = headers.indexOf(fieldToHeader[key]);
      if (colIndex !== -1) sheet.getRange(rowIndex + 1, colIndex + 1).setValue(payload[key]);
    }
  }

  // Sincroniza com a Agenda (2ª auditoria UX 2026-09-23): marcar "não
  // compareceu" aqui, no painel de detalhe, também tira o evento
  // correspondente da Agenda ativa do Dashboard — sem isso, o evento
  // continuava pedindo "Baixar/Remarcar/Desmarcar" como se nada tivesse
  // acontecido. Usa 🚫, diferente de ✅ (concluído normal) e ❌
  // (desmarcado), pra manter o motivo real visível no histórico. Falha
  // aqui nunca derruba a atualização da planilha, que já aconteceu acima.
  if (leadStatus === 'não compareceu') {
    try {
      const confirmedCol = headers.indexOf('Agendamento Confirmado (JSON)');
      const raw = confirmedCol !== -1 ? data[rowIndex][confirmedCol] : '';
      const confirmedAppointment = raw ? JSON.parse(raw) : null;
      if (confirmedAppointment && confirmedAppointment.id) {
        markCalendarEvent(confirmedAppointment.id, STATUS_PREFIXES.NO_SHOW);
      }
    } catch (err) {
      // Silencioso de propósito — a planilha já foi atualizada com sucesso.
    }
  }

  return { ok: true };
}

/**
 * Migra a planilha real sozinha se COLUMNS ganhou colunas novas desde a
 * última vez (ex.: "Tipo de Atendimento", 2ª auditoria UX 2026-09-23) —
 * completa os cabeçalhos que faltam no fim da linha 1. Idempotente e
 * barata, roda toda vez sem custo perceptível.
 */
function ensureSheetColumns(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < COLUMNS.length) {
    const missing = COLUMNS.slice(lastCol);
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }
}

function upsertLeadRow(fields) {
  const sheet = getSheet();
  ensureSheetColumns(sheet);
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
    leadStatus: 'Status',
    appointmentType: 'Tipo de Atendimento',
  };

  const existingRow = rowIndex >= 0 ? data[rowIndex] : COLUMNS.map(() => '');

  // hasOwnProperty (não só "!== undefined") pra distinguir "campo não veio"
  // (preserva o valor antigo) de "campo veio como null" (limpa a célula de
  // propósito) — precisa pros campos JSON quando um cancelamento precisa
  // apagar o agendamento confirmado, não só sobrescrever.
  const has = (key) => Object.prototype.hasOwnProperty.call(fields, key);

  const newRow = COLUMNS.map((header, colIndex) => {
    for (const key in fieldToHeader) {
      if (fieldToHeader[key] === header && has(key)) return fields[key];
    }
    if (header === 'Horário Escolhido (JSON)' && has('chosenSlot')) {
      return fields.chosenSlot ? JSON.stringify(fields.chosenSlot) : '';
    }
    if (header === 'Agendamento Confirmado (JSON)' && has('confirmedAppointment')) {
      return fields.confirmedAppointment ? JSON.stringify(fields.confirmedAppointment) : '';
    }
    return existingRow[colIndex] || '';
  });

  const targetRowNumber = rowIndex === -1 ? sheet.getLastRow() + 1 : rowIndex + 1;
  const range = sheet.getRange(targetRowNumber, 1, 1, newRow.length);
  // Formata a linha como texto simples ANTES de escrever — sem isso, o
  // Sheets interpreta "+5515..." como início de fórmula e corta o "+"
  // (mesmo bug já resolvido do lado Node em
  // src/googleSheetsClient.js#persistLead, que usa valueInputOption=RAW).
  range.setNumberFormat('@');
  range.setValues([newRow]);
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

/**
 * Resumo dos atendimentos de amanhã, enviado no horário fixo de
 * DAILY_SUMMARY_HOUR — pedido do usuário 2026-09-23: o lembrete nativo do
 * Calendar só sabe contar "X horas antes do evento", nunca um horário
 * fixo da noite anterior independente do horário de cada consulta. Chamada
 * automaticamente pelo gatilho criado em setupDailyTrigger() — nunca
 * precisa rodar isso na mão no dia a dia.
 */
function sendTomorrowSummary() {
  if (!NOTIFICATION_EMAIL) return;

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) return;

  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const dayStart = new Date(tomorrow.getFullYear(), tomorrow.getMonth(), tomorrow.getDate(), 0, 0, 0);
  const dayEnd = new Date(tomorrow.getFullYear(), tomorrow.getMonth(), tomorrow.getDate(), 23, 59, 59);

  // Ignora eventos já baixados/desmarcados (✅/❌) — não fazem sentido num
  // resumo do que ainda vai acontecer amanhã.
  const events = calendar
    .getEvents(dayStart, dayEnd)
    .filter((e) => {
      const title = e.getTitle() || '';
      return !title.startsWith('✅') && !title.startsWith('❌');
    })
    .sort((a, b) => a.getStartTime() - b.getStartTime());

  const dateLabel = Utilities.formatDate(dayStart, TIME_ZONE, 'dd/MM/yyyy');
  const subject =
    events.length > 0
      ? `Agenda de amanhã (${dateLabel}) — ${events.length} atendimento(s)`
      : `Agenda de amanhã (${dateLabel}) — nenhum atendimento`;

  const lines = [`Resumo dos atendimentos de amanhã (${dateLabel}):`, ''];
  if (events.length === 0) {
    lines.push('Nenhum atendimento agendado.');
  } else {
    events.forEach((e) => {
      const time = e.isAllDayEvent() ? 'Dia todo' : Utilities.formatDate(e.getStartTime(), TIME_ZONE, 'HH:mm');
      lines.push(`${time} — ${e.getTitle()}`);
    });
  }

  MailApp.sendEmail(NOTIFICATION_EMAIL, subject, lines.join('\n'));
}

/**
 * Configura o gatilho que roda sendTomorrowSummary() todo dia, no horário
 * de DAILY_SUMMARY_HOUR. RODE ESTA FUNÇÃO MANUALMENTE UMA ÚNICA VEZ: no
 * editor do Apps Script, selecione "setupDailyTrigger" no menu de funções
 * (ao lado do botão Executar) e clique em Executar. Depois disso o
 * gatilho fica salvo — não precisa rodar de novo. Remove qualquer gatilho
 * anterior desta mesma função antes de criar um novo, pra nunca duplicar
 * e mandar o resumo em dobro.
 */
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'sendTomorrowSummary')
    .forEach((t) => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sendTomorrowSummary').timeBased().atHour(DAILY_SUMMARY_HOUR).everyDays(1).create();
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
