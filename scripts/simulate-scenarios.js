'use strict';

/**
 * Simulação end-to-end de TODAS as possibilidades de atendimento da IA,
 * rodando o orchestrator.js real contra a planilha e a agenda REAIS (não
 * mocks de infraestrutura). Como LLM_PROVIDER ainda não está configurado
 * (.env), o classificador de intenção e o gerador de texto são stubs
 * determinísticos injetados por cenário — mas toda a máquina de estados,
 * scoring, guardrails, Google Sheets e Google Calendar são os módulos reais
 * de src/.
 *
 * Cada cenário usa um telefone de teste novo (TESTE 7+), nunca reaproveita
 * os TESTE 1-6 já existentes na planilha. Eventos de agenda criados aqui são
 * nomeados com prefixo "[TESTE SIMULACAO]" para ficarem óbvios como teste.
 *
 * Uso: node scripts/simulate-scenarios.js
 */

const fs = require('fs');
const path = require('path');

const { handleIncomingMessage } = require('../src/orchestrator');
const { getAccessToken } = require('../src/googleAuth');
const { createGoogleSheetsLeadStore } = require('../src/googleSheetsClient');
const { createGoogleCalendarStore, CALENDAR_EVENTS_SCOPE } = require('../src/googleCalendarClient');
const { createSchedulingService } = require('../src/schedulingService');
const businessRules = require('../config/business_rules.json');

function loadEnv(filePath) {
  const content = fs.readFileSync(filePath, 'utf8');
  const env = {};
  for (const rawLine of content.split('\n')) {
    const line = rawLine.trim();
    if (!line || line.startsWith('#')) continue;
    const idx = line.indexOf('=');
    if (idx === -1) continue;
    const key = line.slice(0, idx).trim();
    let value = line.slice(idx + 1).trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    env[key] = value.replace(/\\n/g, '\n');
  }
  return env;
}

const env = loadEnv(path.join(__dirname, '..', '.env'));

const SPREADSHEET_ID = env.GOOGLE_SHEETS_SPREADSHEET_ID;
const SHEET_NAME = env.GOOGLE_SHEETS_SHEET_NAME || 'Leads';
const CALENDAR_ID = env.GOOGLE_CALENDAR_ID;

let cachedSheetsToken = null;
let cachedCalendarToken = null;

async function getSheetsAccessToken() {
  if (cachedSheetsToken) return cachedSheetsToken;
  cachedSheetsToken = await getAccessToken({
    clientEmail: env.GOOGLE_SHEETS_CLIENT_EMAIL,
    privateKey: env.GOOGLE_SHEETS_PRIVATE_KEY,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    fetchImpl: fetch,
  });
  return cachedSheetsToken;
}

async function getCalendarAccessToken() {
  if (cachedCalendarToken) return cachedCalendarToken;
  cachedCalendarToken = await getAccessToken({
    clientEmail: env.GOOGLE_SHEETS_CLIENT_EMAIL,
    privateKey: env.GOOGLE_SHEETS_PRIVATE_KEY,
    scope: CALENDAR_EVENTS_SCOPE,
    fetchImpl: fetch,
  });
  return cachedCalendarToken;
}

const leadStore = createGoogleSheetsLeadStore({
  spreadsheetId: SPREADSHEET_ID,
  sheetName: SHEET_NAME,
  getAccessTokenImpl: getSheetsAccessToken,
  fetchImpl: fetch,
});

const calendarStore = createGoogleCalendarStore({
  calendarId: CALENDAR_ID,
  getAccessTokenImpl: getCalendarAccessToken,
  fetchImpl: fetch,
});

const schedulingService = createSchedulingService({
  calendarStore,
  businessHoursByWeekday: businessRules.scheduling.business_hours_by_weekday,
  slotDurationMinutes: businessRules.scheduling.default_slot_duration_minutes,
  bufferMinutes: businessRules.scheduling.buffer_minutes_between_appointments,
});

// --- stubs determinísticos (sem LLM real configurado ainda) ---

function makeFakeIntentEngine(getForcedIntent) {
  return {
    async classify() {
      const forced = getForcedIntent();
      if (forced === 'LOW_CONFIDENCE') {
        return {
          intent: 'INTENT_01_CURIOSIDADE',
          procedimentoMencionado: null,
          urgenciaEmocional: 'baixa',
          requerHandoffHumano: false,
          motivoHandoff: null,
          lowConfidence: true,
        };
      }
      return {
        intent: forced,
        procedimentoMencionado: null,
        urgenciaEmocional: 'media',
        requerHandoffHumano: forced === 'INTENT_09_HUMANO',
        motivoHandoff: forced === 'INTENT_09_HUMANO' ? 'paciente pediu atendimento humano' : null,
        lowConfidence: false,
      };
    },
  };
}

async function fakeGenerateResponse({ action, availableSlots, bookedAppointment, confirmationDeclined }) {
  switch (action.action) {
    case 'ACOLHER':
      return 'Oi! Que bom te ver por aqui. Me conta, o que te trouxe ate a gente hoje?';
    case 'INVESTIGAR':
      return 'Entendi. Me ajuda a entender melhor o que voce esta buscando?';
    case 'EDUCAR':
      return 'Otima pergunta, vou te explicar rapidinho como funciona.';
    case 'APRESENTAR_VALOR':
      return 'Pelo que voce me contou, uma avaliacao individual vai ajudar a definir o melhor caminho pra voce.';
    case 'TRATAR_OBJECAO':
      return 'Entendo sua preocupacao, e super comum. Posso te explicar melhor como funciona?';
    case 'OFERECER_AVALIACAO':
      return 'Que tal agendarmos uma avaliacao pra conversarmos com calma sobre isso?';
    case 'OFERECER_HORARIO':
      return availableSlots && availableSlots.length
        ? `Temos horarios disponiveis, por exemplo ${availableSlots[0].date} as ${availableSlots[0].slots[0].start}. Qual funciona melhor pra voce?`
        : 'Deixa eu verificar os horarios disponiveis pra voce.';
    case 'CONFIRMAR':
      if (bookedAppointment) return 'Prontinho, seu horario esta confirmado! Te esperamos aqui.';
      if (confirmationDeclined) return 'Sem problemas, me avisa quando quiser escolher outro horario.';
      return 'Posso confirmar esse horario pra voce?';
    case 'HANDOFF':
      return 'Vou te conectar com a nossa equipe agora, um momento.';
    default:
      return 'Certo, entendi.';
  }
}

async function fakeMatchChosenSlot({ offeredSlots }) {
  return offeredSlots && offeredSlots.length ? offeredSlots[0] : null;
}

function makeFakeMatchConfirmation(getReply) {
  return async function matchConfirmation() {
    return getReply();
  };
}

async function persistLead(lead) {
  await leadStore.persistLead(lead);
}

async function noopSendMessage() {}
async function noopNotifyHandoff() {}

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

async function runScenario({ label, phone, name, source, turns }) {
  console.log(`\n=== ${label} (${phone} — ${name}) ===`);
  let lead = { phone, name, source };
  let consecutiveLowConfidence = 0;
  let turnIndex = 0;

  for (const turn of turns) {
    turnIndex++;
    const intentEngine = makeFakeIntentEngine(() => turn.intent);
    const matchConfirmation = makeFakeMatchConfirmation(() => (turn.confirmReply === undefined ? null : turn.confirmReply));

    const deps = {
      intentEngine,
      generateResponse: fakeGenerateResponse,
      persistLead,
      sendMessage: noopSendMessage,
      notifyHumanHandoff: noopNotifyHandoff,
      getAvailableSlots: async () => schedulingService.listAvailableSlots({ fromDate: todayIso(), daysAhead: 7 }),
      bookAppointment: async ({ lead: l, slot }) =>
        schedulingService.bookAppointment({
          date: slot.date,
          start: slot.start,
          end: slot.end,
          summary: `[TESTE SIMULACAO] Avaliacao - ${l.name}`,
          description: `Lead de teste automatizado (simulacao end-to-end). Telefone: ${l.phone}. Origem: ${l.source}.`,
        }),
      matchChosenSlot: fakeMatchChosenSlot,
      matchConfirmation,
    };

    const result = await handleIncomingMessage({
      leadRecord: lead,
      incomingMessage: { text: turn.message },
      deps,
      consecutiveLowConfidence,
    });

    consecutiveLowConfidence = result.lowConfidenceStreak || 0;
    lead = result.lead;

    const actionLabel = result.skipped ? `(pulado: ${result.skipReason})` : result.action ? result.action.action : '(sem acao)';
    console.log(`  T${turnIndex}: "${turn.message}"`);
    console.log(`     -> estado=${lead.conversationState || '—'} acao=${actionLabel} handoff=${result.handoff}`);
    if (result.bookedAppointment) {
      console.log(`     -> EVENTO CRIADO NA AGENDA: ${result.bookedAppointment.htmlLink}`);
    }
    if (result.guardrailViolations) {
      console.log(`     -> GUARDRAIL BLOQUEOU: ${JSON.stringify(result.guardrailViolations)}`);
    }

    if (turn.patch) {
      lead = { ...lead, ...turn.patch(lead) };
    }
  }

  return lead;
}

async function runPreregisteredHumanOnlyScenario() {
  const label = 'TESTE 14 — Contato marcado Atendimento Manual (IA nunca responde)';
  const phone = '+5515900000114';
  const name = 'TESTE 14 - Historico Direto Dra';
  console.log(`\n=== ${label} (${phone} — ${name}) ===`);

  await leadStore.persistLead({
    phone,
    name,
    source: 'WHATSAPP',
    manualOnlyNote: 'Atendida diretamente pela Dra. Marcella — pedido da equipe, IA nao deve responder',
  });
  console.log('  Cadastro manual gravado na planilha (coluna Atendimento Manual preenchida).');

  const leadRecord = await leadStore.findLeadByPhone(phone);

  const deps = {
    intentEngine: makeFakeIntentEngine(() => 'INTENT_02_INTERESSE'),
    generateResponse: fakeGenerateResponse,
    persistLead,
    sendMessage: noopSendMessage,
    notifyHumanHandoff: noopNotifyHandoff,
  };

  const result = await handleIncomingMessage({
    leadRecord,
    incomingMessage: { text: 'Oi, gostaria de agendar uma avaliacao' },
    deps,
  });

  console.log(`  T1: "Oi, gostaria de agendar uma avaliacao"`);
  console.log(`     -> pulado pela IA = ${result.skipped} (motivo: ${result.skipReason})`);
  return leadRecord;
}

async function markNoShow(lead) {
  const updated = { ...lead, leadStatus: 'não compareceu' };
  await leadStore.persistLead(updated);
  console.log(`  -> Atualizacao pos-consulta: leadStatus = "nao compareceu" (agendamento confirmado permanece registrado).`);
  return updated;
}

(async () => {
  console.log('Iniciando simulacao completa de todos os cenarios de atendimento...');
  console.log(`Planilha: ${SPREADSHEET_ID} (aba ${SHEET_NAME})`);
  console.log(`Agenda: ${CALENDAR_ID}`);

  const results = {};

  try {
    results.teste7 = await runScenario({
      label: 'TESTE 7 — Fluxo completo (curiosidade -> necessidade -> agendamento confirmado)',
      phone: '+5515900000107',
      name: 'TESTE 7 - Juliana Prado',
      source: 'INSTAGRAM',
      turns: [
        { message: 'Oi, vi o story de voces sobre o Skinglow', intent: 'INTENT_01_CURIOSIDADE' },
        { message: 'O que e o Skinglow, pra que serve?', intent: 'INTENT_01_CURIOSIDADE' },
        { message: 'Entendi! Eu tenho bastante flacidez, acho que preciso disso', intent: 'INTENT_03_NECESSIDADE' },
        { message: 'Faz sentido, quero entender melhor como funciona a avaliacao', intent: 'INTENT_02_INTERESSE' },
        { message: 'Quero agendar', intent: 'INTENT_06_AGENDAMENTO' },
        { message: 'Pode ser esse horario, o primeiro', intent: 'INTENT_06_AGENDAMENTO' },
        { message: 'Sim, confirmo', intent: 'INTENT_06_AGENDAMENTO', confirmReply: true },
      ],
    });
  } catch (err) {
    console.error('  ERRO no TESTE 7:', err.message);
  }

  try {
    results.teste8 = await runScenario({
      label: 'TESTE 8 — Objecao de preco resolvida, depois agenda',
      phone: '+5515900000108',
      name: 'TESTE 8 - Renata Souza',
      source: 'WHATSAPP',
      turns: [
        { message: 'Oi, queria informacao sobre botox', intent: 'INTENT_04_PROCEDIMENTO_DEFINIDO' },
        { message: 'Quero saber mais sobre o procedimento', intent: 'INTENT_04_PROCEDIMENTO_DEFINIDO' },
        { message: 'Faz sentido pra mim', intent: 'INTENT_04_PROCEDIMENTO_DEFINIDO' },
        { message: 'Entendi o valor disso', intent: 'INTENT_02_INTERESSE' },
        {
          message: 'Poxa, mas achei que ia ser caro...',
          intent: 'INTENT_08_OBJECAO',
          patch: (lead) => ({ ...lead, pendingObjectionResolved: true }),
        },
        { message: 'Ah entendi, faz sentido investir nisso', intent: 'INTENT_02_INTERESSE' },
        { message: 'Quero agendar uma avaliacao', intent: 'INTENT_06_AGENDAMENTO' },
        { message: 'Pode ser esse horario', intent: 'INTENT_06_AGENDAMENTO' },
        { message: 'Sim, confirmo', intent: 'INTENT_06_AGENDAMENTO', confirmReply: true },
      ],
    });
  } catch (err) {
    console.error('  ERRO no TESTE 8:', err.message);
  }

  try {
    results.teste9 = await runScenario({
      label: 'TESTE 9 — Objecao de preco NAO resolvida (trava, sem agendar)',
      phone: '+5515900000109',
      name: 'TESTE 9 - Beatriz Lima',
      source: 'SITE',
      turns: [
        { message: 'Oi, queria saber sobre valores', intent: 'INTENT_05_PRECO' },
        { message: 'Quanto custa mais ou menos?', intent: 'INTENT_05_PRECO' },
        { message: 'Mas quanto custa mesmo?', intent: 'INTENT_05_PRECO' },
        { message: 'Achei muito caro, nao sei se vale a pena', intent: 'INTENT_08_OBJECAO' },
        { message: 'Ainda acho caro, vou pensar', intent: 'INTENT_08_OBJECAO' },
      ],
    });
  } catch (err) {
    console.error('  ERRO no TESTE 9:', err.message);
  }

  try {
    results.teste10 = await runScenario({
      label: 'TESTE 10 — Alta intencao direto (pula todo o funil)',
      phone: '+5515900000110',
      name: 'TESTE 10 - Fernanda Alves',
      source: 'INSTAGRAM',
      turns: [
        { message: 'Ja decidi, quero fazer o preenchimento labial, pode agendar pra mim?', intent: 'INTENT_07_ALTA_INTENCAO' },
        { message: 'Pode ser esse horario', intent: 'INTENT_07_ALTA_INTENCAO' },
        { message: 'Sim, confirma', intent: 'INTENT_07_ALTA_INTENCAO', confirmReply: true },
      ],
    });
  } catch (err) {
    console.error('  ERRO no TESTE 10:', err.message);
  }

  try {
    results.teste11 = await runScenario({
      label: 'TESTE 11 — Escolhe horario mas desiste na confirmacao',
      phone: '+5515900000111',
      name: 'TESTE 11 - Camila Duarte',
      source: 'WHATSAPP',
      turns: [
        { message: 'Quero agendar uma avaliacao', intent: 'INTENT_06_AGENDAMENTO' },
        { message: 'Pode ser esse horario', intent: 'INTENT_06_AGENDAMENTO' },
        { message: 'Na verdade vou deixar pra depois, nao quero confirmar agora', intent: 'INTENT_06_AGENDAMENTO', confirmReply: false },
      ],
    });
  } catch (err) {
    console.error('  ERRO no TESTE 11:', err.message);
  }

  try {
    results.teste12 = await runScenario({
      label: 'TESTE 12 — Pede atendimento humano explicitamente',
      phone: '+5515900000112',
      name: 'TESTE 12 - Sandra Mattos',
      source: 'SITE',
      turns: [{ message: 'Prefiro falar direto com uma atendente, pode ser?', intent: 'INTENT_09_HUMANO' }],
    });
  } catch (err) {
    console.error('  ERRO no TESTE 12:', err.message);
  }

  try {
    results.teste13 = await runScenario({
      label: 'TESTE 13 — Baixa confianca repetida -> handoff automatico',
      phone: '+5515900000113',
      name: 'TESTE 13 - Patricia Gomes',
      source: 'INSTAGRAM',
      turns: [
        { message: 'kkkkk oi', intent: 'LOW_CONFIDENCE' },
        { message: '???', intent: 'LOW_CONFIDENCE' },
      ],
    });
  } catch (err) {
    console.error('  ERRO no TESTE 13:', err.message);
  }

  try {
    results.teste14 = await runPreregisteredHumanOnlyScenario();
  } catch (err) {
    console.error('  ERRO no TESTE 14:', err.message);
  }

  try {
    const teste15Confirmed = await runScenario({
      label: 'TESTE 15 — Agenda, confirma e depois marca NAO COMPARECEU',
      phone: '+5515900000115',
      name: 'TESTE 15 - Vanessa Cardoso',
      source: 'SITE',
      turns: [
        { message: 'Quero agendar uma avaliacao o quanto antes', intent: 'INTENT_07_ALTA_INTENCAO' },
        { message: 'Pode ser esse horario', intent: 'INTENT_07_ALTA_INTENCAO' },
        { message: 'Sim, confirma', intent: 'INTENT_07_ALTA_INTENCAO', confirmReply: true },
      ],
    });
    results.teste15 = await markNoShow(teste15Confirmed);
  } catch (err) {
    console.error('  ERRO no TESTE 15:', err.message);
  }

  console.log('\n\n=== RESUMO FINAL ===');
  for (const [key, lead] of Object.entries(results)) {
    if (!lead) continue;
    console.log(
      `${key}: ${lead.name} | estado=${lead.conversationState || '—'} | status=${lead.leadStatus || '—'} | confirmado=${Boolean(
        lead.confirmedAppointment
      )} | humanOnly=${Boolean(lead.humanOnly)}`
    );
  }
  console.log('\nSimulacao concluida. Nenhum teste anterior (TESTE 1-6) foi apagado ou alterado.');
})().catch((err) => {
  console.error('Erro fatal na simulacao:', err);
  process.exit(1);
});
