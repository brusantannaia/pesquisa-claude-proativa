/**
 * Google Apps Script — Webhook da Pesquisa Claude na Proativa
 *
 * Como configurar:
 * 1. Crie uma planilha vazia no Google Sheets ("Respostas - Pesquisa Claude Proativa")
 * 2. Copie o ID da planilha da URL (entre /d/ e /edit) e cole em SHEET_ID abaixo
 * 3. Em script.google.com → Novo projeto → cole este código
 * 4. Salve, rode a função `testar` uma vez (autorize o acesso)
 * 5. Confira que apareceu uma linha de teste na planilha
 * 6. Implantar → Nova implantação → Tipo: App da Web
 *    - Executar como: Eu
 *    - Quem tem acesso: Qualquer pessoa
 * 7. Copie a URL gerada e cole no index.html (variável GOOGLE_SCRIPT_URL)
 */

// ID da planilha já preenchido (planilha criada automaticamente em 2026-05-26)
// https://docs.google.com/spreadsheets/d/16JHBg1gmH0YHlfcBtbF_n3YX8UFZNJZRqtQzWN4QwmU/edit
const SHEET_ID = '16JHBg1gmH0YHlfcBtbF_n3YX8UFZNJZRqtQzWN4QwmU';
const SHEET_NAME = 'Respostas';

// Cabeçalhos da planilha (ordem das colunas)
const HEADERS = [
  'Timestamp',
  'Nome',
  'Q2 — Nível de IA',
  'Q3 — Experiência com Claude',
  'Q4 — Ferramentas usadas',
  'Q4a — Outra ferramenta',
  'Q5 — Confiança em identificar onde IA se aplica',
  'Q6 — Compreendo limites da IA',
  'Q7 — Avaliar confiabilidade de resposta',
  'Q8 — Contextos de uso na Proativa',
  'Q8a — Outro contexto',
  'Q9 — Postura emocional',
  'Q10 — O que quer aprender',
  'Q11 — Dúvidas, receios, temas'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Adiciona cabeçalhos se planilha estiver vazia
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#005847')
        .setFontColor('#ffffff')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle');
      sheet.setRowHeight(1, 32);
      sheet.setFrozenRows(1);
      // Larguras sugeridas
      const widths = [160, 180, 220, 260, 280, 200, 240, 220, 240, 280, 200, 240, 280, 280];
      widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    }

    const row = [
      data.timestamp || new Date().toISOString(),
      data.q1_nome || '',
      data.q2_nivel || '',
      data.q3_claude || '',
      data.q4_ferramentas || '',
      data.q4_outro || '',
      data.q5_onde_aplicar || '',
      data.q6_limites || '',
      data.q7_confiabilidade || '',
      data.q8_contexto || '',
      data.q8_outro || '',
      data.q9_postura || '',
      data.q10_aprender || '',
      data.q11_duvidas || ''
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error('Erro ao processar resposta:', err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Rode esta função uma vez para autorizar o script
 * e validar que escreve corretamente na planilha.
 */
function testar() {
  const e = {
    postData: {
      contents: JSON.stringify({
        timestamp: new Date().toISOString(),
        q1_nome: 'Teste — Apagar depois',
        q2_nivel: '5 - Uso avançado, exploro recursos e construo fluxos',
        q3_claude: 'Uso intensamente (Projects, artifacts, MCPs, skills...)',
        q4_ferramentas: 'Claude (Anthropic); ChatGPT (OpenAI); Gemini (Google)',
        q4_outro: '',
        q5_onde_aplicar: '5 - Confiança total',
        q6_limites: '5 - Concordo totalmente',
        q7_confiabilidade: '5 - Concordo totalmente',
        q8_contexto: 'Análise de dados e relatórios; Tomada de decisão estratégica',
        q8_outro: '',
        q9_postura: 'Animado(a) com as possibilidades',
        q10_aprender: 'Validar que o webhook funciona end to end.',
        q11_duvidas: ''
      })
    }
  };
  const result = doPost(e);
  Logger.log(result.getContent());
}
