/**
 * ============================================================
 * ARILUB — PROCESSO SELETIVO | ASSESSMENT v2.0
 * Vaga: Assistente Comercial — Fortaleza/CE
 * ============================================================
 *
 * Responsabilidades:
 * - Gerar tokens individuais
 * - Montar links individuais
 * - Validar acesso
 * - Registrar início
 * - Salvar respostas progressivamente
 * - Concluir assessment
 * - Bloquear duplicidade
 * - Registrar eventos no LOG
 *
 * IMPORTANTE:
 * A decisão de contratação permanece humana.
 * O tempo de realização é registrado para controle, não para
 * decisão automática.
 * ============================================================
 */

const ARILUB = {
  VERSION: 'v2.0',

  SHEETS: {
    CONFIG: 'CONFIG',
    CANDIDATOS: 'CANDIDATOS',
    AVALIACAO: 'AVALIACAO',
    RESPOSTAS: 'RESPOSTAS',
    ANALISE: 'ANALISE',
    CONTRAPROVAS: 'CONTRAPROVAS',
    ENTREVISTAS: 'ENTREVISTAS',
    LOG: 'LOG'
  },

  STATUS: {
    CONVIDAR: 'CONVIDAR',
    CONVITE_ENVIADO: 'CONVITE ENVIADO',
    INICIADO: 'ASSESSMENT INICIADO',
    CONCLUIDO: 'ASSESSMENT CONCLUÍDO',
    ANALISE: 'ANÁLISE CONCLUÍDA',
    ENTREVISTA: 'ENTREVISTA',
    FINALISTA: 'FINALISTA',
    RESERVA: 'RESERVA',
    ENCERRADO: 'ENCERRADO'
  }
};


/* ============================================================
   1. UTILIDADES
   ============================================================ */

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}


function getSheet_(name) {
  const sheet = getSS_().getSheetByName(name);

  if (!sheet) {
    throw new Error('Aba não encontrada: ' + name);
  }

  return sheet;
}


function getHeaders_(sheet) {
  if (sheet.getLastColumn() === 0) return [];

  return sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(v => String(v).trim());
}


function headerMap_(sheet) {
  const headers = getHeaders_(sheet);
  const map = {};

  headers.forEach((h, i) => {
    if (h) map[h] = i + 1;
  });

  return map;
}


function normalize_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toUpperCase();
}


function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


function now_() {
  return new Date();
}


function isoDate_(date) {
  if (!date) return '';

  return Utilities.formatDate(
    new Date(date),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ss"
  );
}


/* ============================================================
   2. CONFIGURAÇÃO
   ============================================================ */

function getConfig_() {
  const sheet = getSheet_(ARILUB.SHEETS.CONFIG);

  const values = sheet
    .getDataRange()
    .getValues();

  const config = {};

  values.forEach(row => {
    const parametro = String(row[1] || '').trim();

    if (parametro) {
      config[parametro] = row[2];
    }
  });

  return config;
}


function validarConfiguracao() {
  const config = getConfig_();

  const obrigatorios = [
    'TOTAL_QUESTOES',
    'NOME_VAGA',
    'NOME_EMPRESA',
    'VERSAO_ASSESSMENT',
    'STATUS_SISTEMA',
    'BLOQUEAR_DUPLICIDADE'
  ];

  const faltando = obrigatorios.filter(k =>
    config[k] === '' ||
    config[k] === null ||
    typeof config[k] === 'undefined'
  );

  if (faltando.length) {
    throw new Error(
      'Parâmetros ausentes em CONFIG: ' +
      faltando.join(', ')
    );
  }

  return {
    ok: true,
    versao: config.VERSAO_ASSESSMENT,
    status: config.STATUS_SISTEMA,
    totalQuestoes: Number(config.TOTAL_QUESTOES)
  };
}


/* ============================================================
   3. TOKEN
   ============================================================ */

function gerarToken_() {
  return Utilities
    .getUuid()
    .replace(/-/g, '')
    .substring(0, 16)
    .toUpperCase();
}


function gerarTokens() {

  const sheet = getSheet_(ARILUB.SHEETS.CANDIDATOS);
  const map = headerMap_(sheet);

  if (!map.ID_CANDIDATO) {
    throw new Error('Coluna ID_CANDIDATO não encontrada.');
  }

  if (!map.TOKEN) {
    throw new Error('Coluna TOKEN não encontrada.');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'Nenhum candidato encontrado.';
  }

  const tokensExistentes = new Set();

  const tokenValues = sheet
    .getRange(2, map.TOKEN, lastRow - 1, 1)
    .getValues();

  tokenValues.forEach(r => {
    if (r[0]) tokensExistentes.add(String(r[0]));
  });

  let gerados = 0;

  for (let row = 2; row <= lastRow; row++) {

    const id = sheet
      .getRange(row, map.ID_CANDIDATO)
      .getValue();

    if (!id) continue;

    let token = sheet
      .getRange(row, map.TOKEN)
      .getValue();

    if (!token) {

      do {
        token = gerarToken_();
      } while (tokensExistentes.has(token));

      tokensExistentes.add(token);

      sheet
        .getRange(row, map.TOKEN)
        .setValue(token);

      gerados++;

      log_(
        id,
        'TOKEN_GERADO',
        'Token individual criado.',
        '',
        token
      );
    }
  }

  return gerados + ' token(s) gerado(s).';
}


/* ============================================================
   4. LINKS INDIVIDUAIS
   ============================================================ */

function prepararLinks() {

  const config = getConfig_();

  const baseUrl =
    String(config.GITHUB_PAGES_URL || '').trim();

  if (!baseUrl) {
    throw new Error(
      'GITHUB_PAGES_URL ainda não foi preenchida na CONFIG.'
    );
  }

  const sheet = getSheet_(ARILUB.SHEETS.CANDIDATOS);
  const map = headerMap_(sheet);

  if (!map.TOKEN || !map.LINK_AVALIACAO) {
    throw new Error(
      'As colunas TOKEN e LINK_AVALIACAO são obrigatórias.'
    );
  }

  const lastRow = sheet.getLastRow();

  let criados = 0;

  for (let row = 2; row <= lastRow; row++) {

    const token =
      sheet.getRange(row, map.TOKEN).getValue();

    if (!token) continue;

    const link =
      baseUrl.replace(/\/$/, '') +
      '/?t=' +
      encodeURIComponent(token);

    sheet
      .getRange(row, map.LINK_AVALIACAO)
      .setValue(link);

    criados++;
  }

  return criados + ' link(s) preparado(s).';
}


/* ============================================================
   5. LOCALIZAR CANDIDATO
   ============================================================ */

function findCandidateByToken_(token) {

  if (!token) return null;

  const sheet =
    getSheet_(ARILUB.SHEETS.CANDIDATOS);

  const map = headerMap_(sheet);

  if (!map.TOKEN) {
    throw new Error(
      'Coluna TOKEN não encontrada em CANDIDATOS.'
    );
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return null;

  const values =
    sheet.getRange(
      2,
      1,
      lastRow - 1,
      sheet.getLastColumn()
    ).getValues();

  for (let i = 0; i < values.length; i++) {

    const candidateToken =
      String(values[i][map.TOKEN - 1] || '').trim();

    if (candidateToken === String(token).trim()) {

      const obj = {
        row: i + 2
      };

      Object.keys(map).forEach(header => {
        obj[header] =
          values[i][map[header] - 1];
      });

      return obj;
    }
  }

  return null;
}


/* ============================================================
   6. VALIDAÇÃO DE TOKEN
   ============================================================ */

function validarToken(token) {

  try {

    const config = getConfig_();

    const candidato =
      findCandidateByToken_(token);

    if (!candidato) {

      return {
        ok: false,
        codigo: 'TOKEN_INVALIDO',
        mensagem:
          'Este link de participação não é válido.'
      };
    }

    const status =
      normalize_(candidato.STATUS);

    const bloquear =
      normalize_(config.BLOQUEAR_DUPLICIDADE) === 'SIM';

    if (
      bloquear &&
      (
        status === normalize_(ARILUB.STATUS.CONCLUIDO) ||
        status === normalize_(ARILUB.STATUS.ANALISE) ||
        status === normalize_(ARILUB.STATUS.ENTREVISTA) ||
        status === normalize_(ARILUB.STATUS.FINALISTA) ||
        status === normalize_(ARILUB.STATUS.RESERVA) ||
        status === normalize_(ARILUB.STATUS.ENCERRADO)
      )
    ) {

      return {
        ok: false,
        codigo: 'JA_CONCLUIDO',
        mensagem:
          'Sua participação já foi registrada.'
      };
    }

    return {
      ok: true,

      candidato: {
        id: candidato.ID_CANDIDATO,
        nome: candidato.NOME
      },

      vaga: config.NOME_VAGA,
      empresa: config.NOME_EMPRESA,
      introducao: config.TEXTO_INTRO,
      totalQuestoes:
        Number(config.TOTAL_QUESTOES || 30),
      tempoReferencia:
        Number(config.TEMPO_LIMITE_MIN || 20),
      versao:
        config.VERSAO_ASSESSMENT || ARILUB.VERSION,

      status: candidato.STATUS
    };

  } catch (e) {

    return {
      ok: false,
      codigo: 'ERRO_INTERNO',
      mensagem:
        'Não foi possível validar o acesso.'
    };
  }
}


/* ============================================================
   7. CARREGAR QUESTÕES
   ============================================================ */

function carregarQuestoes_() {

  const sheet =
    getSheet_(ARILUB.SHEETS.AVALIACAO);

  const values =
    sheet.getDataRange().getValues();

  if (values.length < 2) return [];

  const headers =
    values[0].map(v => String(v).trim());

  return values
    .slice(1)
    .filter(row => row[0] !== '')
    .map(row => {

      const q = {};

      headers.forEach((h, i) => {
        q[h] = row[i];
      });

      return {
        id: q.ID_QUESTAO,
        bloco: q.BLOCO,
        ordem: q.ORDEM,
        tipo: q.TIPO,
        texto: q.TEXTO_QUESTAO,

        alternativas: {
          A: q.ALT_A || '',
          B: q.ALT_B || '',
          C: q.ALT_C || '',
          D: q.ALT_D || ''
        }
      };
    });
}


/* ============================================================
   8. INICIAR ASSESSMENT
   ============================================================ */

function iniciarAssessment_(token) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(10000);

  try {

    const candidato =
      findCandidateByToken_(token);

    if (!candidato) {
      throw new Error('TOKEN_INVALIDO');
    }

    const sheet =
      getSheet_(ARILUB.SHEETS.CANDIDATOS);

    const map = headerMap_(sheet);

    const status =
      normalize_(candidato.STATUS);

    const concluido =
      status === normalize_(ARILUB.STATUS.CONCLUIDO) ||
      status === normalize_(ARILUB.STATUS.ANALISE) ||
      status === normalize_(ARILUB.STATUS.ENTREVISTA) ||
      status === normalize_(ARILUB.STATUS.FINALISTA) ||
      status === normalize_(ARILUB.STATUS.RESERVA) ||
      status === normalize_(ARILUB.STATUS.ENCERRADO);

    if (concluido) {
      throw new Error('JA_CONCLUIDO');
    }

    if (
      map.DATA_INICIO &&
      !candidato.DATA_INICIO
    ) {

      sheet
        .getRange(
          candidato.row,
          map.DATA_INICIO
        )
        .setValue(now_());
    }

    if (
      status !== normalize_(ARILUB.STATUS.INICIADO)
    ) {

      sheet
        .getRange(
          candidato.row,
          map.STATUS
        )
        .setValue(ARILUB.STATUS.INICIADO);
    }

    if (status !== normalize_(ARILUB.STATUS.INICIADO)) {
      log_(
        candidato.ID_CANDIDATO,
        'ASSESSMENT_INICIADO',
        'Avaliação iniciada.',
        '',
        token
      );
    }

    const candidatoAtualizado = findCandidateByToken_(token);
    const tempo = tempoAssessment_(candidatoAtualizado);

    return {
      ok: true,
      candidato: {
        id: candidato.ID_CANDIDATO,
        nome: candidato.NOME
      },
      questoes: carregarQuestoes_(),
      respostasExistentes:
        getExistingResponses_(candidato.ID_CANDIDATO),
      tempo: tempo
    };

  } finally {
    lock.releaseLock();
  }
}



function tempoAssessment_(candidato) {
  const config = getConfig_();
  const limiteMin = Math.max(1, Number(config.TEMPO_LIMITE_MIN || 25));
  const inicio = candidato.DATA_INICIO ? new Date(candidato.DATA_INICIO) : null;
  const agora = new Date();
  const decorridoSeg = inicio
    ? Math.max(0, Math.floor((agora.getTime() - inicio.getTime()) / 1000))
    : 0;
  const limiteSeg = limiteMin * 60;
  return {
    limiteMin: limiteMin,
    inicioIso: inicio ? inicio.toISOString() : null,
    servidorIso: agora.toISOString(),
    decorridoSeg: decorridoSeg,
    restanteSeg: Math.max(0, limiteSeg - decorridoSeg),
    expirado: !!inicio && decorridoSeg >= limiteSeg
  };
}

function encerrarPorTempo_(token) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const candidato = findCandidateByToken_(token);
    if (!candidato) throw new Error('TOKEN_INVALIDO');

    const tempo = tempoAssessment_(candidato);
    if (!tempo.expirado) {
      return {ok:false, codigo:'TEMPO_AINDA_ATIVO', restanteSeg:tempo.restanteSeg};
    }

    const sheet = getSheet_(ARILUB.SHEETS.CANDIDATOS);
    const map = headerMap_(sheet);
    const fim = now_();
    const totalRespondido = contarRespostas_(candidato.ID_CANDIDATO);

    if (map.DATA_FIM) sheet.getRange(candidato.row, map.DATA_FIM).setValue(fim);
    if (map.TEMPO_TOTAL_MIN) sheet.getRange(candidato.row, map.TEMPO_TOTAL_MIN).setValue(tempo.limiteMin);
    sheet.getRange(candidato.row, map.STATUS).setValue(ARILUB.STATUS.CONCLUIDO);

    log_(candidato.ID_CANDIDATO,'ASSESSMENT_TEMPO_ESGOTADO',
      'Tempo esgotado. Respostas registradas: ' + totalRespondido + '.', '', token);

    return {
      ok:true,
      expirado:true,
      respondidas:totalRespondido,
      mensagem:'Tempo encerrado. As respostas já registradas foram preservadas.'
    };
  } finally {
    lock.releaseLock();
  }
}


/* ============================================================
   9. RESPOSTAS EXISTENTES
   ============================================================ */

function getExistingResponses_(idCandidato) {

  const sheet =
    getSheet_(ARILUB.SHEETS.RESPOSTAS);

  const map = headerMap_(sheet);
  const result = {};
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return result;

  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        sheet.getLastColumn()
      )
      .getValues();

  data.forEach((r, index) => {

    const id =
      String(r[map.ID_CANDIDATO - 1] || '').trim();

    if (id !== String(idCandidato).trim()) return;

    const q =
      String(r[map.ID_QUESTAO - 1] || '').trim();

    if (!q) return;

    result[q] = {
      row: index + 2,
      resposta:
        map.RESPOSTA_BRUTA
          ? r[map.RESPOSTA_BRUTA - 1]
          : '',
      alternativa:
        map.VALOR_ALTERNATIVA
          ? r[map.VALOR_ALTERNATIVA - 1]
          : '',
      tipo:
        map.TIPO
          ? r[map.TIPO - 1]
          : '',
      bloco:
        map.BLOCO
          ? r[map.BLOCO - 1]
          : ''
    };
  });

  return result;
}


/* ============================================================
   10. SALVAR BLOCO
   ============================================================ */

function validarRespostasBloco_(bloco, respostas) {

  if (!Array.isArray(respostas)) {
    throw new Error('RESPOSTAS_INVALIDAS');
  }

  const questoesBloco =
    carregarQuestoes_()
      .filter(q =>
        normalize_(q.bloco) === normalize_(bloco)
      );

  if (!questoesBloco.length) {
    throw new Error('BLOCO_INVALIDO');
  }

  const porId = {};

  respostas.forEach(resp => {
    const id = String(resp.idQuestao || '').trim();

    if (!id) {
      throw new Error('QUESTAO_INVALIDA');
    }

    if (porId[id]) {
      throw new Error('QUESTAO_DUPLICADA_' + id);
    }

    porId[id] = resp;
  });

  questoesBloco.forEach(q => {

    const id = String(q.id);
    const resp = porId[id];

    if (!resp) {
      throw new Error('QUESTAO_NAO_RESPONDIDA_' + id);
    }

    const resposta =
      String(
        resp.resposta == null ? '' : resp.resposta
      ).trim();

    if (!resposta) {
      throw new Error('QUESTAO_NAO_RESPONDIDA_' + id);
    }

    const tipo =
      normalize_(q.tipo);

    if (tipo.indexOf('MULTIPLA') !== -1) {

      const alternativa =
        String(resp.alternativa || '')
          .trim()
          .toUpperCase();

      const permitidas =
        Object.keys(q.alternativas || {})
          .filter(letra =>
            String(q.alternativas[letra] || '').trim() !== ''
          );

      if (!permitidas.includes(alternativa)) {
        throw new Error('ALTERNATIVA_INVALIDA_' + id);
      }
    }
  });

  if (Object.keys(porId).length !== questoesBloco.length) {
    throw new Error('QUANTIDADE_RESPOSTAS_INVALIDA');
  }

  return true;
}


function salvarBloco_(token, bloco, respostas) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(15000);

  try {

    const candidato =
      findCandidateByToken_(token);

    if (!candidato) {
      throw new Error('TOKEN_INVALIDO');
    }

    const validacao =
      validarToken(token);

    if (!validacao.ok) {
      throw new Error(validacao.codigo);
    }

    const tempo = tempoAssessment_(candidato);
    if (tempo.expirado) {
      return {
        ok: false,
        codigo: 'TEMPO_ESGOTADO',
        mensagem: 'O tempo da avaliação foi encerrado.'
      };
    }

    validarRespostasBloco_(bloco, respostas);

    const sheet =
      getSheet_(ARILUB.SHEETS.RESPOSTAS);

    const map = headerMap_(sheet);

    const existentes =
      getExistingResponses_(
        candidato.ID_CANDIDATO
      );

    const config = getConfig_();

    respostas.forEach(resp => {

      const idQuestao =
        String(resp.idQuestao);

      const respostaBruta =
        resp.resposta == null
          ? ''
          : String(resp.resposta).trim();

      if (!respostaBruta) return;

      const valorAlternativa =
        resp.alternativa
          ? String(resp.alternativa).trim()
          : '';

      const linha = [
        candidato.ID_CANDIDATO,
        candidato.NOME,
        resp.idQuestao,
        bloco,
        resp.tipo || '',
        respostaBruta,
        valorAlternativa,
        now_(),
        config.VERSAO_ASSESSMENT || ARILUB.VERSION,
        token,
        bloco
      ];

      if (existentes[idQuestao]) {

        sheet
          .getRange(
            existentes[idQuestao].row,
            1,
            1,
            linha.length
          )
          .setValues([linha]);

      } else {

        sheet.appendRow(linha);
      }
    });

    log_(
      candidato.ID_CANDIDATO,
      'BLOCO_SALVO',
      'Bloco salvo: ' + bloco,
      '',
      token
    );

    return {
      ok: true,
      bloco: bloco,
      mensagem:
        'Respostas salvas com sucesso.'
    };

  } finally {
    lock.releaseLock();
  }
}


/* ============================================================
   11. CONTAR RESPOSTAS
   ============================================================ */

function contarRespostas_(idCandidato) {

  const sheet =
    getSheet_(ARILUB.SHEETS.RESPOSTAS);

  const map = headerMap_(sheet);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return 0;

  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        sheet.getLastColumn()
      )
      .getValues();

  const questoes = new Set();

  data.forEach(r => {

    if (
      String(r[map.ID_CANDIDATO - 1]) ===
      String(idCandidato)
    ) {

      questoes.add(
        String(r[map.ID_QUESTAO - 1])
      );
    }
  });

  return questoes.size;
}


/* ============================================================
   12. CONCLUIR ASSESSMENT
   ============================================================ */

function concluirAssessment_(token) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(15000);

  try {

    const candidato =
      findCandidateByToken_(token);

    if (!candidato) {
      throw new Error('TOKEN_INVALIDO');
    }

    const config = getConfig_();
    const tempo = tempoAssessment_(candidato);

    if (tempo.expirado) {
      return {
        ok: false,
        codigo: 'TEMPO_ESGOTADO',
        mensagem: 'O tempo da avaliação foi encerrado.'
      };
    }

    const totalEsperado =
      Number(
        config.TOTAL_QUESTOES || 30
      );

    const totalRespondido =
      contarRespostas_(
        candidato.ID_CANDIDATO
      );

    if (totalRespondido < totalEsperado) {

      return {
        ok: false,
        codigo: 'INCOMPLETO',
        respondidas: totalRespondido,
        total: totalEsperado,
        mensagem:
          'Ainda existem questões sem resposta.'
      };
    }

    const sheet =
      getSheet_(ARILUB.SHEETS.CANDIDATOS);

    const map = headerMap_(sheet);

    const fim = now_();

    let inicio =
      candidato.DATA_INICIO;

    if (!inicio) {
      inicio = fim;
    }

    const minutos =
      Math.max(
        0,
        Math.round(
          (
            new Date(fim).getTime() -
            new Date(inicio).getTime()
          ) / 60000
        )
      );

    if (map.DATA_FIM) {
      sheet
        .getRange(
          candidato.row,
          map.DATA_FIM
        )
        .setValue(fim);
    }

    if (map.TEMPO_TOTAL_MIN) {
      sheet
        .getRange(
          candidato.row,
          map.TEMPO_TOTAL_MIN
        )
        .setValue(minutos);
    }

    sheet
      .getRange(
        candidato.row,
        map.STATUS
      )
      .setValue(
        ARILUB.STATUS.CONCLUIDO
      );

    log_(
      candidato.ID_CANDIDATO,
      'ASSESSMENT_CONCLUIDO',
      'Assessment concluído com ' +
        totalRespondido +
        ' respostas.',
      '',
      token
    );

    return {
      ok: true,
      mensagem:
        'Sua participação foi registrada com sucesso.',
      respondidas: totalRespondido
    };

  } finally {
    lock.releaseLock();
  }
}


/* ============================================================
   13. LOG
   ============================================================ */

function log_(
  idCandidato,
  acao,
  detalhe,
  erro,
  token
) {

  try {

    const sheet =
      getSheet_(ARILUB.SHEETS.LOG);

    const config = getConfig_();

    sheet.appendRow([
      now_(),
      idCandidato || '',
      acao || '',
      detalhe || '',
      erro || '',
      token || '',
      config.VERSAO_ASSESSMENT ||
        ARILUB.VERSION
    ]);

  } catch (e) {
    console.error(
      'Falha ao registrar LOG:',
      e
    );
  }
}


/* ============================================================
   14. API — GET
   ============================================================ */

function doGet(e) {

  try {

    const action =
      String(
        e.parameter.action || 'validate'
      ).trim();

    const token =
      String(
        e.parameter.t ||
        e.parameter.token ||
        ''
      ).trim();

    if (!token) {
      return json_({
        ok: false,
        codigo: 'TOKEN_AUSENTE',
        mensagem:
          'Token não informado.'
      });
    }

    if (action === 'validate') {
      return json_(
        validarToken(token)
      );
    }

    if (action === 'start') {
      return json_(
        iniciarAssessment_(token)
      );
    }

    return json_({
      ok: false,
      codigo: 'ACAO_INVALIDA'
    });

  } catch (e) {

    return json_({
      ok: false,
      codigo: 'ERRO_INTERNO',
      mensagem:
        'Não foi possível processar a solicitação.'
    });
  }
}


/* ============================================================
   15. API — POST
   ============================================================ */

function doPost(e) {

  let payload = {};

  try {

    payload =
      JSON.parse(
        e.postData.contents || '{}'
      );

    const action =
      String(
        payload.action || ''
      ).trim();

    const token =
      String(
        payload.token || ''
      ).trim();

    if (!token) {
      return json_({
        ok: false,
        codigo: 'TOKEN_AUSENTE'
      });
    }

    if (action === 'saveBlock') {

      return json_(
        salvarBloco_(
          token,
          payload.bloco,
          payload.respostas || []
        )
      );
    }

    if (action === 'finish') {

      return json_(
        concluirAssessment_(token)
      );
    }

    if (action === 'timeout') {

      return json_(
        encerrarPorTempo_(token)
      );
    }

    return json_({
      ok: false,
      codigo: 'ACAO_INVALIDA'
    });

  } catch (e) {

    try {

      log_(
        '',
        'ERRO_API',
        payload.action || '',
        e.message || String(e),
        payload.token || ''
      );

    } catch (_) {}

    return json_({
      ok: false,
      codigo: 'ERRO_INTERNO',
      mensagem:
        'Não foi possível processar a solicitação.'
    });
  }
}


/* ============================================================
   16. PREPARAÇÃO DO SISTEMA
   ============================================================ */

function prepararSistema() {

  validarConfiguracao();

  const resultadoTokens =
    gerarTokens();

  let resultadoLinks =
    'Links ainda não gerados.';

  const config = getConfig_();

  if (
    String(
      config.GITHUB_PAGES_URL || ''
    ).trim()
  ) {

    resultadoLinks =
      prepararLinks();
  }

  return {
    ok: true,
    tokens: resultadoTokens,
    links: resultadoLinks
  };
}


/* ============================================================
   17. TESTE INTERNO
   ============================================================ */

function testeSistema() {

  const resultado =
    validarConfiguracao();

  const questoes =
    carregarQuestoes_();

  const sheet =
    getSheet_(ARILUB.SHEETS.CANDIDATOS);

  const map =
    headerMap_(sheet);

  return {
    configuracao: resultado,
    totalQuestoesCarregadas:
      questoes.length,
    possuiToken:
      Boolean(map.TOKEN),
    possuiLink:
      Boolean(map.LINK_AVALIACAO),
    possuiStatus:
      Boolean(map.STATUS),
    versao:
      ARILUB.VERSION
  };
}

/**
 * ============================================================
 * ARILUB — GERAR TOKEN EXCLUSIVAMENTE PARA CANDIDATO DE TESTE
 * ============================================================
 *
 * Atua somente sobre:
 * ID_CANDIDATO = TESTE-0001
 *
 * Não altera os candidatos reais.
 * Se TESTE-0001 já possuir token, mantém o token existente.
 * ============================================================
 */

function gerarTokenTeste() {

  const ID_TESTE = 'TESTE-0001';

  const sheet =
    getSheet_(ARILUB.SHEETS.CANDIDATOS);

  const map =
    headerMap_(sheet);

  // Validação das colunas necessárias
  if (!map.ID_CANDIDATO) {
    throw new Error(
      'Coluna ID_CANDIDATO não encontrada.'
    );
  }

  if (!map.TOKEN) {
    throw new Error(
      'Coluna TOKEN não encontrada.'
    );
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error(
      'Nenhum candidato encontrado.'
    );
  }

  // Procura SOMENTE o candidato TESTE-0001
  const ids =
    sheet
      .getRange(
        2,
        map.ID_CANDIDATO,
        lastRow - 1,
        1
      )
      .getValues();

  let linhaTeste = null;

  for (let i = 0; i < ids.length; i++) {

    const id =
      String(ids[i][0] || '').trim();

    if (id === ID_TESTE) {

      linhaTeste = i + 2;
      break;
    }
  }

  if (!linhaTeste) {
    throw new Error(
      'Candidato técnico TESTE-0001 não encontrado.'
    );
  }

  // Verifica se já existe token
  const tokenAtual =
    String(
      sheet
        .getRange(
          linhaTeste,
          map.TOKEN
        )
        .getValue() || ''
    ).trim();

  if (tokenAtual) {

    return {
      ok: true,
      id: ID_TESTE,
      token: tokenAtual,
      novoToken: false,
      mensagem:
        'TESTE-0001 já possuía token. Nenhuma alteração foi realizada.'
    };
  }

  // Gera token
  const novoToken =
    gerarToken_();

  sheet
    .getRange(
      linhaTeste,
      map.TOKEN
    )
    .setValue(novoToken);

  // Garante a versão do assessment, caso a coluna exista
  if (map.VERSAO_ASSESSMENT) {

    sheet
      .getRange(
        linhaTeste,
        map.VERSAO_ASSESSMENT
      )
      .setValue(ARILUB.VERSION);
  }

  // Registra no LOG
  log_(
    ID_TESTE,
    'TOKEN_TESTE_GERADO',
    'Token criado exclusivamente para o candidato técnico TESTE-0001.',
    '',
    novoToken
  );

  return {
    ok: true,
    id: ID_TESTE,
    token: novoToken,
    novoToken: true,
    mensagem:
      'Token técnico gerado com sucesso.'
  };
}

function testeSalvarBloco1() {

  const token = '69FE84AE345548BE';

  const respostas = [
    {
      idQuestao: 1,
      tipo: 'multipla',
      resposta: 'B',
      alternativa: 'B'
    },
    {
      idQuestao: 2,
      tipo: 'multipla',
      resposta: 'C',
      alternativa: 'C'
    },
    {
      idQuestao: 3,
      tipo: 'aberta',
      resposta: 'Organizaria os retornos no CRM por data e prioridade, deixando programada a próxima ação.'
    },
    {
      idQuestao: 4,
      tipo: 'multipla',
      resposta: 'B',
      alternativa: 'B'
    },
    {
      idQuestao: 5,
      tipo: 'multipla',
      resposta: 'C',
      alternativa: 'C'
    },
    {
      idQuestao: 6,
      tipo: 'aberta',
      resposta: 'Reviso a atividade, verifico pendências e confirmo se os registros e próximos passos foram realizados.'
    },
    {
      idQuestao: 7,
      tipo: 'multipla',
      resposta: 'C',
      alternativa: 'C'
    },
    {
      idQuestao: 8,
      tipo: 'aberta',
      resposta: 'Primeiro trataria o retorno prometido para ontem, depois a proposta que vence hoje e por último a pesquisa sem prazo, registrando os próximos passos.'
    },
    {
      idQuestao: 9,
      tipo: 'multipla',
      resposta: 'B',
      alternativa: 'B'
    },
    {
      idQuestao: 10,
      tipo: 'aberta',
      resposta: 'Organizo as atividades por prazo e prioridade e acompanho durante o dia o que foi concluído e o que continua pendente.'
    }
  ];

  const resultado = salvarBloco_(
    token,
    'BLOCO 1',
    respostas
  );

  Logger.log(
    JSON.stringify(resultado, null, 2)
  );

  return resultado;
}

function testeFecharAssessment() {

  const token = '69FE84AE345548BE';

  // =========================
  // BLOCO 2 — Questões 11–20
  // =========================

  const bloco2 = [
    { idQuestao: 11, tipo: 'multipla', resposta: 'B', alternativa: 'B' },

    {
      idQuestao: 12,
      tipo: 'aberta',
      resposta: 'Informaria ao cliente que estou aguardando uma informação interna e combinaria um prazo objetivo para retornar.'
    },

    { idQuestao: 13, tipo: 'multipla', resposta: 'C', alternativa: 'C' },

    {
      idQuestao: 14,
      tipo: 'aberta',
      resposta: 'Definiria critérios de potencial, pesquisaria as empresas em fontes confiáveis e organizaria as informações para priorizar os contatos.'
    },

    { idQuestao: 15, tipo: 'multipla', resposta: 'B', alternativa: 'B' },

    {
      idQuestao: 16,
      tipo: 'aberta',
      resposta: 'Registraria contexto, necessidade do cliente, informações relevantes, compromissos assumidos e o próximo passo.'
    },

    {
      idQuestao: 17,
      tipo: 'aberta',
      resposta: 'Consultaria primeiro os materiais disponíveis e fontes confiáveis. Se necessário, confirmaria a informação com alguém da equipe antes de responder.'
    },

    { idQuestao: 18, tipo: 'multipla', resposta: 'C', alternativa: 'C' },

    {
      idQuestao: 19,
      tipo: 'aberta',
      resposta: 'Identificaria a situação mais crítica, daria retorno inicial aos dois clientes e priorizaria o problema que pode comprometer a negociação.'
    },

    {
      idQuestao: 20,
      tipo: 'aberta',
      resposta: 'Um bom atendimento entende a necessidade, comunica com clareza, cumpre os combinados e acompanha o cliente até o próximo passo.'
    }
  ];


  // ========================
  // BLOCO 3 — Questões 21–30
  // =========================

  const bloco3 = [
    { idQuestao: 21, tipo: 'multipla', resposta: 'C', alternativa: 'C' },

    {
      idQuestao: 22,
      tipo: 'aberta',
      resposta: 'Verificaria o histórico, último contato, estágio da oportunidade, compromissos existentes e a próxima ação prevista.'
    },

    { idQuestao: 23, tipo: 'multipla', resposta: 'B', alternativa: 'B' },

    {
      idQuestao: 24,
      tipo: 'aberta',
      resposta: 'Buscaria entender a orientação, consultaria materiais disponíveis, praticaria na ferramenta e tiraria dúvidas específicas.'
    },

    {
      idQuestao: 25,
      tipo: 'aberta',
      resposta: 'Revisaria a causa do erro, confirmaria o procedimento correto e criaria uma forma de conferência para evitar a repetição.'
    },

    {
      idQuestao: 26,
      tipo: 'aberta',
      resposta: 'Utilizaria site da empresa, Google, redes e fontes institucionais ou comerciais confiáveis, organizando as informações encontradas.'
    },

    { idQuestao: 27, tipo: 'multipla', resposta: 'C', alternativa: 'C' },

    { idQuestao: 28, tipo: 'multipla', resposta: 'B', alternativa: 'B' },

    {
      idQuestao: 29,
      tipo: 'aberta',
      resposta: 'Cliente demonstrou interesse. Aguarda informação adicional antes da decisão. Retorno combinado amanhã às 10h. Criar tarefa de retorno para 10h.'
    },

    {
      idQuestao: 30,
      tipo: 'aberta',
      resposta: 'Nas primeiras semanas organizaria meu aprendizado por processos, ferramentas e produtos, praticaria as atividades, registraria dúvidas e buscaria feedback até conseguir executar com autonomia.'
    }
  ];


  // =========================
  // SALVAMENTO
  // =========================

  const resultadoBloco2 =
    salvarBloco_(
      token,
      'BLOCO 2',
      bloco2
    );

  const resultadoBloco3 =
    salvarBloco_(
      token,
      'BLOCO 3',
      bloco3
    );


  // =========================
  // CONFERÊNCIA
  // =========================

  const candidato =
    findCandidateByToken_(token);

  if (!candidato) {
    throw new Error(
      'TESTE-0001 não localizado.'
    );
  }

  const total =
    contarRespostas_(
      candidato.ID_CANDIDATO
    );

  if (total !== 30) {

    throw new Error(
      'Esperadas 30 respostas únicas, mas foram encontradas ' +
      total +
      '. Assessment NÃO foi concluído.'
    );
  }


  // =========================
  // CONCLUSÃO
  // =========================

  const conclusao =
    concluirAssessment_(token);


  // =========================
  // RESULTADO
  // =========================

  const resultado = {
    ok: conclusao.ok === true,
    bloco2: resultadoBloco2,
    bloco3: resultadoBloco3,
    respostasUnicas: total,
    conclusao: conclusao
  };

  Logger.log(
    JSON.stringify(
      resultado,
      null,
      2
    )
  );

  return resultado;
}

function resetarCandidatoTeste() {

  const ID_TESTE = 'TESTE-0001';

  const ss = getSS_();
  const shCand = getSheet_(ARILUB.SHEETS.CANDIDATOS);
  const shResp = getSheet_(ARILUB.SHEETS.RESPOSTAS);

  // =========================
  // 1. LOCALIZA TESTE-0001
  // =========================

  const dados = shCand.getDataRange().getValues();
  const headers = dados[0];
  const map = {};

  headers.forEach((h, i) => {
    map[String(h).trim()] = i;
  });

  let linha = -1;

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][map.ID_CANDIDATO]).trim() === ID_TESTE) {
      linha = i + 1;
      break;
    }
  }

  if (linha === -1) {
    throw new Error('TESTE-0001 não encontrado.');
  }


  // =========================
  // 2. APAGA SOMENTE RESPOSTAS
  //    DO TESTE-0001
  // =========================

  const respostas = shResp.getDataRange().getValues();

  if (respostas.length > 1) {

    const hResp = respostas[0];
    const idxID = hResp.indexOf('ID_CANDIDATO');

    for (let i = respostas.length - 1; i >= 1; i--) {

      if (
        String(respostas[i][idxID]).trim() === ID_TESTE
      ) {
        shResp.deleteRow(i + 1);
      }
    }
  }


  // =========================
  // 3. RESETA DADOS OPERACIONAIS
  // =========================

  const limpar = [
    'LINK_AVALIACAO',
    'DATA_ENVIO',
    'DATA_INICIO',
    'DATA_FIM',
    'TEMPO_TOTAL_MIN',
    'TOKEN'
  ];

  limpar.forEach(campo => {

    if (map[campo] !== undefined) {
      shCand
        .getRange(linha, map[campo] + 1)
        .clearContent();
    }

  });


  // =========================
  // 4. VOLTA STATUS
  // =========================

  shCand
    .getRange(linha, map.STATUS + 1)
    .setValue('CONVIDAR');


  // =========================
  // 5. GARANTE VERSÃO
  // =========================

  if (map.VERSAO_ASSESSMENT !== undefined) {

    shCand
      .getRange(
        linha,
        map.VERSAO_ASSESSMENT + 1
      )
      .setValue('v2.0');
  }


  // =========================
  // 6. GERA NOVO TOKEN
  // =========================

  const novoToken = gerarToken_();

  shCand
    .getRange(linha, map.TOKEN + 1)
    .setValue(novoToken);


  // =========================
  // 7. LOG
  // =========================

  log_(
    ID_TESTE,
    'RESET_TESTE_INTERFACE',
    'Candidato técnico resetado para teste completo via GitHub Pages.',
    '',
    novoToken
  );


  // =========================
  // RESULTADO
  // =========================

  const resultado = {
    ok: true,
    candidato: ID_TESTE,
    status: 'CONVIDAR',
    token: novoToken,
    mensagem:
      'TESTE-0001 resetado. Utilize o novo token no GitHub Pages.'
  };

  Logger.log(
    JSON.stringify(resultado, null, 2)
  );

  return resultado;
}