/**
 * ============================================================================
 *  BM ADVOCACIA — Motor de sincronizacao  DataJud (CNJ) + DJEN  ->  Planilha
 *  Planilha: "Processos [DataJud + DJEN]" - aba "Base geral"
 * ----------------------------------------------------------------------------
 *  O QUE FAZ
 *   1. Monta a base de processos (planilha-mae "Clientes e Processos" +
 *      processos descobertos no DJEN pelas OABs do escritorio).
 *   2. Consulta a API Publica do DataJud (CNJ) em lote, por tribunal, e grava
 *      classe, assunto, orgao julgador, grau, sistema, formato, sigilo,
 *      data de ajuizamento e TODAS as movimentacoes.
 *   3. Varre o DJEN (Diario de Justica Eletronico Nacional) pelas OABs do
 *      escritorio e grava TODAS as publicacoes de cada processo.
 *   4. Calcula colunas de gestao (fase, dias parado, situacao, ultimos eventos).
 *
 *  QUANDO RODA
 *   - Todo dia as 6h (gatilho de tempo criado por configurarTudo()).
 *   - Sempre que o botao "Atualizar" do dashboard chamar o Web App (doGet).
 *   - Pelo menu "BM - DataJud/DJEN" dentro da propria planilha.
 *
 *  Execucoes longas sao retomadas automaticamente: o script salva um cursor e
 *  cria um gatilho de continuacao antes de estourar o limite de tempo.
 * ============================================================================
 */

/* --------------------------- CONFIGURACAO ------------------------------- */

var CFG_PADRAO = {
  /* ID desta planilha. So e usado quando o script roda como projeto
     independente (nao vinculado a planilha) - ai getActive() volta null. */
  PLANILHA_ID:         '1MXpGlo3U1bDV34l_jj98ZNPwgguvD-hCFOCaARNQaYs',
  ABA_BASE:            'Base geral',
  PLANILHA_MAE_ID:     '1XKMeYEapBqBq_IIaLu2uN-ceB3btArIYmrPBuyxsLHU',
  PLANILHA_MAE_ABA:    'Todos os Processos',
  INCLUIR_PLANILHA_MAE:'SEMPRE',
  /* OABs varridas no DJEN. Aceita "41438/BA", "BA/41438" ou "OAB/BA 41438". */
  OABS:                '41438/BA, 271081/RJ, 536843/SP, 63805/BA',
  DJEN_DATA_INICIAL:   '2023-01-01',
  DJEN_JANELA_DIAS:    '45',
  DATAJUD_APIKEY:      'APIKey cDZHYzlZa0JadVREZDJCendQbXY6SkJlTzNjLV9TRENyQk1RdnFKZGRQdw==',
  DESCOBRIR_NO_DJEN:   'SIM',
  MAX_MOVIMENTOS:      '400',
  MAX_PUBLICACOES:     '200',
  TOKEN_WEBAPP:        'bm-datajud',
  DIAS_ESTAGNADO:      '90'
};

var ABA_SYNC = '_Sync';
var ABA_CFG  = '_Config';
var ABA_DJEN = '_DJEN';

var LIMITE_MS        = 4.5 * 60 * 1000;   // recomeca antes do teto de 6 min
var LOTE_DATAJUD     = 40;                // processos por requisicao
var PARALELO_DATAJUD = 8;                 // requisicoes simultaneas (fetchAll)
var DJEN_ITENS_PAG   = 500;
var DJEN_PAUSA_MS    = 3300;              // limite observado: 20 req/min
var MAX_CHARS_CEL    = 45000;             // teto seguro por celula (limite 50k)

/* Os nomes abaixo sao reescritos na linha 1 da Base geral. Mantidos exatamente
   como estao na planilha (as 15 primeiras) para nao desfigurar o cabecalho. */
var COLUNAS = [
  'Tipo de Processo','Número do Processo','Partes','Tribunal','Vara','Grau','Classe',
  'Assunto','Órgão Julgador','Data do Ajuizamento','Sistema','Formato','Nível de Sigilo',
  'Movimentos','Publicações no DJEN',
  /* gestão */
  'Última Atualização','Qtd. Movimentos','Data da Última Movimentação','Última Movimentação',
  'Dias sem Movimentação','Qtd. Publicações','Data da Última Publicação','Última Publicação',
  'Fase Processual','Situação','Valor da Causa','Cliente / Parte Representada','Origem do Cadastro',
  /* espelho da planilha-mãe — classificação do escritório */
  'Resultado da Sentença','Resultado do Recurso','Fase Processual (cadastro)','Situação do Alvará',
  /* duas dimensoes independentes, gravadas para valerem tambem fora do painel */
  'Status do Processo','Status de Movimentação'
];
var N_COLUNAS = 34;

/* --------------------------- MENU / SETUP ------------------------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BM - DataJud/DJEN')
    .addItem('Sincronizar agora (completa)', 'sincronizarCompletaMenu')
    .addItem('Sincronizar incremental', 'sincronizarIncrementalMenu')
    .addSeparator()
    .addItem('Importar processos da planilha-mae', 'importarDaPlanilhaMae')
    .addItem('Incluir processos por numero (colar lista)', 'incluirProcessosPorNumero')
    .addItem('Buscar processos novos no diario (DJEN)', 'buscarProcessosNoDiario')
    .addItem('Aplicar OABs padrao no _Config', 'aplicarOABsPadrao')
    .addItem('Configurar tudo (1a vez)', 'configurarTudo')
    .addItem('Recriar gatilho das 6h', 'criarGatilhoDiario')
    .addItem('Cancelar sincronizacao', 'cancelarSincronizacao')
    .addToUi();
}

/** Cria abas de apoio, cabecalhos e o gatilho diario das 6h. */
function configurarTudo() {
  var ss = planilha_();
  garantirConfig_(ss);
  garantirSync_(ss);
  garantirDJEN_(ss);
  garantirCabecalho_(ss);
  criarGatilhoDiario();
  try {
    SpreadsheetApp.getUi().alert(
      'Configuracao concluida.\n\n' +
      '- Abas _Config, _Sync e _DJEN criadas.\n' +
      '- Cabecalho da "Base geral" ajustado (' + N_COLUNAS + ' colunas).\n' +
      '- Gatilho diario as 6h criado.\n\n' +
      'Agora publique o Web App (Implantar > Nova implantacao > App da Web, ' +
      'executar como Eu, acesso Qualquer pessoa) e cole a URL no dashboard.');
  } catch (e) {}
}

function criarGatilhoDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sincronizacaoDiaria') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sincronizacaoDiaria')
    .timeBased().atHour(6).nearMinute(0).everyDays(1)
    .inTimezone(Session.getScriptTimeZone() || 'America/Bahia').create();
}

/* ------------------------- PONTOS DE ENTRADA ---------------------------- */

function sincronizacaoDiaria() {
  /* Antes de sincronizar, procura no diario processos que a base ainda nao
     tem. Falha aqui nao pode derrubar a sincronizacao: se o DJEN recusar a
     chamada, segue-se com a base como esta. */
  try { semearDoDJEN_(30); } catch (e) {}
  iniciarSincronizacao_('completa', 'gatilho 6h');
  executarEtapas_();
}
function sincronizarCompletaMenu()   { iniciarSincronizacao_('completa', 'menu');        executarEtapas_(); }
function sincronizarIncrementalMenu(){ iniciarSincronizacao_('incremental', 'menu');     executarEtapas_(); }
function continuarSincronizacao()    { limparGatilhosContinuacao_();                     executarEtapas_(); }

/**
 * Traz para a Base geral todo processo de "Clientes e Processos" que ainda nao
 * esteja aqui, sem esperar a proxima sincronizacao. Nao remove nada.
 */
function importarDaPlanilhaMae() {
  var ss = planilha_();
  garantirConfig_(ss); garantirSync_(ss); garantirDJEN_(ss); garantirCabecalho_(ss);
  var aba = ss.getSheetByName(cfg_('ABA_BASE'));
  var antes = Math.max(0, aba.getLastRow() - 1);
  props_().setProperty('forcarSemear', '1');
  try {
    etapaBase_({ modo: 'completa' });
  } finally {
    props_().deleteProperty('forcarSemear');
  }
  var depois = Math.max(0, aba.getLastRow() - 1);
  gravarSync_({ status: 'ocioso', etapa: 'importacao',
    mensagem: 'Importados ' + (depois - antes) + ' processos da planilha-mae.' });
  try {
    SpreadsheetApp.getUi().alert([
      'Importacao concluida.',
      '',
      (depois - antes) + ' processos novos vieram de "Clientes e Processos".',
      'Base geral: ' + depois + ' processos.',
      '',
      'Rode "Sincronizar agora (completa)" para buscar os dados deles no DataJud e no DJEN.'
    ].join('\n'));
  } catch (e) {}
}

/**
 * Inclui na Base geral processos informados por numero.
 *
 * A semeadura normal so olha para a propria Base geral e para a planilha-mae
 * "Clientes e Processos". Um processo que saiu no diario e nao esta em nenhuma
 * das duas nunca entra sozinho - e o painel Publicacoes DJEN chama esses de
 * publicacoes orfas. Este item fecha essa lacuna: cola-se a lista de numeros
 * (um por linha, com ou sem mascara) e eles passam a existir na base; a
 * sincronizacao seguinte busca os dados de cada um no DataJud e no DJEN.
 *
 * Nao sobrescreve nada: numero ja existente e ignorado.
 */
function incluirProcessosPorNumero() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Incluir processos na Base geral',
    'Cole os numeros dos processos, um por linha (com ou sem mascara).\n' +
    'Os que ja existirem serao ignorados.',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var brutos = String(resp.getResponseText() || '').split(/[\r\n,;]+/);
  var novos = [], vistos = {};
  for (var i = 0; i < brutos.length; i++) {
    var d = soDigitos_(brutos[i]);
    if (d.length !== 20 || vistos[d]) continue;
    vistos[d] = 1;
    novos.push(d);
  }
  if (!novos.length) {
    ui.alert('Nenhum numero valido. O numero do CNJ tem 20 digitos.');
    return;
  }

  var res = incluirNumeros_(novos, 'menu');
  ui.alert([
    res.incluidos + ' processo(s) incluido(s) na Base geral.',
    res.jaExistiam + ' ja existiam e foram ignorados.',
    '',
    res.incluidos ? 'Rode "Sincronizar agora (completa)" para buscar os dados deles no DataJud e no DJEN.'
                  : 'Nada a fazer.'
  ].join('\n'));
}

/**
 * Nucleo da inclusao, usado pelo menu, pelo Web App e pela semeadura do diario.
 * Recebe digitos ja normalizados; devolve o que fez. Nao sobrescreve nada:
 * numero que ja existe e ignorado.
 */
function incluirNumeros_(digitos, origem) {
  var ss = planilha_();
  garantirConfig_(ss); garantirSync_(ss); garantirDJEN_(ss); garantirCabecalho_(ss);
  var aba = ss.getSheetByName(cfg_('ABA_BASE'));

  /* Ja existentes na base — comparacao por digitos, porque a mascara varia. */
  var existentes = {};
  var ult = aba.getLastRow();
  if (ult > 1) {
    var nums = aba.getRange(2, 2, ult - 1, 1).getValues();
    for (var j = 0; j < nums.length; j++) {
      var n = soDigitos_(nums[j][0]);
      if (n) existentes[n] = 1;
    }
  }

  var incluir = [];
  for (var k = 0; k < digitos.length; k++) {
    if (!existentes[digitos[k]]) { incluir.push(digitos[k]); existentes[digitos[k]] = 1; }
  }
  if (!incluir.length) {
    return { incluidos: 0, jaExistiam: digitos.length, numeros: [] };
  }

  /* Grava so o numero, com mascara: as demais colunas sao preenchidas pela
     sincronizacao, que e quem fala com o DataJud. */
  var linhas = incluir.map(function (d) {
    var linha = new Array(N_COLUNAS);
    for (var c = 0; c < N_COLUNAS; c++) linha[c] = '';
    linha[1] = mascaraCNJ_(d);
    linha[27] = 'Incluido via ' + (origem || 'comando');   // AB — Origem do Cadastro
    return linha;
  });
  aba.getRange(aba.getLastRow() + 1, 1, linhas.length, N_COLUNAS).setValues(linhas);

  gravarSync_({ status: 'ocioso', etapa: 'inclusao',
    mensagem: 'Incluidos ' + linhas.length + ' processos (' + (origem || 'comando') + ').' });

  return { incluidos: incluir.length, jaExistiam: digitos.length - incluir.length,
           numeros: incluir.map(mascaraCNJ_) };
}

/** Normaliza uma lista solta (texto colado, parametro de URL, array). */
function numerosValidos_(entrada) {
  var brutos = Array.isArray(entrada) ? entrada
             : String(entrada || '').split(/[\r\n,;\s]+/);
  var out = [], vistos = {};
  for (var i = 0; i < brutos.length; i++) {
    var d = soDigitos_(brutos[i]);
    if (d.length !== 20 || vistos[d]) continue;
    vistos[d] = 1;
    out.push(d);
  }
  return out;
}

/**
 * Semeadura automatica pelo diario: pergunta ao DJEN quais processos tiveram
 * publicacao para as OABs do escritorio e inclui os que a base ainda nao tem.
 *
 * E o caminho que dispensa qualquer acao manual. Se a API responder 403 - ela
 * costuma recusar chamadas de fora do Brasil, e o Apps Script sai dos IPs do
 * Google - a funcao apenas nao encontra nada e a sincronizacao segue; nesse
 * caso o painel Publicacoes DJEN, que roda no navegador, continua sendo o
 * caminho que funciona.
 */
function semearDoDJEN_(dias) {
  dias = dias || 30;
  var oabs = listaOABs_();
  if (!oabs.length) return { incluidos: 0, lidas: 0, erro: 'sem OAB no _Config' };

  var fim = new Date();
  var ini = new Date(fim.getTime() - dias * 86400000);
  var fmt = function (d) { return Utilities.formatDate(d, 'America/Bahia', 'yyyy-MM-dd'); };

  var achados = {}, lidas = 0, falhas = 0;
  for (var i = 0; i < oabs.length; i++) {
    var pagina = 1, total = 1;
    while (pagina <= total && pagina <= 60) {
      var url = 'https://comunicaapi.pje.jus.br/api/v1/comunicacao' +
        '?numeroOab=' + encodeURIComponent(oabs[i].n) +
        '&ufOab=' + encodeURIComponent(oabs[i].uf) +
        '&dataDisponibilizacaoInicio=' + fmt(ini) +
        '&dataDisponibilizacaoFim=' + fmt(fim) +
        '&pagina=' + pagina + '&itensPorPagina=50';
      var resp;
      try {
        resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      } catch (e) { falhas++; break; }
      if (resp.getResponseCode() !== 200) { falhas++; break; }
      var corpo;
      try { corpo = JSON.parse(resp.getContentText()); } catch (e) { falhas++; break; }
      var itens = corpo.items || [];
      total = Math.ceil((corpo.count || 0) / 50) || 1;
      for (var k = 0; k < itens.length; k++) {
        lidas++;
        var d = soDigitos_(itens[k].numero_processo || itens[k].numeroprocessocommascara);
        if (d.length === 20) achados[d] = 1;
      }
      pagina++;
      Utilities.sleep(400);
    }
  }

  var lista = Object.keys(achados);
  if (!lista.length) return { incluidos: 0, lidas: lidas, falhas: falhas };
  var res = incluirNumeros_(lista, 'diario');
  res.lidas = lidas;
  res.falhas = falhas;
  return res;
}

/* ===================================================================
   RESULTADO VINDO DO DIARIO
   =================================================================== */

/* Confianca minima para tocar na base. A leitura do dispositivo e heuristica;
   abaixo disto a divergencia continua exposta no painel para decisao humana,
   e nada e escrito. */
var CONF_MINIMA_RESULTADO = 80;

var ABA_LOG_RES = '_ResultadosDJEN';

/* Codigo compacto vindo do painel -> valor gravado na planilha. O painel manda
   a letra para caber na URL; o vocabulario gravado e o da planilha. */
function valorSentenca_(cod) {
  var m = { P: 'Proced\u00eancia', R: 'Proced\u00eancia', I: 'Improced\u00eancia',
            E: 'Extin\u00e7\u00e3o do Processo', A: 'Acordo Homologado' };
  return m[cod] || '';
}
function valorRecurso_(cod) {
  var m = { V: 'Provido', W: 'Provido', N: 'N\u00e3o Provido' };
  return m[cod] || '';
}
/* Rotulo do que foi realmente lido, para o log nao perder o parcial. */
function rotuloDispositivo_(cod) {
  var m = { P: 'Procedente', R: 'Parcialmente procedente', I: 'Improcedente',
            E: 'Extin\u00e7\u00e3o do processo', A: 'Acordo homologado',
            V: 'Recurso provido', W: 'Recurso parcialmente provido', N: 'Recurso n\u00e3o provido' };
  return m[cod] || cod;
}

function garantirLogResultados_(ss) {
  var aba = ss.getSheetByName(ABA_LOG_RES);
  if (!aba) {
    aba = ss.insertSheet(ABA_LOG_RES);
    aba.getRange(1, 1, 1, 9).setValues([[
      'Quando', 'Processo', 'Campo', 'Valor anterior', 'Valor gravado',
      'Dispositivo lido no diario', 'Confianca', 'Data da publicacao', 'Certidao'
    ]]);
    aba.setFrozenRows(1);
  }
  return aba;
}

/* "Vazio" para efeito de escrita: campo em branco ou o valor de espera. Um
   valor posto por gente nunca e considerado vazio. */
function podeGravarSentenca_(atual) {
  var t = norm_(String(atual || ''));
  return !t || t.indexOf('nao sentenciado') === 0;
}
function podeGravarRecurso_(atual) {
  var t = norm_(String(atual || ''));
  return !t || t.indexOf('nao houve recurso') === 0 || t.indexOf('aguardando julgamento') === 0;
}

/**
 * Aplica os resultados lidos no diario.
 * Cada item: { n: 20 digitos, c: codigo, k: 'S'|'R', cf: confianca,
 *              d: data da publicacao, h: hash da certidao }
 */
function aplicarResultados_(itens) {
  var ss = planilha_();
  garantirCabecalho_(ss);
  var aba = ss.getSheetByName(cfg_('ABA_BASE'));
  var ult = aba.getLastRow();
  if (ult < 2) return { gravados: 0, ignorados: itens.length, baixaConfianca: 0 };

  var qtd = ult - 1;
  var nums = aba.getRange(2, 2, qtd, 1).getValues();
  var linhaDe = {};
  for (var i = 0; i < qtd; i++) {
    var n = soDigitos_(nums[i][0]);
    if (n.length === 20 && linhaDe[n] === undefined) linhaDe[n] = i;
  }

  /* AC (29) e AD (30) — resultado da sentenca e do recurso. */
  var col = aba.getRange(2, 29, qtd, 2).getValues();
  var antes = JSON.stringify(col);

  var log = [], baixa = 0, semLinha = 0, ocupado = 0;
  for (var k = 0; k < itens.length; k++) {
    var it = itens[k] || {};
    var conf = Number(it.cf || 0);
    if (conf < CONF_MINIMA_RESULTADO) { baixa++; continue; }
    var idx = linhaDe[String(it.n || '')];
    if (idx === undefined) { semLinha++; continue; }

    var eSentenca = String(it.k || 'S') === 'S';
    var atual = eSentenca ? col[idx][0] : col[idx][1];
    var novo = eSentenca ? valorSentenca_(it.c) : valorRecurso_(it.c);
    if (!novo) continue;
    var liberado = eSentenca ? podeGravarSentenca_(atual) : podeGravarRecurso_(atual);
    if (!liberado) { ocupado++; continue; }

    if (eSentenca) col[idx][0] = novo; else col[idx][1] = novo;
    log.push([new Date(), formatarCNJ_(it.n),
              eSentenca ? 'Resultado da Sentenca' : 'Resultado do Recurso',
              String(atual || ''), novo, rotuloDispositivo_(it.c), conf,
              String(it.d || ''),
              it.h ? 'https://comunicaapi.pje.jus.br/api/v1/comunicacao/' + it.h + '/certidao' : '']);
  }

  if (JSON.stringify(col) !== antes) {
    aba.getRange(2, 29, qtd, 2).setValues(col);
  }
  if (log.length) {
    var abaLog = garantirLogResultados_(ss);
    abaLog.getRange(abaLog.getLastRow() + 1, 1, log.length, 9).setValues(log);
  }
  gravarSync_({ status: 'ocioso', etapa: 'resultado',
    mensagem: 'Gravados ' + log.length + ' resultados vindos do diario.' });
  return { gravados: log.length, baixaConfianca: baixa,
           semCadastro: semLinha, jaPreenchidos: ocupado };
}

/* Formato compacto para caber na URL: n:codigo:tipo:confianca:data:hash,
   separados por barra vertical. */
function lerResultados_(texto) {
  var out = [];
  String(texto || '').split('|').forEach(function (parte) {
    var c = String(parte).split(':');
    if (c.length < 4) return;
    var n = soDigitos_(c[0]);
    if (n.length !== 20) return;
    out.push({ n: n, c: c[1], k: c[2], cf: Number(c[3] || 0), d: c[4] || '', h: c[5] || '' });
  });
  return out;
}

/** Item de menu: procura processos novos no diario e inclui. */
function buscarProcessosNoDiario() {
  var ui = SpreadsheetApp.getUi();
  var res = semearDoDJEN_(60);
  if (res.erro) { ui.alert('Nao foi possivel: ' + res.erro); return; }
  ui.alert([
    'Diario consultado: ' + (res.lidas || 0) + ' publicacoes lidas.',
    '',
    (res.incluidos || 0) + ' processo(s) novo(s) incluido(s) na Base geral.',
    (res.jaExistiam || 0) + ' ja existiam.',
    res.falhas ? '' : null,
    res.falhas ? 'Atencao: ' + res.falhas + ' consulta(s) falharam. A API do DJEN costuma' : null,
    res.falhas ? 'recusar chamadas de fora do Brasil, e o Apps Script sai dos IPs do Google.' : null,
    res.falhas ? 'Nesse caso use o botao "Incluir na base" do painel Publicacoes DJEN,' : null,
    res.falhas ? 'que roda no seu navegador.' : null
  ].filter(function (x) { return x !== null; }).join('\n'));
}

/** Aplica a mascara do CNJ a 20 digitos: NNNNNNN-DD.AAAA.J.TR.OOOO */
function mascaraCNJ_(d) {
  d = soDigitos_(d);
  if (d.length !== 20) return d;
  return d.slice(0, 7) + '-' + d.slice(7, 9) + '.' + d.slice(9, 13) + '.' +
         d.slice(13, 14) + '.' + d.slice(14, 16) + '.' + d.slice(16, 20);
}

function cancelarSincronizacao() {
  limparGatilhosContinuacao_();
  props_().deleteProperty('cursor');
  gravarSync_({ status: 'ocioso', etapa: '-', mensagem: 'Cancelado pelo usuario.' });
}

/**
 * Web App - chamado pelo botao "Atualizar" do dashboard.
 *   ?acao=status                -> estado atual da sincronizacao
 *   ?acao=atualizar&token=XXXX  -> dispara sincronizacao completa
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var acao = p.acao || 'status';
  var saida;
  try {
    if (acao === 'atualizar') {
      if (String(p.token || '') !== String(cfg_('TOKEN_WEBAPP'))) {
        saida = { ok: false, erro: 'Token invalido.' };
      } else {
        var st = lerSync_();
        if (st.status === 'rodando') {
          saida = { ok: true, jaRodando: true, sync: st };
        } else {
          iniciarSincronizacao_(p.modo === 'incremental' ? 'incremental' : 'completa', 'dashboard');
          limparGatilhosContinuacao_();
          ScriptApp.newTrigger('continuarSincronizacao').timeBased().after(5 * 1000).create();
          saida = { ok: true, iniciada: true, sync: lerSync_() };
        }
      }
    } else if (acao === 'incluir') {
      /* Permite que o painel Publicacoes DJEN mande a lista de processos que
         sairam no diario e nao estao na base — sem colar nada e sem tocar em
         codigo. Mesmo token do acao=atualizar. */
      if (String(p.token || '') !== String(cfg_('TOKEN_WEBAPP'))) {
        saida = { ok: false, erro: 'Token invalido.' };
      } else {
        var lista = numerosValidos_(p.numeros || '');
        if (!lista.length) {
          saida = { ok: false, erro: 'Nenhum numero valido recebido.' };
        } else {
          var r = incluirNumeros_(lista, 'painel DJEN');
          saida = { ok: true, incluidos: r.incluidos, jaExistiam: r.jaExistiam,
                    numeros: r.numeros };
        }
      }
    } else if (acao === 'resultado') {
      /* O painel Publicacoes DJEN manda o que saiu no diario; aqui so entra o
         que passa do limiar de confianca e cai em campo ainda vazio. */
      if (String(p.token || '') !== String(cfg_('TOKEN_WEBAPP'))) {
        saida = { ok: false, erro: 'Token invalido.' };
      } else {
        var itens = lerResultados_(p.dados || '');
        if (!itens.length) {
          saida = { ok: false, erro: 'Nenhum resultado valido recebido.' };
        } else {
          saida = { ok: true, resultado: aplicarResultados_(itens),
                    limiar: CONF_MINIMA_RESULTADO };
        }
      }
    } else if (acao === 'semear') {
      if (String(p.token || '') !== String(cfg_('TOKEN_WEBAPP'))) {
        saida = { ok: false, erro: 'Token invalido.' };
      } else {
        saida = { ok: true, resultado: semearDoDJEN_(Number(p.dias) || 30) };
      }
    } else {
      saida = { ok: true, sync: lerSync_() };
    }
  } catch (err) {
    saida = { ok: false, erro: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(saida))
    .setMimeType(ContentService.MimeType.JSON);
}

/* --------------------------- ORQUESTRACAO ------------------------------- */

function iniciarSincronizacao_(modo, origem) {
  var ss = planilha_();
  garantirConfig_(ss); garantirSync_(ss); garantirDJEN_(ss); garantirCabecalho_(ss);
  props_().setProperty('cursor', JSON.stringify({ etapa: 'base', modo: modo, i: 0 }));
  gravarSync_({
    status: 'rodando', etapa: 'base', inicio: new Date(), fim: '',
    progresso: 0, mensagem: 'Iniciada (' + modo + ') via ' + origem + '.',
    novos: 0, movimentos: 0, publicacoes: 0
  });
}

function executarEtapas_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;              // ja ha execucao em curso
  var t0 = Date.now();
  try {
    var cur = JSON.parse(props_().getProperty('cursor') || 'null');
    if (!cur) { lock.releaseLock(); return; }

    while (true) {
      if (Date.now() - t0 > LIMITE_MS) { agendarContinuacao_(); break; }

      if (cur.etapa === 'base')            cur = etapaBase_(cur);
      else if (cur.etapa === 'datajud')    cur = etapaDataJud_(cur, t0);
      else if (cur.etapa === 'djen')       cur = etapaDJEN_(cur, t0);
      else if (cur.etapa === 'descobrir')  cur = etapaDescobrir_(cur);
      else if (cur.etapa === 'consolidar') cur = etapaConsolidar_(cur, t0);
      else { finalizar_(); break; }

      props_().setProperty('cursor', JSON.stringify(cur));
      if (cur.etapa === 'fim') { finalizar_(); break; }
    }
  } catch (err) {
    gravarSync_({ status: 'erro', mensagem: 'Erro: ' + ((err && err.message) || err) });
    try { console.error(err); } catch (e) {}
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function agendarContinuacao_() {
  limparGatilhosContinuacao_();
  ScriptApp.newTrigger('continuarSincronizacao').timeBased().after(30 * 1000).create();
  gravarSync_({ mensagem: 'Pausa tecnica - retomando em ~30s.' });
}

function limparGatilhosContinuacao_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'continuarSincronizacao') ScriptApp.deleteTrigger(t);
  });
}

function finalizar_() {
  limparGatilhosContinuacao_();
  props_().deleteProperty('cursor');
  gravarSync_({
    status: 'ocioso', etapa: 'concluida', fim: new Date(), progresso: 100,
    mensagem: 'Sincronizacao concluida.', ultimaOk: new Date()
  });
}

/* --------------------------- ETAPA 1 - BASE ----------------------------- */

function etapaBase_(cur) {
  gravarSync_({ etapa: 'base', progresso: 3, mensagem: 'Montando a base de processos...' });
  var ss = planilha_();
  var aba = ss.getSheetByName(cfg_('ABA_BASE'));

  /* 1a. processos ja na planilha — esta lista manda; a planilha-mae so enriquece */
  var existentes = {};
  var ordem = [];
  var ult = aba.getLastRow();
  var colC = [], colZAB = [], colACAF = [], linhas = [];
  if (ult > 1) {
    var qtdL = ult - 1;
    var nums = aba.getRange(2, 2, qtdL, 1).getValues();
    colC   = aba.getRange(2, 3, qtdL, 1).getValues();
    colZAB = aba.getRange(2, 26, qtdL, 3).getValues();
    colACAF = aba.getRange(2, 29, qtdL, 4).getValues();
    for (var i = 0; i < qtdL; i++) {
      var n = soDigitos_(nums[i][0]);
      if (n.length !== 20 || existentes[n]) continue;
      existentes[n] = { linha: i + 2, idx: i, origem: colZAB[i][2] };
      ordem.push(n);
    }
    linhas = nums;
  }

  /* 1b. planilha-mae "Clientes e Processos" */
  var meta = {};
  try {
    var mae = SpreadsheetApp.openById(cfg_('PLANILHA_MAE_ID'))
                            .getSheetByName(cfg_('PLANILHA_MAE_ABA'));
    if (mae) {
      var mv = mae.getDataRange().getDisplayValues();
      var hh = mv[0].map(function (c) { return chave_(c); });
      var ix = function (nome) { return hh.indexOf(chave_(nome)); };
      var cNum  = ix('Nº DO PROCESSO');
      if (cNum < 0) cNum = acharColunaNumero_(hh);
      var cTipo = ix('TIPO DE PROCESSO'), cAut = ix('AUTOR'), cReu = ix('RÉU'),
          cVal  = ix('VALOR DA CAUSA'),   cPar = ix('PARTE REPRESENTADA'),
          cSen  = ix('RESULTADO DE SENTENÇA'), cRec = ix('RESULTADO DO RECURSO'),
          cFas  = ix('FASE PROCESSUAL'),       cAlv = ix('SITUAÇÃO DO ALVARÁ');
      for (var r = 1; r < mv.length; r++) {
        var num = soDigitos_(mv[r][cNum]);
        if (num.length !== 20 || meta[num]) continue;
        var autor = cAut >= 0 ? String(mv[r][cAut] || '').trim() : '';
        var reu   = cReu >= 0 ? String(mv[r][cReu] || '').trim() : '';
        /* PARTE REPRESENTADA guarda o POLO ("Autor"/"Reu"), nao o nome.
           O cliente e o nome de quem o escritorio representa. */
        var polo = cPar >= 0 ? String(mv[r][cPar] || '').trim() : '';
        var cliente = /r[eé]u/i.test(polo) ? reu : autor;
        meta[num] = {
          tipo:    cTipo >= 0 ? String(mv[r][cTipo] || '').trim() : '',
          partes:  (autor && reu) ? (autor + ' x ' + reu) : (autor || reu),
          valor:   cVal >= 0 ? mv[r][cVal] : '',
          polo:    polo,
          cliente: cliente || autor || reu,
          sentenca: cSen >= 0 ? String(mv[r][cSen] || '').trim() : '',
          recurso:  cRec >= 0 ? String(mv[r][cRec] || '').trim() : '',
          fase:     cFas >= 0 ? String(mv[r][cFas] || '').trim() : '',
          alvara:   cAlv >= 0 ? String(mv[r][cAlv] || '').trim() : ''
        };
      }
    }
  } catch (e) {
    gravarSync_({ mensagem: 'Aviso: planilha-mae nao pode ser lida (' + e.message + '). Seguindo com a base atual.' });
  }

  /* 1c. enriquece as linhas que ja estao na planilha (sem sobrescrever o que
         alguem preencheu a mao) */
  if (ordem.length) {
    var mudouC = false, mudouZ = false, mudouM = false;
    ordem.forEach(function (n) {
      var e = existentes[n]; if (e.idx == null) return;
      var m = meta[n]; if (!m) return;
      if (!String(colC[e.idx][0] || '').trim() && m.partes) { colC[e.idx][0] = m.partes; mudouC = true; }
      if (!String(colZAB[e.idx][0] || '').trim() && m.valor)   { colZAB[e.idx][0] = m.valor;   mudouZ = true; }
      if (m.cliente && String(colZAB[e.idx][1] || '').trim() !== m.cliente) {
        colZAB[e.idx][1] = m.cliente; mudouZ = true;
      }
      if (!String(colZAB[e.idx][2] || '').trim())              { colZAB[e.idx][2] = 'Clientes e Processos'; mudouZ = true; }
      /* espelho: a planilha-mae manda nesses quatro, entao sempre reescreve */
      var espelho = [m.sentenca || '', m.recurso || '', m.fase || '', m.alvara || ''];
      for (var q = 0; q < 4; q++) {
        if (String(colACAF[e.idx][q] || '') !== espelho[q]) { colACAF[e.idx][q] = espelho[q]; mudouM = true; }
      }
    });
    if (mudouC) aba.getRange(2, 3, colC.length, 1).setValues(colC);
    if (mudouZ) aba.getRange(2, 26, colZAB.length, 3).setValues(colZAB);
    if (mudouM) aba.getRange(2, 29, colACAF.length, 4).setValues(colACAF);
  }

  /* 1d. semeia a base pela planilha-mae — por padrao SO se a aba estiver vazia,
         para nao inflar uma lista de processos curada a mao */
  var modoMae = norm_(cfg_('INCLUIR_PLANILHA_MAE') || 'SEMPRE');
  var semear = props_().getProperty('forcarSemear') === '1' ||
               modoMae.indexOf('sempre') === 0 ||
               (modoMae.indexOf('so_se_vazia') === 0 && !ordem.length);
  if (semear) {
    Object.keys(meta).forEach(function (num) {
      if (existentes[num]) return;
      existentes[num] = { linha: 0, origem: 'Clientes e Processos' };
      ordem.push(num);
    });
  }

  /* 1e. grava esqueleto das linhas que ainda nao existem
         (processos que so aparecem no DJEN entram na etapa "descobrir") */
  var novos = 0;
  var linhasNovas = [];
  ordem.forEach(function (n) {
    if (existentes[n].linha) return;
    var m = meta[n] || {};
    var linha = [];
    for (var c = 0; c < N_COLUNAS; c++) linha.push('');
    linha[0]  = m.tipo || '';
    linha[1]  = formatarCNJ_(n);
    linha[2]  = m.partes || existentes[n].partes || '';
    linha[25] = m.valor || '';
    linha[26] = m.cliente || '';
    linha[27] = existentes[n].origem || 'Clientes e Processos';
    linha[28] = m.sentenca || '';
    linha[29] = m.recurso || '';
    linha[30] = m.fase || '';
    linha[31] = m.alvara || '';
    linhasNovas.push(linha);
  });
  if (linhasNovas.length) {
    var inicio = Math.max(aba.getLastRow(), 1) + 1;
    if (aba.getMaxRows() < inicio + linhasNovas.length) {
      aba.insertRowsAfter(aba.getMaxRows(), inicio + linhasNovas.length - aba.getMaxRows());
    }
    for (var off = 0; off < linhasNovas.length; off += 500) {
      var bloco = linhasNovas.slice(off, off + 500);
      aba.getRange(inicio + off, 1, bloco.length, N_COLUNAS).setValues(bloco);
    }
  }

  gravarSync_({ progresso: 10, processos: ordem.length, novos: novos,
    mensagem: 'Base montada: ' + ordem.length + ' processos (' + novos + ' novos).' });
  return { etapa: 'datajud', modo: cur.modo, i: 0, novos: novos };
}

function acharColunaNumero_(hh) {
  for (var i = 0; i < hh.length; i++) if (/processo/.test(hh[i]) && /^n/.test(hh[i])) return i;
  return 1;
}
