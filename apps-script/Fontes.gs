/**
 * ============================================================================
 *  BM ADVOCACIA - Fontes de dados (DataJud + DJEN), consolidacao e utilitarios.
 *  Este arquivo complementa Codigo.gs (o Apps Script junta os dois).
 * ============================================================================
 */

var LOTE_LINHAS = 60;   // linhas contiguas processadas por vez no DataJud

/* ------------------------- ETAPA 2 - DATAJUD ---------------------------- */

function etapaDataJud_(cur, t0) {
  var ss  = planilha_();
  var aba = ss.getSheetByName(cfg_('ABA_BASE'));
  var ult = aba.getLastRow();
  var proxima = cur.segundaPassada
    ? { etapa: 'consolidar', modo: cur.modo, novos: cur.novos }
    : { etapa: 'djen', modo: cur.modo, oab: 0, pagina: 1, novos: cur.novos };
  if (ult < 2) return proxima;

  var linhaBase = Math.max(2, Number(cur.base || 2));
  if (linhaBase > ult) return proxima;
  var totalLinhas = ult - linhaBase + 1;
  var blocos = Math.ceil(totalLinhas / LOTE_LINHAS);
  var b = cur.i || 0;
  var maxMov = Number(cfg_('MAX_MOVIMENTOS')) || 400;
  var movTotal = Number(lerSync_().movimentos || 0);

  while (b < blocos) {
    if (Date.now() - t0 > LIMITE_MS) {
      gravarSync_({ etapa: 'datajud', progresso: 10 + Math.round(45 * b / blocos),
        movimentos: movTotal, mensagem: 'DataJud: bloco ' + b + '/' + blocos + '.' });
      return { etapa: 'datajud', modo: cur.modo, i: b, base: linhaBase,
               novos: cur.novos, segundaPassada: cur.segundaPassada };
    }

    var linhaIni = linhaBase + b * LOTE_LINHAS;
    var qtd = Math.min(LOTE_LINHAS, ult - linhaIni + 1);
    if (qtd <= 0) break;

    var nums = aba.getRange(linhaIni, 2, qtd, 1).getValues()
      .map(function (r) { return soDigitos_(r[0]); });

    /* agrupa o bloco por endpoint do DataJud */
    var grupos = {};
    nums.forEach(function (n) {
      if (n.length !== 20) return;
      var ep = endpointDataJud_(n);
      if (!ep) return;
      (grupos[ep] = grupos[ep] || []).push(n);
    });

    var reqs = [], numsPorReq = [];
    Object.keys(grupos).forEach(function (ep) {
      for (var i = 0; i < grupos[ep].length; i += LOTE_DATAJUD) {
        numsPorReq.push(grupos[ep].slice(i, i + LOTE_DATAJUD));
        reqs.push({
          url: 'https://api-publica.datajud.cnj.jus.br/' + ep + '/_search',
          method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          headers: { Authorization: cfg_('DATAJUD_APIKEY') },
          payload: JSON.stringify({
            size: LOTE_DATAJUD * 5,
            query: { terms: { numeroProcesso: grupos[ep].slice(i, i + LOTE_DATAJUD) } }
          })
        });
      }
    });

    var porNumero = {}, consultados = {};
    if (reqs.length) {
      var resp = [];
      try { resp = UrlFetchApp.fetchAll(reqs); }
      catch (e) { Utilities.sleep(5000); try { resp = UrlFetchApp.fetchAll(reqs); } catch (e2) { resp = []; } }
      resp.forEach(function (r, ri) {
        if (!r || r.getResponseCode() !== 200) return;
        var body; try { body = JSON.parse(r.getContentText()); } catch (e) { return; }
        /* a consulta desses numeros foi respondida: o que nao voltou e porque
           a fonte nao tem - a linha pode ser limpa sem risco de apagar dado bom */
        (numsPorReq[ri] || []).forEach(function (n) { consultados[n] = 1; });
        var hits = (body.hits && body.hits.hits) || [];
        hits.forEach(function (h) {
          var s = h._source || {};
          var n = soDigitos_(s.numeroProcesso);
          (porNumero[n] = porNumero[n] || []).push(s);
        });
      });
    }

    /* le o bloco: linhas cuja consulta falhou ficam como estao */
    var colA  = aba.getRange(linhaIni, 1, qtd, 1).getValues();
    var colDM = aba.getRange(linhaIni, 4, qtd, 10).getValues();
    var colN  = aba.getRange(linhaIni, 14, qtd, 1).getValues();
    var colQS = aba.getRange(linhaIni, 17, qtd, 3).getValues();
    var colX  = aba.getRange(linhaIni, 24, qtd, 1).getValues();

    for (var i2 = 0; i2 < qtd; i2++) {
      var regs = porNumero[nums[i2]];
      if (regs && regs.length) {
        /* a planilha reflete sempre a ultima pesquisa: sobrescreve tudo */
        var d = consolidarDataJud_(regs, maxMov);
        movTotal += d.qtdMov;
        if (d.tipo) colA[i2][0] = d.tipo;
        colDM[i2] = [d.tribunal, d.vara, d.grau, d.classe, d.assunto, d.orgao,
                     d.ajuizamento, d.sistema, d.formato, d.sigilo];
        colN[i2][0]  = d.movimentos;
        colQS[i2]    = [d.qtdMov, d.ultMovData, d.ultMovNome];
        colX[i2][0]  = d.fase;
      } else if (consultados[nums[i2]]) {
        /* consultado e sem retorno: limpa para nao deixar dado velho na planilha */
        colDM[i2]   = ['', '', '', '', '', '', '', '', '', ''];
        colN[i2][0] = '';
        colQS[i2]   = ['', '', ''];
        colX[i2][0] = '';
      }
    }

    aba.getRange(linhaIni, 1, qtd, 1).setValues(colA);
    aba.getRange(linhaIni, 4, qtd, 10).setValues(colDM);
    aba.getRange(linhaIni, 14, qtd, 1).setValues(colN);
    aba.getRange(linhaIni, 17, qtd, 3).setValues(colQS);
    aba.getRange(linhaIni, 24, qtd, 1).setValues(colX);

    b++;
    gravarSync_({ etapa: 'datajud', progresso: 10 + Math.round(45 * b / blocos),
      movimentos: movTotal, mensagem: 'DataJud: bloco ' + b + '/' + blocos + ' - ' +
      movTotal + ' movimentacoes coletadas.' });
  }
  return proxima;
}

/** Junta os registros do DataJud (um por grau) de um mesmo numero de processo. */
function consolidarDataJud_(regs, maxMov) {
  regs.sort(function (a, b) {
    return String(b.dataHoraUltimaAtualizacao || '')
      .localeCompare(String(a.dataHoraUltimaAtualizacao || ''));
  });

  var graus = [], classes = [], assuntos = [], orgaos = [], sistemas = [], formatos = [];
  var movs = [], sigilo = 0, ajuiz = '', tribunal = '', vara = '';
  var ordemGrau = { G1: 1, JE: 1, G2: 2, TR: 2, TRU: 3, SUP: 4 };
  var menor = null, classePrincipal = '';

  regs.forEach(function (s) {
    var g = String(s.grau || '');
    if (graus.indexOf(g) < 0) graus.push(g);
    tribunal = tribunal || s.tribunal || '';

    var cl = (s.classe && s.classe.nome) || '';
    if (cl && classes.indexOf(g + ': ' + cl) < 0) classes.push(g + ': ' + cl);
    if (!classePrincipal && cl) classePrincipal = cl;

    (s.assuntos || []).forEach(function (a) {
      if (a && a.nome && assuntos.indexOf(a.nome) < 0) assuntos.push(a.nome);
    });

    var org = (s.orgaoJulgador && s.orgaoJulgador.nome) || '';
    if (org && orgaos.indexOf(g + ': ' + org) < 0) orgaos.push(g + ': ' + org);

    var sis = (s.sistema && s.sistema.nome) || '';
    if (sis && sistemas.indexOf(sis) < 0) sistemas.push(sis);
    var fmt = (s.formato && s.formato.nome) || '';
    if (fmt && formatos.indexOf(fmt) < 0) formatos.push(fmt);

    if (Number(s.nivelSigilo || 0) > sigilo) sigilo = Number(s.nivelSigilo || 0);

    var dj = dataDataJud_(s.dataAjuizamento);
    if (dj && (!ajuiz || dj < ajuiz)) ajuiz = dj;

    var og = ordemGrau[g] || 9;
    if (menor === null || og < menor) { menor = og; vara = org; }

    (s.movimentos || []).forEach(function (m) {
      movs.push({
        g: g,
        dt: String(m.dataHora || ''),
        nome: String(m.nome || ''),
        comp: (m.complementosTabelados || [])
                .map(function (c) { return c.nome || c.descricao; })
                .filter(Boolean).join(', ')
      });
    });
  });

  movs.sort(function (a, b) { return String(b.dt).localeCompare(String(a.dt)); });
  var qtd = movs.length;
  var recorte = movs.slice(0, maxMov);
  var texto = recorte.map(function (m) {
    return '[' + m.g + '] ' + dataBR_(m.dt) + ' - ' + m.nome + (m.comp ? ' (' + m.comp + ')' : '');
  }).join('\n');
  if (qtd > recorte.length) {
    texto += '\n... (' + (qtd - recorte.length) + ' movimentacoes mais antigas omitidas)';
  }

  var ultNome = recorte.length ? recorte[0].nome : '';
  var trilha  = recorte.slice(0, 25).map(function (m) { return m.nome; }).join('\n');
  var recente = recorte.slice(0, 5).map(function (m) { return m.nome; }).join('\n');

  return {
    tribunal: tribunal, vara: vara, grau: graus.join(' | '), classe: classes.join(' | '),
    assunto: assuntos.join(' | '), orgao: orgaos.join(' | '),
    ajuizamento: ajuiz ? dataBR_(ajuiz) : '',
    sistema: sistemas.join(' | '), formato: formatos.join(' | '), sigilo: sigilo,
    movimentos: cortar_(texto), qtdMov: qtd,
    ultMovData: recorte.length ? dataBR_(recorte[0].dt) : '',
    ultMovNome: ultNome,
    tipo: classificarTipo_(regs, numero),
    statusProcesso: statusPelosMovimentos_(movs),
    fase: inferirFase_(trilha, classes.join(' | '), ultNome, recente)
  };
}

/**
 * Coluna A - Tipo de Processo.
 * Percorre os registros do mais recente para o mais antigo e devolve o primeiro
 * que caia em um dos tipos pedidos; se nenhum cair, e Processo Principal.
 */
function classificarTipo_(regs) {
  /* A numeracao do CNJ decide antes da classe: um feito com origem 0000 ou
     9000 nasceu no tribunal, nunca e processo principal de 1o grau. Isso
     tambem resolve os casos em que a classe veio vazia do DataJud. */
  var originaria = regs.length && origemOriginaria_(regs[0].numeroProcesso);

  for (var i = 0; i < regs.length; i++) {
    var c = norm_((regs[i].classe && regs[i].classe.nome) || '');
    if (/mandado de seguranca/.test(c))  return 'Mandado de Segurança';
    if (/conflito de competencia/.test(c)) return 'Conflito de Competência';
    if (/agravo interno/.test(c))        return 'Agravo Interno';
    if (/embargos de declaracao/.test(c)) return 'Embargos de Declaração';
    if (/agravo de instrumento/.test(c)) return 'Agravo de Instrumento';
    if (/apelacao/.test(c))              return 'Recurso de Apelação';
    if (/cumprimento de sentenca|cumprimento provisorio|cumprimento de decisao/.test(c)) {
      return 'Cumprimento de Sentença';
    }
  }
  /* Sem classe reconhecida, a origem ainda diz que nao e principal. */
  if (originaria) return 'Originário de tribunal';
  return 'Processo Principal';
}

/* Numeracao unificada do CNJ: NNNNNNN-DD.AAAA.J.TR.OOOO
   Os quatro ultimos digitos (OOOO) sao a UNIDADE DE ORIGEM. Quando valem
   0000, o feito nasceu no proprio tribunal (competencia originaria de 2o
   grau). Alguns tribunais — TJBA, TJPE, TJRJ nesta base — usam 9000 para os
   orgaos de 2o grau dos juizados (turmas recursais). Nos dois casos NAO se
   trata de processo principal de 1o grau; o que ele e vem da classe. */
function origemOriginaria_(numero) {
  var m = String(numero || '').match(/(\d{4})\s*$/);
  if (!m) return false;
  return m[1] === '0000' || m[1] === '9000';
}

/* Mesma regra de statusPelosMovimentos_, aplicada ao texto ja gravado na
   coluna N — uma linha por movimento, da mais recente para a mais antiga. */
function statusPeloTextoDeMovimentos_(texto) {
  if (!texto) return 'Ativo';
  var linhas = texto.split('\n');
  for (var i = 0; i < linhas.length; i++) {
    var n = norm_(linhas[i]);
    if (/desarquiv|reativacao/.test(n)) return 'Ativo';
    if (/baixa definitiva|arquivamento definitivo|arquivado definitivamente/.test(n))
      return 'Arquivado definitivamente';
    if (/arquivamento provisorio|sobrestamento|suspensao do processo/.test(n))
      return 'Arquivado provisoriamente';
    if (/arquivamento|baixa/.test(n)) return 'Arquivado provisoriamente';
  }
  return 'Ativo';
}

/* Arquivamento lido dos ANDAMENTOS, na ordem em que eles acontecem.
   A lista chega do mais recente para o mais antigo, entao vale o primeiro
   movimento encontrado: um processo com Baixa Definitiva seguida de
   Desarquivamento voltou a tramitar, e olhar so "contem arquivamento"
   marcaria como arquivado quem esta ativo.
   Na tabela do CNJ, "Baixa Definitiva" encerra o feito; "Arquivamento" sem
   qualificacao e provisorio (sobrestamento, suspensao, arquivo do cartorio). */
function statusPelosMovimentos_(movs) {
  for (var i = 0; i < movs.length; i++) {
    var n = norm_(movs[i].nome + ' ' + (movs[i].comp || ''));
    if (/desarquiv|reativacao/.test(n)) return 'Ativo';
    if (/baixa definitiva|arquivamento definitivo|arquivado definitivamente/.test(n))
      return 'Arquivado definitivamente';
    if (/arquivamento provisorio|sobrestamento|suspensao do processo/.test(n))
      return 'Arquivado provisoriamente';
    if (/arquivamento|baixa/.test(n)) return 'Arquivado provisoriamente';
  }
  return 'Ativo';
}

function inferirFase_(trilha, classe, ultima, recente) {
  var t = norm_(trilha), c = norm_(classe), u = norm_(ultima), r = norm_(recente || ultima);
  if (/arquivamento definitivo|baixa definitiva|arquivado definitivamente/.test(r)) return 'Arquivado / baixado';
  if (/cumprimento de sentenca|cumprimento provisorio|execucao de titulo/.test(c)) return 'Cumprimento / execução';
  if (/agravo de instrumento/.test(c)) return 'Agravo de instrumento';
  if (/apelacao|recurso inominado|embargos de declaracao/.test(c + '\n' + u)) return 'Recursal / 2º grau';
  if (/transito em julgado/.test(r)) return 'Trânsito em julgado';
  if (/procedencia|improcedencia|homologacao de acordo|julgamento|sentenca/.test(u + '\n' + t)) return 'Sentenciado';
  if (/conclusao para julgamento|conclusos para sentenca|conclusao para decisao/.test(u + '\n' + t)) return 'Concluso';
  if (/audiencia|pericia|contestacao|replica|especificacao de provas/.test(u + '\n' + t)) return 'Instrução';
  if (/citacao|distribuicao|peticao inicial|liminar|tutela/.test(u + '\n' + t)) return 'Inicial / citação';
  return 'Em andamento';
}

/* --------------------------- ETAPA 3 - DJEN ----------------------------- */

function etapaDJEN_(cur, t0) {
  var ss   = planilha_();
  var abaD = ss.getSheetByName(ABA_DJEN);
  var oabs = listaOABs_();
  var oi = cur.oab || 0, pagina = cur.pagina || 1;

  var vistos = {};
  var ultD = abaD.getLastRow();
  if (ultD > 1) {
    abaD.getRange(2, 1, ultD - 1, 1).getValues()
        .forEach(function (r) { vistos[String(r[0])] = 1; });
  }

  var per = periodoDJEN_(cur.modo);
  var pubTotal = Number(lerSync_().publicacoes || 0);
  var buffer = [];

  while (oi < oabs.length) {
    if (Date.now() - t0 > LIMITE_MS) {
      if (buffer.length) gravarDJEN_(abaD, buffer);
      gravarSync_({ etapa: 'djen', publicacoes: pubTotal,
        mensagem: 'DJEN: OAB ' + oabs[oi].n + '/' + oabs[oi].uf + ', pagina ' + pagina + '.' });
      return { etapa: 'djen', modo: cur.modo, oab: oi, pagina: pagina, novos: cur.novos };
    }

    var url = 'https://comunicaapi.pje.jus.br/api/v1/comunicacao'
      + '?numeroOab=' + encodeURIComponent(oabs[oi].n)
      + '&ufOab='     + encodeURIComponent(oabs[oi].uf)
      + '&dataDisponibilizacaoInicio=' + per.ini
      + '&dataDisponibilizacaoFim='    + per.fim
      + '&pagina=' + pagina + '&itensPorPagina=' + DJEN_ITENS_PAG;

    var r = fetchComRetentativa_(url);
    var body = null; try { body = JSON.parse(r.getContentText()); } catch (e) { body = null; }
    var itens = (body && body.items) || [];

    itens.forEach(function (it) {
      var id = String(it.id || (it.numeroComunicacao + '-' + it.numero_processo));
      if (vistos[id]) return;
      vistos[id] = 1;
      buffer.push([
        id,
        soDigitos_(it.numero_processo || it.numeroprocessocommascara),
        String(it.data_disponibilizacao || ''),
        String(it.siglaTribunal || ''),
        String(it.tipoComunicacao || ''),
        String(it.nomeOrgao || ''),
        String(it.link || ''),
        cortar_(String(it.texto || '')),
        (it.destinatarios || []).map(function (d) { return d.nome; }).filter(Boolean).join(' | ')
      ]);
      pubTotal++;
    });

    if (buffer.length >= 400) { gravarDJEN_(abaD, buffer); buffer = []; }

    gravarSync_({ etapa: 'djen',
      progresso: 55 + Math.round(30 * oi / Math.max(1, oabs.length)),
      publicacoes: pubTotal,
      mensagem: 'DJEN: OAB ' + oabs[oi].n + '/' + oabs[oi].uf + ' - pagina ' + pagina +
                ' - ' + pubTotal + ' publicacoes.' });

    if (itens.length < DJEN_ITENS_PAG) { oi++; pagina = 1; } else { pagina++; }
    Utilities.sleep(DJEN_PAUSA_MS);
  }
  if (buffer.length) gravarDJEN_(abaD, buffer);
  return { etapa: 'descobrir', modo: cur.modo, novos: cur.novos };
}

/**
 * Etapa 3b - processos que aparecem no DJEN mas ainda nao estao na base.
 * Entram como linhas novas e voltam para o DataJud numa segunda passada.
 */
function etapaDescobrir_(cur) {
  if (String(cfg_('DESCOBRIR_NO_DJEN')).toUpperCase().indexOf('S') !== 0) {
    return { etapa: 'consolidar', modo: cur.modo, novos: cur.novos };
  }
  gravarSync_({ etapa: 'descobrir', progresso: 84,
    mensagem: 'Procurando processos do DJEN que ainda nao estao na base...' });

  var ss   = planilha_();
  var aba  = ss.getSheetByName(cfg_('ABA_BASE'));
  var abaD = ss.getSheetByName(ABA_DJEN);

  var naBase = {};
  var ult = aba.getLastRow();
  if (ult > 1) {
    aba.getRange(2, 2, ult - 1, 1).getValues().forEach(function (r) {
      var n = soDigitos_(r[0]); if (n.length === 20) naBase[n] = 1;
    });
  }

  var novos = [], vistos = {};
  var ultD = abaD.getLastRow();
  if (ultD > 1) {
    var dv = abaD.getRange(2, 2, ultD - 1, 8).getValues();   // numero .. destinatarios
    for (var i = 0; i < dv.length; i++) {
      var n = soDigitos_(dv[i][0]);
      if (n.length !== 20 || naBase[n] || vistos[n]) continue;
      vistos[n] = 1;
      var linha = []; for (var c = 0; c < N_COLUNAS; c++) linha.push('');
      linha[1]  = formatarCNJ_(n);
      linha[2]  = String(dv[i][7] || '');
      linha[27] = 'Descoberto no DJEN';
      novos.push(linha);
    }
  }
  if (!novos.length) {
    return { etapa: 'consolidar', modo: cur.modo, novos: cur.novos };
  }

  var inicio = Math.max(aba.getLastRow(), 1) + 1;
  if (aba.getMaxRows() < inicio + novos.length) {
    aba.insertRowsAfter(aba.getMaxRows(), inicio + novos.length - aba.getMaxRows());
  }
  for (var off = 0; off < novos.length; off += 500) {
    var bloco = novos.slice(off, off + 500);
    aba.getRange(inicio + off, 1, bloco.length, N_COLUNAS).setValues(bloco);
  }
  var total = Number(cur.novos || 0) + novos.length;
  gravarSync_({ novos: total, processos: aba.getLastRow() - 1,
    mensagem: novos.length + ' processos novos vieram do DJEN - consultando no DataJud...' });

  /* segunda passada do DataJud, so nas linhas novas */
  return { etapa: 'datajud', modo: cur.modo, i: 0, base: inicio,
           novos: total, segundaPassada: true };
}

function gravarDJEN_(abaD, buffer) {
  if (!buffer.length) return;
  var inicio = Math.max(abaD.getLastRow(), 1) + 1;
  if (abaD.getMaxRows() < inicio + buffer.length) {
    abaD.insertRowsAfter(abaD.getMaxRows(), inicio + buffer.length - abaD.getMaxRows());
  }
  abaD.getRange(inicio, 1, buffer.length, 9).setValues(buffer);
}

function periodoDJEN_(modo) {
  var tz = Session.getScriptTimeZone() || 'America/Bahia';
  var hoje = new Date();
  var fim = Utilities.formatDate(hoje, tz, 'yyyy-MM-dd');
  var ini;
  if (modo === 'completa') {
    /* completa = varredura do periodo inteiro toda vez, para a planilha
       refletir exatamente o que o DJEN tem hoje */
    ini = cfg_('DJEN_DATA_INICIAL');
  } else {
    var dias = Number(cfg_('DJEN_JANELA_DIAS')) || 45;
    ini = Utilities.formatDate(new Date(hoje.getTime() - dias * 864e5), tz, 'yyyy-MM-dd');
  }
  return { ini: ini, fim: fim };
}

function fetchComRetentativa_(url) {
  for (var tent = 0; tent < 5; tent++) {
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var c = r.getResponseCode();
    if (c === 200) return r;
    if (c === 429 || c >= 500) { Utilities.sleep(6000 * (tent + 1)); continue; }
    return r;
  }
  return UrlFetchApp.fetch(url, { muteHttpExceptions: true });
}

/* ------------------ ETAPA 4 - CONSOLIDAR / GESTAO ----------------------- */

function etapaConsolidar_(cur, t0) {
  gravarSync_({ etapa: 'consolidar', progresso: 88,
    mensagem: 'Consolidando publicacoes e indicadores de gestao...' });

  var ss   = planilha_();
  var aba  = ss.getSheetByName(cfg_('ABA_BASE'));
  var abaD = ss.getSheetByName(ABA_DJEN);
  var ult  = aba.getLastRow();
  if (ult < 2) return { etapa: 'fim', modo: cur.modo };

  /* agrupa publicacoes por processo */
  var pubs = {};
  var ultD = abaD.getLastRow();
  if (ultD > 1) {
    var dv = abaD.getRange(2, 1, ultD - 1, 9).getValues();
    for (var i = 0; i < dv.length; i++) {
      var n = soDigitos_(dv[i][1]); if (n.length !== 20) continue;
      (pubs[n] = pubs[n] || []).push({
        data: String(dv[i][2] || ''), trib: dv[i][3], tipo: dv[i][4],
        orgao: dv[i][5], texto: String(dv[i][7] || '')
      });
    }
  }

  var qtdL     = ult - 1;
  var maxPub   = Number(cfg_('MAX_PUBLICACOES')) || 200;
  var diasEst  = Number(cfg_('DIAS_ESTAGNADO')) || 90;
  var tz       = Session.getScriptTimeZone() || 'America/Bahia';
  var agora    = new Date();
  var carimbo  = Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm');
  var soData = function (v) {
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'dd/MM/yyyy');
    return String(v || '');
  };

  /* fatiado por tempo: a coluna O carrega o texto integral das publicacoes */
  var ini      = Math.max(0, Number(cur.off || 0));
  var FATIA    = 150;
  var fim      = Math.min(qtdL, ini + FATIA * 12);
  var qtdF     = fim - ini;
  if (qtdF <= 0) {
    gravarSync_({ progresso: 98, mensagem: 'Indicadores de gestao recalculados.' });
    return { etapa: 'fim', modo: cur.modo };
  }

  var nums  = aba.getRange(2 + ini, 2, qtdF, 1).getValues();
  var colQS = aba.getRange(2 + ini, 17, qtdF, 3).getValues();   // Q,R,S
  var colX  = aba.getRange(2 + ini, 24, qtdF, 1).getValues();   // X - fase
  var colN  = aba.getRange(2 + ini, 14, qtdF, 1).getValues();   // N - movimentos

  var outOP = [], outTW = [], outY = [], outAGAH = [];
  var totalPub = Number(lerSync_().publicacoes || 0);
  if (ini === 0) totalPub = 0;

  for (var r2 = 0; r2 < qtdF; r2++) {
    var num = soDigitos_(nums[r2][0]);
    var lista = (pubs[num] || []).sort(function (a, b) {
      return String(b.data).localeCompare(String(a.data));
    });
    var recorte = lista.slice(0, maxPub);
    var textoPub = recorte.map(function (p) {
      return '[' + dataBR_(p.data) + '] ' + p.trib + ' - ' + p.tipo + ' - ' + p.orgao +
             '\n' + p.texto;
    }).join('\n\n----------\n\n');
    if (lista.length > recorte.length) {
      textoPub += '\n\n... (' + (lista.length - recorte.length) + ' publicacoes mais antigas omitidas)';
    }
    totalPub += lista.length;

    var qtdMov   = Number(colQS[r2][0] || 0);
    var ultMovBR = soData(colQS[r2][1]);
    var fase     = String(colX[r2][0] || '');
    var dtMov    = brParaData_(ultMovBR);
    var dtPub    = lista.length ? isoParaData_(lista[0].data) : null;
    var refer    = (dtMov && dtPub) ? (dtMov > dtPub ? dtMov : dtPub) : (dtMov || dtPub);
    var dias     = refer ? Math.floor((agora - refer) / 864e5) : '';

    /* Coluna Y (Situação) fica como estava, para nao quebrar quem ja a usa. */
    var sit;
    if (!qtdMov && !lista.length)                 sit = 'Sem dados nas fontes';
    else if (/arquiv|baixa/.test(norm_(fase)))    sit = 'Arquivado / baixado';
    else if (dias !== '' && dias > diasEst)       sit = 'Estagnado';
    else                                          sit = 'Ativo';

    /* AG e AH: as duas dimensoes separadas, que a coluna Y funde.
       O status do processo sai do texto dos andamentos ja gravado na coluna N
       — e a leitura respeita a ordem, entao Desarquivamento posterior a uma
       Baixa Definitiva devolve o processo para "Ativo". */
    var statusProc = (!qtdMov && !lista.length)
      ? 'Sem dados nas fontes'
      : statusPeloTextoDeMovimentos_(String(colN[r2][0] || ''));
    var statusMov = dias === '' ? 'Sem data de movimento'
                  : (dias > diasEst ? 'Estagnado' : 'Não estagnado');
    outAGAH.push([statusProc, statusMov]);

    outOP.push([cortar_(textoPub), carimbo]);
    outTW.push([dias, lista.length,
                recorte.length ? dataBR_(recorte[0].data) : '',
                recorte.length ? (recorte[0].tipo + ' - ' + recorte[0].orgao) : '']);
    outY.push([sit]);
  }

  for (var off = 0; off < qtdF; off += FATIA) {
    var n2 = Math.min(FATIA, qtdF - off);
    var base = 2 + ini + off;
    aba.getRange(base, 15, n2, 2).setValues(outOP.slice(off, off + n2));
    aba.getRange(base, 20, n2, 4).setValues(outTW.slice(off, off + n2));
    aba.getRange(base, 25, n2, 1).setValues(outY.slice(off, off + n2));
    aba.getRange(base, 33, n2, 2).setValues(outAGAH.slice(off, off + n2));   // AG, AH
  }

  gravarSync_({ progresso: 88 + Math.round(10 * fim / qtdL), publicacoes: totalPub,
    processos: qtdL,
    mensagem: fim >= qtdL ? 'Indicadores de gestao recalculados.'
                          : 'Consolidando: ' + fim + '/' + qtdL + ' processos.' });
  if (fim >= qtdL) return { etapa: 'fim', modo: cur.modo };
  return { etapa: 'consolidar', modo: cur.modo, off: fim, novos: cur.novos };
}

/* ---------------------------- UTILITARIOS ------------------------------- */

function props_() { return PropertiesService.getScriptProperties(); }

/**
 * A planilha de trabalho. Funciona nos dois modos:
 *  - script vinculado a planilha  -> getActive()
 *  - projeto independente          -> openById(PLANILHA_ID)
 */
var _SS = null;
function planilha_() {
  if (_SS) return _SS;
  try { _SS = SpreadsheetApp.getActive(); } catch (e) { _SS = null; }
  if (!_SS) {
    var id = (props_().getProperty('cfg_PLANILHA_ID') || CFG_PADRAO.PLANILHA_ID || '').trim();
    if (!id) throw new Error('Defina PLANILHA_ID em CFG_PADRAO: o script nao esta vinculado a uma planilha.');
    _SS = SpreadsheetApp.openById(id);
  }
  return _SS;
}

function cfg_(chave) {
  var cache = props_().getProperty('cfg_' + chave);
  if (cache) return cache;
  var aba = planilha_().getSheetByName(ABA_CFG);
  if (aba) {
    var v = aba.getDataRange().getValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][0]).trim() === chave) {
        var val = String(v[i][1]).trim();
        if (val) { props_().setProperty('cfg_' + chave, val); return val; }
      }
    }
  }
  return CFG_PADRAO[chave] || '';
}

/**
 * Le a lista de OABs do _Config. Tolerante ao formato: "41438/BA", "BA/41438",
 * "OAB/BA 41438" e "41438 BA" caem todos no mesmo lugar. Duplicatas somem.
 */
function listaOABs_() {
  var vistos = {}, saida = [];
  String(cfg_('OABS')).split(/[,;\n]/).forEach(function (parte) {
    var t = String(parte).replace(/oab/ig, ' ').trim();
    if (!t) return;
    var numero = (t.match(/\d+[A-Za-z\-]*/) || [])[0];
    var uf = (t.match(/(?:^|[^A-Za-z])([A-Za-z]{2})(?:[^A-Za-z]|$)/) || [])[1];
    if (!numero || !uf) return;
    var chave = numero + '/' + uf.toUpperCase();
    if (vistos[chave]) return;
    vistos[chave] = 1;
    saida.push({ n: numero, uf: uf.toUpperCase() });
  });
  return saida;
}

function garantirConfig_(ss) {
  var aba = ss.getSheetByName(ABA_CFG);
  if (!aba) {
    aba = ss.insertSheet(ABA_CFG);
    var desc = {
      ABA_BASE: 'Aba onde os dados sao gravados',
      PLANILHA_MAE_ID: 'ID da planilha "Clientes e Processos" (fonte da base)',
      PLANILHA_MAE_ABA: 'Aba da planilha-mae',
      INCLUIR_PLANILHA_MAE: 'SEMPRE (padrao - todo processo de Clientes e Processos entra na base) | SO_SE_VAZIA (so semeia com a aba vazia) | NUNCA. Em qualquer opcao a planilha-mae enriquece partes, cliente e valor da causa.',
      OABS: 'OABs do escritorio usadas na varredura do DJEN (numero/UF, separadas por virgula)',
      DJEN_DATA_INICIAL: 'Data inicial da 1a varredura historica do DJEN',
      DJEN_JANELA_DIAS: 'Janela (dias) das varreduras seguintes',
      DATAJUD_APIKEY: 'Chave publica da API do DataJud (CNJ)',
      DESCOBRIR_NO_DJEN: 'SIM = inclui na base processos novos achados pela OAB',
      MAX_MOVIMENTOS: 'Maximo de movimentacoes gravadas por processo',
      MAX_PUBLICACOES: 'Maximo de publicacoes gravadas por processo',
      TOKEN_WEBAPP: 'Senha do botao Atualizar do dashboard',
      DIAS_ESTAGNADO: 'Dias sem movimento para marcar como Estagnado'
    };
    var linhas = [['Chave', 'Valor', 'O que e']];
    Object.keys(CFG_PADRAO).forEach(function (k) {
      linhas.push([k, CFG_PADRAO[k], desc[k] || '']);
    });
    aba.getRange(1, 1, linhas.length, 3).setValues(linhas);
    aba.getRange(1, 1, 1, 3).setFontWeight('bold');
    aba.setColumnWidth(1, 190); aba.setColumnWidth(2, 380); aba.setColumnWidth(3, 430);
  }
  else {
    /* _Config ja existe: acrescenta apenas as chaves que ainda nao estao la,
       sem tocar em nada que o usuario tenha ajustado. */
    var v = aba.getDataRange().getValues();
    var tem = {};
    for (var i = 0; i < v.length; i++) tem[String(v[i][0]).trim()] = 1;
    var faltando = Object.keys(CFG_PADRAO).filter(function (k) { return !tem[k]; });
    if (faltando.length) {
      aba.getRange(v.length + 1, 1, faltando.length, 2).setValues(
        faltando.map(function (k) { return [k, CFG_PADRAO[k]]; }));
    }
  }
  Object.keys(CFG_PADRAO).forEach(function (k) { props_().deleteProperty('cfg_' + k); });
  return aba;
}

/**
 * Repoe no _Config a lista de OABs padrao do codigo. Util depois de acrescentar
 * uma OAB nova no CFG_PADRAO, ja que o _Config tem precedencia sobre ele.
 */
function aplicarOABsPadrao() {
  var aba = garantirConfig_(planilha_());
  var v = aba.getDataRange().getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === 'OABS') {
      aba.getRange(i + 1, 2).setValue(CFG_PADRAO.OABS);
      break;
    }
  }
  props_().deleteProperty('cfg_OABS');
  var lidas = listaOABs_().map(function (o) { return o.n + '/' + o.uf; }).join(', ');
  gravarSync_({ mensagem: 'OABs do DJEN: ' + lidas });
  try {
    SpreadsheetApp.getUi().alert('OABs atualizadas no _Config:\n\n' + lidas);
  } catch (e) {}
}

var SYNC_ORDEM = ['status', 'etapa', 'inicio', 'fim', 'progresso', 'mensagem',
                  'processos', 'novos', 'movimentos', 'publicacoes', 'ultimaOk'];

function garantirSync_(ss) {
  var aba = ss.getSheetByName(ABA_SYNC);
  if (!aba) {
    aba = ss.insertSheet(ABA_SYNC);
    aba.getRange(1, 1, SYNC_ORDEM.length, 1)
       .setValues(SYNC_ORDEM.map(function (r) { return [r]; })).setFontWeight('bold');
    aba.getRange(1, 2).setValue('ocioso');
    aba.setColumnWidth(1, 120); aba.setColumnWidth(2, 620);
  }
  return aba;
}

function garantirDJEN_(ss) {
  var aba = ss.getSheetByName(ABA_DJEN);
  if (!aba) {
    aba = ss.insertSheet(ABA_DJEN);
    aba.getRange(1, 1, 1, 9).setValues([['id', 'numero_processo', 'data', 'tribunal',
      'tipo', 'orgao', 'link', 'texto', 'destinatarios']]).setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}

function garantirCabecalho_(ss) {
  var aba = ss.getSheetByName(cfg_('ABA_BASE'));
  if (!aba) aba = ss.insertSheet(cfg_('ABA_BASE'));
  if (aba.getMaxColumns() < N_COLUNAS) aba.insertColumnsAfter(aba.getMaxColumns(), N_COLUNAS - aba.getMaxColumns());
  aba.getRange(1, 1, 1, N_COLUNAS).setValues([COLUNAS]).setFontWeight('bold');
  aba.setFrozenRows(1);
  return aba;
}

function gravarSync_(obj) {
  var aba = garantirSync_(planilha_());
  SYNC_ORDEM.forEach(function (k, i) {
    if (obj[k] !== undefined) aba.getRange(i + 1, 2).setValue(obj[k]);
  });
  SpreadsheetApp.flush();
}

function lerSync_() {
  var aba = garantirSync_(planilha_());
  var v = aba.getRange(1, 2, SYNC_ORDEM.length, 1).getDisplayValues();
  var o = {};
  SYNC_ORDEM.forEach(function (k, i) { o[k] = v[i][0]; });
  return o;
}

function soDigitos_(v) { return String(v == null ? '' : v).replace(/\D/g, ''); }

function formatarCNJ_(n) {
  n = soDigitos_(n);
  if (n.length !== 20) return n;
  return n.slice(0, 7) + '-' + n.slice(7, 9) + '.' + n.slice(9, 13) + '.' +
         n.slice(13, 14) + '.' + n.slice(14, 16) + '.' + n.slice(16);
}

function norm_(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
}

/** Normalizacao agressiva para casar nomes de coluna (ignora acento e pontuacao). */
function chave_(s) { return norm_(s).replace(/[^a-z0-9]/g, ''); }

function cortar_(t) {
  t = String(t == null ? '' : t);
  return t.length > MAX_CHARS_CEL ? t.slice(0, MAX_CHARS_CEL) + ' [...truncado]' : t;
}

function dataDataJud_(v) {
  var s = String(v || '');
  if (/^\d{14}$/.test(s)) return s.slice(0, 4) + '-' + s.slice(4, 6) + '-' + s.slice(6, 8);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  return '';
}

function dataBR_(v) {
  var s = String(v || '');
  if (/^\d{14}$/.test(s)) return s.slice(6, 8) + '/' + s.slice(4, 6) + '/' + s.slice(0, 4);
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? (m[3] + '/' + m[2] + '/' + m[1]) : s;
}

/**
 * Aceita Date (o Sheets converte "dd/mm/aaaa" em data ao gravar), "dd/mm/aaaa"
 * e "aaaa-mm-dd". Sem isso, o round-trip pela planilha perdia a data e
 * "Dias sem Movimentacao" saia vazio.
 */
function brParaData_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = String(v || '');
  var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

function isoParaData_(v) { return brParaData_(v); }

/** Endpoint do DataJud a partir do numero CNJ (...AAAA.J.TR.OOOO). */
var UF_ESTADUAL = { '01': 'ac', '02': 'al', '03': 'ap', '04': 'am', '05': 'ba', '06': 'ce',
  '07': 'df', '08': 'es', '09': 'go', '10': 'ma', '11': 'mt', '12': 'ms', '13': 'mg',
  '14': 'pa', '15': 'pb', '16': 'pr', '17': 'pe', '18': 'pi', '19': 'rj', '20': 'rn',
  '21': 'rs', '22': 'ro', '23': 'rr', '24': 'sc', '25': 'se', '26': 'sp', '27': 'to' };

function endpointDataJud_(n) {
  n = soDigitos_(n);
  if (n.length !== 20) return '';
  var J = n.charAt(13), TR = n.slice(14, 16), k = Number(TR);
  if (J === '8') { var uf = UF_ESTADUAL[TR]; return uf ? 'api_publica_tj' + uf : ''; }
  if (J === '4') return (k >= 1 && k <= 6)  ? 'api_publica_trf' + k : '';
  if (J === '5') return (k >= 1 && k <= 24) ? 'api_publica_trt' + k : '';
  if (J === '6') { var ufe = UF_ESTADUAL[TR]; return ufe ? 'api_publica_tre' + ufe : ''; }
  if (J === '3') return 'api_publica_stj';
  if (J === '7') return 'api_publica_stm';
  if (J === '9') { var m2 = { '13': 'mg', '21': 'rs', '26': 'sp' }[TR]; return m2 ? 'api_publica_tjm' + m2 : ''; }
  return '';
}
