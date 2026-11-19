/*
Lê a coluna R (18) a partir da linha 3 de cada aba de segmento para saber quais leads ainda estão sem BDR responsável. Escreve o nome do BDR na coluna R desses leads.

"Indústria" sozinho (sem qualificador) é tratado como uma aba virtual que distribui a
quantidade solicitada entre as abas listadas em DIST_CONFIG.ABAS_INDUSTRIA, preenchendo
nessa ordem até esgotar a quantidade ou a disponibilidade. Se no mesmo envio também
houver um pedido específico para uma dessas abas (ex: "Ind. Equip-Metalmec"), a
quantidade do pedido específico é reservada primeiro, e "Indústria" só usa o que sobrar.
*/

const DIST_CONFIG = {
  COLUNA_BDR: 18,               // coluna R
  LINHA_INICIAL_DADOS: 3,
  COLUNA_SETOR_TECNOLOGIA: 9,   // coluna I (deslocada 1 coluna à esquerda após a coleta passar a pular a coluna A/Status do LDR)
  ABA_TECNOLOGIA: 'Tecnologia e Software',
  ABAS_INDUSTRIA: ['Ind. Equip-Metalmec', 'Ind. Insumos-Tech', 'Ind. Alimentos-Bebidas'],
};

const DIST_ALIASES = {
  'marketing': 'Marketing',
  'fintech': 'Fintech',
  'fintechs': 'Fintech',
  'healthtech': 'Healthtech',
  'health tech': 'Healthtech',
  'engenharia': 'Engenharia',
  'industria metal': 'Ind. Equip-Metalmec',
  'industria metal mecanica': 'Ind. Equip-Metalmec',
  'metal mecanica': 'Ind. Equip-Metalmec',
  'metalmecanica': 'Ind. Equip-Metalmec',
  'equip metalmec': 'Ind. Equip-Metalmec',
  'industria equipamentos': 'Ind. Equip-Metalmec',
  'industria insumos': 'Ind. Insumos-Tech',
  'insumos tech': 'Ind. Insumos-Tech',
  'industria alimentos': 'Ind. Alimentos-Bebidas',
  'alimentos': 'Ind. Alimentos-Bebidas',
  'alimentos bebidas': 'Ind. Alimentos-Bebidas',
  'industria alimentos bebidas': 'Ind. Alimentos-Bebidas',
};

/**
 * Ponto de entrada do menu "Distribuir entre BDRs".
 */
function abrirDistribuicaoBDRs() {
  const html = HtmlService.createHtmlOutputFromFile('DistribuicaoBDRs')
    .setWidth(560)
    .setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, 'Distribuir entre BDRs');
}

function _normalizarTexto(txt) {
  return (txt || '')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[.\-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function _chaveSemEspacos(txt) {
  return _normalizarTexto(txt).replace(/\s+/g, '');
}

function _resolverSegmento(valorBruto) {
  const v = _normalizarTexto(valorBruto);
  if (!v) return null;

  // Casos especiais de Tecnologia.
  if (v === 'ti software' || v === 'tecnologia software') {
    return { sheetName: DIST_CONFIG.ABA_TECNOLOGIA, filtroJ: 'TI Software', labelExibicao: 'TI Software' };
  }
  if (v === 'ti servico' || v === 'tecnologia servico') {
    return { sheetName: DIST_CONFIG.ABA_TECNOLOGIA, filtroJ: 'TI Serviço', labelExibicao: 'TI Serviço' };
  }
  if (v.indexOf('tecnologia') !== -1) {
    return { sheetName: DIST_CONFIG.ABA_TECNOLOGIA, filtroJ: null, labelExibicao: 'Tecnologia' };
  }

  if (DIST_ALIASES[v]) {
    return { sheetName: DIST_ALIASES[v], filtroJ: null, labelExibicao: DIST_ALIASES[v] };
  }

  // Caso especial de Indústria: sem qualificador específico, distribui entre as abas.
  if (v.indexOf('industria') !== -1) {
    return {
      sheetNames: DIST_CONFIG.ABAS_INDUSTRIA,
      filtroJ: null,
      labelExibicao: 'Indústria (distribuído entre abas)',
      multi: true,
    };
  }

  let match = CONFIG.SEGMENT_SHEET_NAMES.find((n) => _normalizarTexto(n) === v);

  if (!match) {
    const vSemEspaco = _chaveSemEspacos(valorBruto);
    match = CONFIG.SEGMENT_SHEET_NAMES.find((n) => {
      const nSemEspaco = _chaveSemEspacos(n);
      return nSemEspaco.indexOf(vSemEspaco) !== -1 || vSemEspaco.indexOf(nSemEspaco) !== -1;
    });
  }

  if (!match) return null;
  return { sheetName: match, filtroJ: null, labelExibicao: match };
}

function verificarDisponibilidadeDistribuicao(linhas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resolucao = _resolverLinhas(linhas);
  if (resolucao.erros.length > 0) {
    return { ok: false, erros: resolucao.erros };
  }
  const checagem = _checarDisponibilidade(ss, resolucao.linhasResolvidas);
  return { ok: checagem.erros.length === 0, erros: checagem.erros };
}

function _resolverLinhas(linhas) {
  const erros = [];
  const linhasResolvidas = [];

  (linhas || []).forEach((linha) => {
    const bdr = (linha.bdr || '').toString().trim();
    const qtd = Number(linha.quantidade);

    if (!bdr) {
      erros.push('Existe uma linha sem nome de BDR (segmento "' + linha.segmento + '").');
      return;
    }
    if (isNaN(qtd) || qtd < 0) {
      erros.push('Quantidade inválida para "' + bdr + '" em "' + linha.segmento + '".');
      return;
    }
    if (qtd === 0) return;

    const resolvido = _resolverSegmento(linha.segmento);
    if (!resolvido) {
      erros.push('Segmento não reconhecido: "' + linha.segmento + '" (BDR ' + bdr + ').');
      return;
    }

    linhasResolvidas.push({ bdr, quantidade: qtd, ...resolvido });
  });

  return { linhasResolvidas, erros };
}

//Conta leads com R vazio
function _checarDisponibilidade(ss, linhasResolvidas) {
  const erros = [];

  // 1) Agrupa e valida as linhas específicas (não-multi) por aba+filtro, como antes.
  const gruposEspecificos = {};
  linhasResolvidas.filter((l) => !l.multi).forEach((l) => {
    const chave = l.sheetName + '||' + (l.filtroJ || '');
    if (!gruposEspecificos[chave]) {
      gruposEspecificos[chave] = { sheetName: l.sheetName, filtroJ: l.filtroJ, labelExibicao: l.filtroJ || l.sheetName, total: 0 };
    }
    gruposEspecificos[chave].total += l.quantidade;
  });

  Object.keys(gruposEspecificos).forEach((chave) => {
    const grupo = gruposEspecificos[chave];
    const disponivel = _contarDisponiveis(ss, grupo.sheetName, grupo.filtroJ);
    if (grupo.total > disponivel) {
      erros.push(grupo.labelExibicao + ' — solicitados: ' + grupo.total + ', disponíveis: ' + disponivel);
    }
  });

  // 2) Valida as linhas "Indústria" (multi) contra a capacidade que sobra em cada
  // aba da família DEPOIS de reservar o que já foi pedido especificamente (sem
  // filtro) para aquela mesma aba — evita contar o mesmo lead duas vezes.
  const totalMulti = linhasResolvidas.filter((l) => l.multi).reduce((soma, l) => soma + l.quantidade, 0);

  if (totalMulti > 0) {
    let capacidadeRestante = 0;
    DIST_CONFIG.ABAS_INDUSTRIA.forEach((nomeAba) => {
      const chaveEspecifica = nomeAba + '||';
      const reservadoEspecifico = gruposEspecificos[chaveEspecifica] ? gruposEspecificos[chaveEspecifica].total : 0;
      const disponivelAba = _contarDisponiveis(ss, nomeAba, null);
      capacidadeRestante += Math.max(0, disponivelAba - reservadoEspecifico);
    });

    if (totalMulti > capacidadeRestante) {
      erros.push('Indústria (distribuído entre abas) — solicitados: ' + totalMulti + ', disponíveis: ' + capacidadeRestante);
    }
  }

  return { erros };
}

function _contarDisponiveis(ss, sheetName, filtroJ) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow < DIST_CONFIG.LINHA_INICIAL_DADOS) return 0;

  const numRows = lastRow - DIST_CONFIG.LINHA_INICIAL_DADOS + 1;
  const colA = sheet.getRange(DIST_CONFIG.LINHA_INICIAL_DADOS, 1, numRows, 1).getDisplayValues();
  const colR = sheet.getRange(DIST_CONFIG.LINHA_INICIAL_DADOS, DIST_CONFIG.COLUNA_BDR, numRows, 1).getDisplayValues();
  const colJ = filtroJ
    ? sheet.getRange(DIST_CONFIG.LINHA_INICIAL_DADOS, DIST_CONFIG.COLUNA_SETOR_TECNOLOGIA, numRows, 1).getDisplayValues()
    : null;

  let count = 0;
  for (let i = 0; i < numRows; i++) {
    const temLead = (colA[i][0] || '').toString().trim() !== '';
    if (!temLead) continue;
    const semResponsavel = (colR[i][0] || '').toString().trim() === '';
    if (!semResponsavel) continue;
    if (filtroJ) {
      const setor = (colJ[i][0] || '').toString().trim();
      if (_normalizarTexto(setor) !== _normalizarTexto(filtroJ)) continue;
    }
    count++;
  }
  return count;
}

function escreverDistribuicaoFinal(linhas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const resolucao = _resolverLinhas(linhas);
  if (resolucao.erros.length > 0) {
    return { ok: false, erros: resolucao.erros };
  }

  const checagem = _checarDisponibilidade(ss, resolucao.linhasResolvidas);
  if (checagem.erros.length > 0) {
    return { ok: false, erros: checagem.erros };
  }

  // Reserva primeiro a quantidade dos pedidos específicos (sem filtro) para cada
  // aba da família de Indústria, para "Indústria" genérica não sobrepor com eles.
  const reservadoPorSheet = {};
  resolucao.linhasResolvidas
    .filter((l) => !l.multi && !l.filtroJ)
    .forEach((l) => {
      reservadoPorSheet[l.sheetName] = (reservadoPorSheet[l.sheetName] || 0) + l.quantidade;
    });

  // Expande as linhas "Indústria (multi)" em sublinhas por aba, preenchendo as
  // abas na ordem definida em DIST_CONFIG.ABAS_INDUSTRIA, respeitando o que já
  // está reservado (seja por pedidos específicos, seja por outras linhas de
  // Indústria já expandidas).
  const linhasExpandidas = [];

  resolucao.linhasResolvidas.forEach((l) => {
    if (!l.multi) {
      linhasExpandidas.push(l);
      return;
    }

    let restante = l.quantidade;
    l.sheetNames.forEach((nomeAba) => {
      if (restante <= 0) return;
      const jaReservado = reservadoPorSheet[nomeAba] || 0;
      const disponivelNaAba = _contarDisponiveis(ss, nomeAba, null) - jaReservado;
      if (disponivelNaAba <= 0) return;

      const pegar = Math.min(restante, disponivelNaAba);
      linhasExpandidas.push({
        bdr: l.bdr,
        quantidade: pegar,
        sheetName: nomeAba,
        filtroJ: null,
        labelExibicao: l.labelExibicao,
      });
      reservadoPorSheet[nomeAba] = jaReservado + pegar;
      restante -= pegar;
    });
  });

  // Agrupa por aba.
  const porSheet = {};

  linhasExpandidas.forEach((l) => {
    if (!porSheet[l.sheetName]) porSheet[l.sheetName] = [];
    porSheet[l.sheetName].push(l);
  });

  let totalGeral = 0;
  const relatorioPorBdr = {};

  Object.keys(porSheet).forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return; // aba não existe; segmento ignorado

    const lastRow = sheet.getLastRow();
    if (lastRow < DIST_CONFIG.LINHA_INICIAL_DADOS) return;

    const numRows = lastRow - DIST_CONFIG.LINHA_INICIAL_DADOS + 1;
    const colA = sheet.getRange(DIST_CONFIG.LINHA_INICIAL_DADOS, 1, numRows, 1).getDisplayValues();
    const rangeR = sheet.getRange(DIST_CONFIG.LINHA_INICIAL_DADOS, DIST_CONFIG.COLUNA_BDR, numRows, 1);
    const colR = rangeR.getDisplayValues();
    const temFiltroTech = porSheet[sheetName].some((l) => l.filtroJ);
    const colJ = temFiltroTech
      ? sheet.getRange(DIST_CONFIG.LINHA_INICIAL_DADOS, DIST_CONFIG.COLUNA_SETOR_TECNOLOGIA, numRows, 1).getDisplayValues()
      : null;

    const indicesLivres = [];
    for (let i = 0; i < numRows; i++) {
      const temLead = (colA[i][0] || '').toString().trim() !== '';
      const livre = (colR[i][0] || '').toString().trim() === '';
      if (temLead && livre) indicesLivres.push(i);
    }

    const usados = {};
    const novosValoresR = colR.map((linha) => linha[0]); 
    const entradas = porSheet[sheetName];
    const comFiltro = entradas.filter((e) => e.filtroJ);
    const semFiltro = entradas.filter((e) => !e.filtroJ);

    comFiltro.forEach((entrada) => {
      let restante = entrada.quantidade;
      for (let i = 0; i < indicesLivres.length && restante > 0; i++) {
        const idx = indicesLivres[i];
        if (usados[idx]) continue;
        const setor = colJ ? (colJ[idx][0] || '').toString().trim() : '';
        if (_normalizarTexto(setor) === _normalizarTexto(entrada.filtroJ)) {
          novosValoresR[idx] = entrada.bdr;
          usados[idx] = true;
          restante--;
        }
      }
    });

    semFiltro.forEach((entrada) => {
      let restante = entrada.quantidade;
      for (let i = 0; i < indicesLivres.length && restante > 0; i++) {
        const idx = indicesLivres[i];
        if (usados[idx]) continue;
        novosValoresR[idx] = entrada.bdr;
        usados[idx] = true;
        restante--;
      }
    });

    rangeR.setValues(novosValoresR.map((v) => [v]));

    entradas.forEach((entrada) => {
      relatorioPorBdr[entrada.bdr] = relatorioPorBdr[entrada.bdr] || {};
      relatorioPorBdr[entrada.bdr][entrada.labelExibicao] =
        (relatorioPorBdr[entrada.bdr][entrada.labelExibicao] || 0) + entrada.quantidade;
      totalGeral += entrada.quantidade;
    });
  });

  let relatorio = 'Distribuição concluída!\n\n';
  Object.keys(relatorioPorBdr).forEach((bdr) => {
    relatorio += bdr + '\n';
    Object.keys(relatorioPorBdr[bdr]).forEach((seg) => {
      relatorio += '  ' + seg + ': ' + relatorioPorBdr[bdr][seg] + '\n';
    });
  });
  relatorio += '\nTotal distribuído: ' + totalGeral;

  return { ok: true, relatorio };
}