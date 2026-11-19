/*
O script deve ser vinculado à planilha "TOTAL DA SEMANA". Ele localiza automaticamente a pasta onde essa planilha está salva e processa todas as outras planilhas Google Sheets que estiverem dentro dela, EXCETO:
- a própria planilha "TOTAL DA SEMANA";
- qualquer planilha que tenha "Geral", "Unificada" ou "Modelo" no título.
Em cada planilha de LDR restante, percorre todas as abas. Se houver mais de uma aba do mesmo segmento (na mesma planilha de LDR ou em LDRs diferentes), os leads Higienizados de todas elas são somados e enviados juntos para a aba do segmento na TOTAL DA SEMANA. Para cada aba identificada, procura o cabeçalho "Status do lead" na coluna A dentro das primeiras linhas. As linhas cujo status seja "Higienizado" têm os dados copiados (valores finais + hyperlinks) a partir da coluna B do LDR (a coluna A, de status, NÃO é copiada) para a coluna A da TOTAL DA SEMANA em diante. Nenhuma formatação é transferida.
Se a mesma linha de dados (todos os valores idênticos) aparecer mais de uma vez dentro do
mesmo segmento, apenas a primeira ocorrência é mantida.
Se algum dos 8 segmentos não tiver aba correspondente na TOTAL DA SEMANA, ele é simplesmente ignorado (não interrompe a coleta dos demais).
Se nenhuma aba válida for encontrada nos LDRs, o script monta o diagnóstico detalhado e encerra sem alterar a planilha.
*/

const CONFIG = {
  // As 8 abas de segmento.
  SEGMENT_SHEET_NAMES: [
    'Marketing',
    'Tecnologia e Software',
    'Healthtech',
    'Fintech',
    'Engenharia',
    'Ind. Equip-Metalmec',
    'Ind. Insumos-Tech',
    'Ind. Alimentos-Bebidas',
  ],

  IGNORE_NAME_CONTAINS: ['geral', 'unificada', 'modelo'],

  STATUS_HEADER: 'status do lead',

  STATUS_VALUE: 'higienizado',

  HEADER_ROW: 2, // linha padrão do cabeçalho

  HEADER_SEARCH_MAX_ROWS: 10, // procurar o cabeçalho

  SRC_DATA_START_COL: 2, // coluna B do LDR: pula a coluna A (Status do lead), que não é copiada

  NUM_COLS: 17, // quantidade de colunas copiadas: B(2)-R(18) no LDR -> A(1)-Q(17) na TOTAL DA SEMANA

  BOUNDARY_CHARS: [' ', '(', '-'],
};

function normalizarTexto(texto) {
  let t = (texto === null || texto === undefined) ? '' : texto.toString();

  t = t.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  t = t.replace(/[\u200B\u200C\u200D\uFEFF\u200E\u200F]/g, '');
  t = t.replace(/[\u00A0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000\t\n\r]/g, ' ');
  t = t.replace(/ +/g, ' ');
  t = t.trim().toLowerCase();

  return t;
}

function identificarSegmentoPorNomeAba(nomeAbaOriginal, segmentosNormalizados) {
  const nomeAbaNormalizado = normalizarTexto(nomeAbaOriginal);

  let melhorMatch = null;
  let melhorTamanho = -1;

  segmentosNormalizados.forEach((seg) => {
    const segNorm = seg.normalizado;
    if (nomeAbaNormalizado.indexOf(segNorm) !== 0) {
      return; 
    }

    const proximoChar = nomeAbaNormalizado.charAt(segNorm.length);
    const delimitadorValido =
      proximoChar === '' || CONFIG.BOUNDARY_CHARS.indexOf(proximoChar) !== -1;

    if (!delimitadorValido) {
      return;
    }

    if (segNorm.length > melhorTamanho) {
      melhorTamanho = segNorm.length;
      melhorMatch = seg.oficial;
    }
  });

  return melhorMatch;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Coleta Semanal')
    .addItem('Coletar Higienizados', 'coletarHigienizados')
    .addItem('Distribuir entre BDRs', 'abrirDistribuicaoBDRs')
    .addToUi();
}

function coletarHigienizados() {
  const ui = SpreadsheetApp.getUi();

  try {
    const destSs = SpreadsheetApp.getActiveSpreadsheet();
    const destId = destSs.getId();

    const resp = ui.alert(
      'Coletar Higienizados',
      'Isso vai apagar os DADOS (não a formatação/estrutura) das linhas ' +
        'abaixo do cabeçalho de cada aba de segmento e recoletar todos os ' +
        'leads "Higienizado" das planilhas de LDR da pasta desta semana. ' +
        'Deseja continuar?',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) {
      return;
    }

    const destFile = DriveApp.getFileById(destId);
    const parentsIt = destFile.getParents();
    if (!parentsIt.hasNext()) {
      throw new Error(
        'Não foi possível identificar a pasta onde esta planilha está ' +
          'salva. Verifique se a planilha TOTAL DA SEMANA está dentro da ' +
          'pasta semanal correta.'
      );
    }
    const weekFolder = parentsIt.next();

    const ldrFiles = [];
    const filesIt = weekFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (filesIt.hasNext()) {
      const f = filesIt.next();
      if (f.getId() === destId) continue;
      const nomeNormalizado = normalizarTexto(f.getName());
      const deveIgnorar = CONFIG.IGNORE_NAME_CONTAINS.some(
        (termo) => nomeNormalizado.indexOf(normalizarTexto(termo)) !== -1
      );
      if (deveIgnorar) continue;
      ldrFiles.push(f);
    }

    if (ldrFiles.length === 0) {
      throw new Error(
        'Nenhuma planilha de LDR foi encontrada na pasta "' +
          weekFolder.getName() +
          '". Verifique se as planilhas estão na mesma pasta desta TOTAL DA SEMANA.'
      );
    }

    const segmentosNormalizados = CONFIG.SEGMENT_SHEET_NAMES.map((nome) => ({
      oficial: nome,
      normalizado: normalizarTexto(nome),
    }));

    const segmentSheets = {};
    const segmentosFaltando = [];
    CONFIG.SEGMENT_SHEET_NAMES.forEach((nome) => {
      const sheet = destSs.getSheetByName(nome);
      if (!sheet) {
        segmentosFaltando.push(nome);
        return;
      }
      segmentSheets[nome] = sheet;
    });

    const leadsPorSegmento = {};
    const chavesVistasPorSegmento = {};
    CONFIG.SEGMENT_SHEET_NAMES.forEach((nome) => {
      leadsPorSegmento[nome] = [];
      chavesVistasPorSegmento[nome] = {};
    });
    let sheetsAnalisadas = 0;
    let duplicadosIgnorados = 0;

    // Resultado da checagem de cada aba.
    const diagnostico = [];

    ldrFiles.forEach((file) => {
      const srcSs = SpreadsheetApp.openById(file.getId());
      const sheets = srcSs.getSheets();

      const diagArquivo = {
        nomeArquivo: file.getName(),
        abas: [],
      };
      diagnostico.push(diagArquivo);

      sheets.forEach((srcSheet) => {
        const nomeAbaOriginal = srcSheet.getName();
        const nomeSegmento = identificarSegmentoPorNomeAba(
          nomeAbaOriginal,
          segmentosNormalizados
        );

        const diagAba = {
          nomeAba: nomeAbaOriginal,
          segmento: nomeSegmento, 
          headerRow: null, 
          leadsColetados: 0,
        };
        diagArquivo.abas.push(diagAba);

        if (!nomeSegmento) {
          return; // aba não corresponde a nenhum dos 8 segmentos
        }

        const lastRow = srcSheet.getLastRow();
        if (lastRow < 1) {
          return;
        }

        // Procura o "status do lead"
        const maxHeaderScanRow = Math.min(
          CONFIG.HEADER_SEARCH_MAX_ROWS,
          lastRow
        );
        let headerRow = -1;
        if (maxHeaderScanRow >= 1) {
          const colAValues = srcSheet
            .getRange(1, 1, maxHeaderScanRow, 1)
            .getDisplayValues();
          for (let r = 0; r < colAValues.length; r++) {
            if (normalizarTexto(colAValues[r][0]) === CONFIG.STATUS_HEADER) {
              headerRow = r + 1; 
              break;
            }
          }
        }

        if (headerRow === -1) {
          return; // cabeçalho não encontrado nesta aba
        }

        diagAba.headerRow = headerRow;

        const dataStartRow = headerRow + 1;
        if (lastRow < dataStartRow) {
          sheetsAnalisadas++;
          return; // cabeçalho encontrado, mas sem os dados
        }

        sheetsAnalisadas++;

        const numDataRows = lastRow - headerRow;
        const statusValues = srcSheet
          .getRange(dataStartRow, 1, numDataRows, 1)
          .getDisplayValues();

        for (let i = 0; i < numDataRows; i++) {
          const status = normalizarTexto(statusValues[i][0]);
          if (status !== CONFIG.STATUS_VALUE) continue;

          const srcRowNum = dataStartRow + i;
          const srcRange = srcSheet.getRange(
            srcRowNum,
            CONFIG.SRC_DATA_START_COL,
            1,
            CONFIG.NUM_COLS
          );

          const linhaDados = extrairLinhaDados(srcRange);
          const chave = linhaDados.values.join('||');

          if (chavesVistasPorSegmento[nomeSegmento][chave]) {
            duplicadosIgnorados++;
            continue; // linha idêntica já coletada para este segmento; ignora a repetição
          }
          chavesVistasPorSegmento[nomeSegmento][chave] = true;

          leadsPorSegmento[nomeSegmento].push(linhaDados);
          diagAba.leadsColetados++;
        }
      });
    });

    if (sheetsAnalisadas === 0) {
      throw new Error(montarMensagemDiagnostico(diagnostico));
    }

    let leadsColetados = 0;
    const contagemPorSegmento = {};

    CONFIG.SEGMENT_SHEET_NAMES.forEach((nome) => {
      const destSheet = segmentSheets[nome];
      if (!destSheet) return; // aba não existe na TOTAL DA SEMANA; segmento ignorado

      const linhas = leadsPorSegmento[nome];

      const maxRows = destSheet.getMaxRows();
      if (maxRows > CONFIG.HEADER_ROW) {
        destSheet
          .getRange(
            CONFIG.HEADER_ROW + 1,
            1,
            maxRows - CONFIG.HEADER_ROW,
            CONFIG.NUM_COLS
          )
          .clearContent();
      }

      let destRow = CONFIG.HEADER_ROW + 1; // linha 3
      linhas.forEach((linha) => {
        const destRange = destSheet.getRange(destRow, 1, 1, CONFIG.NUM_COLS);
        aplicarLinhaDados(linha, destRange);
        destRow++;
        leadsColetados++;
      });

      contagemPorSegmento[nome] = linhas.length;
    });

    let relatorio = 'Coleta concluída! Os resultados foram:\n\n';
    CONFIG.SEGMENT_SHEET_NAMES.forEach((nome) => {
      if (!segmentSheets[nome]) {
        relatorio += nome + ': aba não existe, ignorado\n';
        return;
      }
      const qtd = contagemPorSegmento[nome];
      relatorio += nome + ': ' + qtd + '\n';
    });
    relatorio += '\nTotal: ' + leadsColetados;
    relatorio += '\n\nAbas válidas analisadas: ' + sheetsAnalisadas;

    if (duplicadosIgnorados > 0) {
      relatorio += '\nDuplicados ignorados: ' + duplicadosIgnorados;
    }

    if (segmentosFaltando.length > 0) {
      relatorio += '\n\nAviso: os seguintes segmentos não têm aba na TOTAL DA SEMANA e foram ignorados: ' + segmentosFaltando.join(', ');
    }

    ui.alert(relatorio);
  } catch (err) {
    ui.alert('Erro na coleta', 'Ocorreu um erro:\n\n' + err.message, ui.ButtonSet.OK);
  }
}

function montarMensagemDiagnostico(diagnostico) {
  let msg = 'Nenhuma aba válida (correspondente a um dos 8 segmentos, com ' +
    'cabeçalho "Status do lead" na coluna A) foi encontrada nas planilhas ' +
    'de LDR.\n\n';

  msg += 'DIAGNÓSTICO\n';
  msg += '-----------\n';
  msg += 'Planilhas de LDR encontradas: ' + diagnostico.length + '\n\n';

  diagnostico.forEach((diagArquivo) => {
    msg += 'LDR: ' + diagArquivo.nomeArquivo + '\n';

    if (diagArquivo.abas.length === 0) {
      msg += '  (nenhuma aba encontrada nesta planilha)\n\n';
      return;
    }

    diagArquivo.abas.forEach((diagAba) => {
      const segmentoTexto = diagAba.segmento
        ? diagAba.segmento
        : 'NÃO IDENTIFICADO';
      const headerTexto = diagAba.headerRow
        ? 'encontrado na linha ' + diagAba.headerRow
        : (diagAba.segmento ? 'NÃO encontrado' : 'não verificado');

      msg += '  Aba: "' + diagAba.nomeAba + '"\n';
      msg += '    Segmento identificado: ' + segmentoTexto + '\n';
      msg += '    Cabeçalho: ' + headerTexto + '\n';
      msg += '    Leads Higienizados: ' + diagAba.leadsColetados + '\n';
    });

    msg += '\n';
  });

  return msg;
}

function extrairLinhaDados(srcRange) {
  return {
    values: srcRange.getValues()[0],
    richTextValues: srcRange.getRichTextValues()[0],
  };
}

function aplicarLinhaDados(linha, destRange) {
  // Valores finais
  destRange.setValues([linha.values]);

  // Hyperlinks
  const destRowIndex = destRange.getRow();
  const destColStart = destRange.getColumn();
  const destSheet = destRange.getSheet();
  for (let c = 0; c < linha.richTextValues.length; c++) {
    const linkUrl = linha.richTextValues[c].getLinkUrl();
    if (linkUrl) {
      destSheet
        .getRange(destRowIndex, destColStart + c)
        .setRichTextValue(linha.richTextValues[c]);
    }
  }
}