/**
 * Monitor Cenas de Uso — API em Tempo Real (Google Apps Script)
 * Somente leitura: nunca altera células, fórmulas ou linhas de nenhuma equipe.
 * Privacidade: na planilha CNR - BRÁS retorna apenas contagens agregadas e status estrutural (zero dados pessoais).
 *
 * COMO IMPLANTAR:
 * 1. Abra https://script.google.com e crie um novo projeto.
 * 2. Cole todo este código em Code.gs e salve.
 * 3. Clique em "Implantar" -> "Nova implantação" -> tipo "App da Web".
 * 4. Executar como: "Eu" | Quem pode acessar: "Qualquer pessoa".
 * 5. Copie a URL gerada (termina em /exec) e cole no painel do Netlify.
 */

const CONFIG = {
  CRS_SUDESTE_ID: '1-UF36937geDEdLTwDqnkYD0zAcoQLIACDPgvK_LlFN8',
  PLANILHA_DIARIA_ID: '11zw_rhfcoXxZNcrXPqtfnk3fgYPBTwsXdJfTQcwGpN4',
  CNR_BRAS_ID: '1u8ZAfZL18CuPdycBnJzlSzVUU5peUUw5MNyzFtwkhGY',
  ABA_DIARIA_PADRAO: 'Novembro26',
  SUFIXO_CRS_PADRAO: '112026'
};

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const abaDiaria = params.abaDiaria || CONFIG.ABA_DIARIA_PADRAO;
  const sufixoCrs = params.sufixoCrs || CONFIG.SUFIXO_CRS_PADRAO;

  const resultado = {
    ok: true,
    atualizadoEm: new Date().toISOString(),
    fonte: 'ao_vivo',
    parametros: { abaDiaria: abaDiaria, sufixoCrs: sufixoCrs },
    planilhaDiaria: null,
    crsSudeste: null,
    cnrBras: null,
    inconsistencias: [],
    errosLeitura: []
  };

  try {
    const pd = auditarPlanilhaDiaria_(CONFIG.PLANILHA_DIARIA_ID, abaDiaria);
    resultado.planilhaDiaria = pd.resumo;
    resultado.inconsistencias = resultado.inconsistencias.concat(pd.inconsistencias);
  } catch (err) {
    resultado.errosLeitura.push('Planilha Diária: ' + err.message);
  }

  try {
    const crs = auditarCrsSudeste_(CONFIG.CRS_SUDESTE_ID, sufixoCrs);
    resultado.crsSudeste = crs.resumo;
    resultado.inconsistencias = resultado.inconsistencias.concat(crs.inconsistencias);
  } catch (err) {
    resultado.errosLeitura.push('CRS Sudeste: ' + err.message);
  }

  try {
    const bras = auditarCnrBrasAgregado_(CONFIG.CNR_BRAS_ID);
    resultado.cnrBras = bras.resumo;
    resultado.inconsistencias = resultado.inconsistencias.concat(bras.inconsistencias);
  } catch (err) {
    resultado.errosLeitura.push('CNR - BRÁS: ' + err.message);
  }

  return ContentService
    .createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

function toNum_(val) {
  if (val === null || val === undefined || val === '') return { n: 0, err: false };
  if (typeof val === 'number') return { n: val, err: false };
  const s = String(val).trim();
  if (s === '' || s === '-' || s === '–' || s.toUpperCase() === 'N/A') return { n: 0, err: false };
  if (s.indexOf('#') === 0) return { n: 0, err: true };
  const parsed = Number(s.replace(',', '.'));
  if (isNaN(parsed)) return { n: 0, err: true };
  return { n: parsed, err: false };
}

function auditarPlanilhaDiaria_(sheetId, abaNome) {
  const ss = SpreadsheetApp.openById(sheetId);
  const ws = ss.getSheetByName(abaNome);
  if (!ws) throw new Error('Aba não encontrada: ' + abaNome);

  const lastRow = Math.min(ws.getLastRow(), 120);
  const lastCol = Math.min(ws.getLastColumn(), 62);
  const vals = ws.getRange(1, 1, lastRow, lastCol).getValues();
  const forms = ws.getRange(1, 1, lastRow, lastCol).getFormulas();

  let totalRowIdx = -1;
  for (let r = 0; r < vals.length; r++) {
    if (String(vals[r][0] || '').trim().toLowerCase() === 'total') {
      totalRowIdx = r;
      break;
    }
  }
  if (totalRowIdx < 2) throw new Error('Linha Total não localizada.');

  const inconsistencias = [];
  const serieDiaria = [];
  const equipesMap = {};
  let equipeAtual = '';

  // Mapear equipes por linha (linhas 3 até totalRowIdx)
  const linhaEquipe = [];
  for (let r = 2; r < totalRowIdx; r++) {
    const colA = String(vals[r][0] || '').replace(/\n/g, ' ').trim();
    if (colA) equipeAtual = colA.split(/novembro|outubro|dezembro|janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro/i)[0].trim();
    linhaEquipe[r] = equipeAtual || 'Sem Equipe';
    if (!equipesMap[linhaEquipe[r]]) {
      equipesMap[linhaEquipe[r]] = { equipe: linhaEquipe[r], cenas: 0, soma10h: 0, soma15h: 0, diasAtivos: 0 };
    }
    equipesMap[linhaEquipe[r]].cenas += 1;
  }

  let diasComDados = 0;
  let maiorSubnotificacao = 0;

  for (let d = 1; d <= 30; d++) {
    const c10 = 2 + (d - 1) * 2;
    const c15 = 3 + (d - 1) * 2;
    if (c15 >= lastCol) break;

    const cabecalhoDia = vals[0][c10];
    if (Number(cabecalhoDia) !== d) {
      inconsistencias.push({
        planilha: 'Planilha Diária',
        severidade: 'Atenção',
        equipe: 'Geral',
        local: 'Cabeçalho Dia ' + d,
        detalhe: 'Célula do cabeçalho indica "' + cabecalhoDia + '" no lugar de ' + d + '.'
      });
    }

    let soma10 = 0;
    let soma15 = 0;
    let temDado = false;

    for (let r = 2; r < totalRowIdx; r++) {
      const raw10 = vals[r][c10];
      const raw15 = vals[r][c15];
      if (typeof raw10 === 'number') {
        soma10 += raw10;
        temDado = true;
        equipesMap[linhaEquipe[r]].soma10h += raw10;
      }
      if (typeof raw15 === 'number') {
        soma15 += raw15;
        temDado = true;
        equipesMap[linhaEquipe[r]].soma15h += raw15;
      }
    }

    if (!temDado) continue;
    diasComDados++;

    // Verificar células em branco no turno 15h para dias já preenchidos
    for (let r = 2; r < totalRowIdx; r++) {
      const raw15 = vals[r][c15];
      if (raw15 === '' || raw15 === null) {
        inconsistencias.push({
          planilha: 'Planilha Diária',
          severidade: 'Atenção',
          equipe: linhaEquipe[r],
          local: 'Dia ' + d + ' (15h) · Linha ' + (r + 1),
          detalhe: 'Célula vazia sem registro (' + String(vals[r][1] || '').slice(0, 42) + ').'
        });
      }
    }

    const rep10 = toNum_(vals[totalRowIdx][c10]).n;
    const rep15 = toNum_(vals[totalRowIdx][c15]).n;

    if (Math.abs(soma10 - rep10) > 0.01) {
      const diff = soma10 - rep10;
      if (Math.abs(diff) > Math.abs(maiorSubnotificacao)) maiorSubnotificacao = diff;
      inconsistencias.push({
        planilha: 'Planilha Diária',
        severidade: 'Crítica',
        equipe: 'Geral',
        local: 'Dia ' + d + ' (10h) · Linha ' + (totalRowIdx + 1),
        detalhe: 'Total informado=' + rep10 + ' difere do calculado=' + soma10 + ' (fórmula: ' + (forms[totalRowIdx][c10] || 'valor fixo') + ').'
      });
    }

    if (Math.abs(soma15 - rep15) > 0.01) {
      inconsistencias.push({
        planilha: 'Planilha Diária',
        severidade: 'Crítica',
        equipe: 'Geral',
        local: 'Dia ' + d + ' (15h) · Linha ' + (totalRowIdx + 1),
        detalhe: 'Total informado=' + rep15 + ' difere do calculado=' + soma15 + ' (fórmula: ' + (forms[totalRowIdx][c15] || 'valor fixo') + ').'
      });
    }

    serieDiaria.push({
      dia: d,
      calc10h: soma10,
      rep10h: rep10,
      calc15h: soma15,
      rep15h: rep15
    });
  }

  const equipes = Object.keys(equipesMap).map(function(k) {
    const e = equipesMap[k];
    return {
      equipe: e.equipe,
      cenas: e.cenas,
      media10h: diasComDados ? Number((e.soma10h / diasComDados).toFixed(1)) : 0,
      media15h: diasComDados ? Number((e.soma15h / diasComDados).toFixed(1)) : 0
    };
  }).sort(function(a, b) { return (b.media10h + b.media15h) - (a.media10h + a.media15h); });

  return {
    resumo: {
      aba: abaNome,
      diasRegistrados: diasComDados,
      maiorSubnotificacao: maiorSubnotificacao,
      serieDiaria: serieDiaria,
      equipes: equipes
    },
    inconsistencias: inconsistencias
  };
}

function auditarCrsSudeste_(sheetId, sufixo) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sheets = ss.getSheets().filter(function(s) {
    return s.getName().trim().slice(-sufixo.length) === sufixo;
  });

  const cortes = [];
  const inconsistencias = [];
  let lacunaTotal = 0;

  sheets.forEach(function(ws) {
    const tab = ws.getName().trim();
    const lastRow = ws.getLastRow();
    if (lastRow < 5) return;
    const vals = ws.getRange(5, 1, lastRow - 4, 15).getValues();
    const forms = ws.getRange(5, 1, lastRow - 4, 15).getFormulas();

    let equipeAtual = '';
    let calcSum = 0;
    let repSum = 0;
    let errosSoma = 0;
    let errosData = 0;

    for (let i = 0; i < vals.length; i++) {
      const row = vals[i];
      const fRow = forms[i];
      const rowNum = i + 5;

      const t = String(row[0] || '').trim();
      if (t) {
        if (/^(TOTAL|SOMA)/i.test(t)) continue;
        equipeAtual = t.split('(')[0].trim();
      }

      const dataVal = row[1];
      const endVal = String(row[3] || '').trim();
      if (!endVal && (dataVal === '' || dataVal === null)) continue;

      const g = toNum_(row[6]).n;
      const h = toNum_(row[7]).n;
      const ii = toNum_(row[8]).n;
      const jObj = toNum_(row[9]);
      const calc = g + h + ii;

      calcSum += calc;
      repSum += jObj.n;

      if (jObj.err) {
        errosSoma++;
        inconsistencias.push({
          planilha: 'CRS Sudeste',
          severidade: 'Crítica',
          equipe: equipeAtual,
          local: tab + ' · Linha ' + rowNum,
          detalhe: 'Valor inválido na coluna TOTAL: "' + row[9] + '" (' + endVal.slice(0, 38) + ').'
        });
      } else if (Math.abs(calc - jObj.n) > 0.01) {
        errosSoma++;
        inconsistencias.push({
          planilha: 'CRS Sudeste',
          severidade: 'Crítica',
          equipe: equipeAtual,
          local: tab + ' · Linha ' + rowNum,
          detalhe: 'G(' + g + ')+H(' + h + ')+I(' + ii + ')=' + calc + ' != Informado J=' + jObj.n + ' (fórmula: ' + (fRow[9] || 'fixo') + ').'
        });
      }

      if (!(dataVal instanceof Date) && (!dataVal || String(dataVal).trim() === '' || String(dataVal).trim() === '-')) {
        errosData++;
        inconsistencias.push({
          planilha: 'CRS Sudeste',
          severidade: 'Atenção',
          equipe: equipeAtual,
          local: tab + ' · Linha ' + rowNum,
          detalhe: 'Data ausente ou inválida (' + endVal.slice(0, 38) + ').'
        });
      }
    }

    const dif = repSum - calcSum;
    lacunaTotal += dif;
    cortes.push({
      aba: tab,
      label: tab.slice(0, 2) + '/' + tab.slice(2, 4),
      calculado: calcSum,
      informado: repSum,
      diferenca: dif,
      errosSoma: errosSoma,
      errosData: errosData
    });
  });

  return {
    resumo: {
      totalCortes: cortes.length,
      lacunaAcumulada: lacunaTotal,
      cortes: cortes
    },
    inconsistencias: inconsistencias
  };
}

function auditarCnrBrasAgregado_(sheetId) {
  const ss = SpreadsheetApp.openById(sheetId);
  const wsReg = ss.getSheetByName('Registro Diário');
  const wsBase = ss.getSheetByName('Base');
  const inconsistencias = [];

  const totalRegistros = wsReg ? Math.max(0, wsReg.getLastRow() - 1) : 0;
  let datasInvalidas = 0;

  if (wsReg && totalRegistros > 0) {
    const colDatas = wsReg.getRange(2, 2, totalRegistros, 1).getValues();
    for (let i = 0; i < colDatas.length; i++) {
      const v = colDatas[i][0];
      if (!v || (!(v instanceof Date) && String(v).trim() === '')) datasInvalidas++;
    }
    if (datasInvalidas > 0) {
      inconsistencias.push({
        planilha: 'CNR - BRÁS',
        severidade: 'Atenção',
        equipe: 'CnR Brás',
        local: 'Registro Diário',
        detalhe: datasInvalidas + ' registros sem data válida (contagem agregada, sem exposição de dados nominais).'
      });
    }
  }

  let baseDesalinhada = false;
  if (wsBase && wsBase.getLastRow() >= 2) {
    const amostra = wsBase.getRange(1, 1, 2, 2).getValues();
    if (String(amostra[0][0]).toLowerCase().indexOf('data') !== -1 && !(amostra[1][0] instanceof Date)) {
      baseDesalinhada = true;
      inconsistencias.push({
        planilha: 'CNR - BRÁS',
        severidade: 'Atenção',
        equipe: 'CnR Brás',
        local: 'Aba Base · Colunas A/B',
        detalhe: 'Colunas Data de Inclusão e Equipe/Líder invertidas na amostra inicial e deslocamento a partir da col. V.'
      });
    }
  }

  return {
    resumo: {
      totalRegistros: totalRegistros,
      datasInvalidas: datasInvalidas,
      baseDesalinhada: baseDesalinhada
    },
    inconsistencias: inconsistencias
  };
}
