/*
 * MODELO DE INTEGRAÇÃO — PCP / ESTOQUE DE CHAPAS
 * IMPORTANTE:
 * 1) Use o ID da planilha ORIGINAL/editável, não o token "2PACX..." da URL publicada.
 * 2) Ajuste SHEET_NAME e os nomes exatos das colunas conforme o cabeçalho real.
 * 3) Faça o deploy como Web App e restrinja o acesso conforme as regras da empresa.
 * 4) Não implante com acesso público se a planilha contiver dados internos/confidenciais.
 */
const CONFIG = {
  SPREADSHEET_ID: 'COLE_AQUI_O_ID_DA_PLANILHA_ORIGINAL',
  SHEET_NAME: 'Estoque',
  HEADER_ROW: 1,
  // Ajuste estes nomes para corresponder exatamente ao cabeçalho da planilha:
  HEADERS: {
    material: 'Material',
    thickness: 'Espessura',
    size: 'Tamanho',
    quantity: 'Quantidade'
  }
};

function doGet(e) {
  try {
    const sheet = getSheet_();
    const values = sheet.getDataRange().getDisplayValues();
    if (values.length < CONFIG.HEADER_ROW) return json_({ok:true, records:[]});
    const headers = values[CONFIG.HEADER_ROW - 1];
    const ix = headerIndexes_(headers);
    const records = values.slice(CONFIG.HEADER_ROW).filter(r => r.some(v => String(v).trim())).map((r,i) => ({
      id: String(i + 1),
      material: String(r[ix.material] || '').trim(),
      thickness: String(r[ix.thickness] || '').trim(),
      size: String(r[ix.size] || '').trim(),
      quantity: Number(String(r[ix.quantity] || '0').replace(',', '.')) || 0,
      unit: 'chapas',
      notes: ''
    })).filter(r => r.material && r.thickness && r.size);
    return json_({ok:true, records:records});
  } catch (err) {
    return json_({ok:false, error:String(err.message || err)});
  }
}

// Grava quantidade na linha que corresponde a material + espessura + tamanho.
// Para inserir novas combinações ou alterar outros campos, amplie este método após validar o esquema.
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.token !== PropertiesService.getScriptProperties().getProperty('API_TOKEN')) {
      return json_({ok:false, error:'Não autorizado'});
    }
    if (body.action !== 'updateQuantity' || !body.record) throw new Error('Ação inválida');
    const r = body.record;
    const sheet = getSheet_(), values = sheet.getDataRange().getDisplayValues();
    const ix = headerIndexes_(values[CONFIG.HEADER_ROW - 1]);
    let rowNumber = -1;
    for (let i=CONFIG.HEADER_ROW;i<values.length;i++) {
      if (norm_(values[i][ix.material])===norm_(r.material) &&
          norm_(values[i][ix.thickness])===norm_(r.thickness) &&
          norm_(values[i][ix.size])===norm_(r.size)) { rowNumber=i+1; break; }
    }
    if (rowNumber < 0) throw new Error('Combinação não encontrada; nenhum dado foi gravado.');
    const qty = Number(r.quantity);
    if (!Number.isInteger(qty) || qty < 0) throw new Error('Quantidade inválida.');
    sheet.getRange(rowNumber, ix.quantity+1).setValue(qty);
    return json_({ok:true, message:'Quantidade atualizada', row:rowNumber});
  } catch (err) {
    return json_({ok:false, error:String(err.message || err)});
  }
}
function getSheet_() {
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID.startsWith('COLE_')) throw new Error('Configure o ID da planilha original.');
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error('A aba configurada não foi encontrada: ' + CONFIG.SHEET_NAME);
  return sh;
}
function headerIndexes_(headers) {
  const normHeaders = headers.map(norm_);
  const out = {};
  Object.keys(CONFIG.HEADERS).forEach(k => {
    const idx = normHeaders.indexOf(norm_(CONFIG.HEADERS[k]));
    if (idx < 0) throw new Error('Cabeçalho não encontrado: ' + CONFIG.HEADERS[k]);
    out[k] = idx;
  });
  return out;
}
function norm_(s) {
  return String(s == null ? '' : s).trim().toLocaleLowerCase('pt-BR').normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
