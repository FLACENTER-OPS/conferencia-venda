function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);

    // Aba Deals
    let sheet = ss.getSheetByName('Deals');
    if (!sheet) {
      sheet = ss.insertSheet('Deals');
      sheet.appendRow([
        'Timestamp', 'Deal ID', 'VIN', 'Veículo', 'Vendedor','Data Venda',
        'Status', 'Financiamento', 'Lendbuzz', 'Tag', 'Parcelado',
        'Trade-In', 'Warranty', 'GAP', 'Rebuilt', 'Obs Vendedor',
        'Motivos Bloqueio', 'Obs Gerente', 'Revisado Em', 'Histórico'
      ]);
      sheet.setFrozenRows(1);
    }

    // Aba Documentos
    let docsSheet = ss.getSheetByName('Documentos');
    if (!docsSheet) {
      docsSheet = ss.insertSheet('Documentos');
      docsSheet.appendRow(['Deal ID', 'Documento', 'Status']);
      docsSheet.setFrozenRows(1);
    }

    const action = data.action;

    if (action === 'submit') {
      // Vendedor submete novo deal ou atualização
      const existingRow = findDealRow(sheet, data.id);

const rowData = [
  new Date().toISOString(),
  data.id, data.vin, data.veiculo, data.vendedor,
  data.dataVenda || '',
        'aguardando_revisao',
        data.finance || '', data.lendbuzz || '', data.tag || '',
        data.parcelado || '', data.tradein || '', data.warranty || '',
        data.gap || '', data.rebuilt || '', data.obs || '',
        '', '', '', JSON.stringify(data.history || [])
      ];

      if (existingRow > 0) {
        // Atualização — preserva histórico
        const oldHistory = sheet.getRange(existingRow, 20).getValue();
        rowData[19] = oldHistory;
        sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
        // Limpa docs antigos e reinsere
        clearDealDocs(docsSheet, data.id);
      } else {
        sheet.appendRow(rowData);
      }

      // Insere documentos
      (data.docs || []).forEach(doc => {
        docsSheet.appendRow([data.id, doc.doc, doc.status]);
      });

      return ContentService
        .createTextOutput(JSON.stringify({ success: true, id: data.id }))
        .setMimeType(ContentService.MimeType.JSON);
} else if (action === 'correction') {
  const row = findDealRow(sheet, data.id);
  if (row > 0) {
    sheet.getRange(row, 7).setValue('em_correcao');
    const docsData = docsSheet.getDataRange().getValues();
    for (let i = 1; i < docsData.length; i++) {
      if (docsData[i][0] === data.id) {
        if (data.corrigidos && data.corrigidos.includes(docsData[i][1])) {
          docsSheet.getRange(i + 1, 3).setValue('corrigido');
        }
      }
    }
    if (data.obs) {
      const existingObs = sheet.getRange(row, 16).getValue();
sheet.getRange(row, 16).setValue(existingObs + '\n[Correção] ' + data.obs);
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'review') {
      // Giovanna revisa
      const row = findDealRow(sheet, data.id);
      if (row > 0) {
        sheet.getRange(row, 7).setValue(data.status); // Status
        sheet.getRange(row, 17).setValue((data.motivos || []).join(', ')); // Motivos
        sheet.getRange(row, 18).setValue(data.obsGerente || ''); // Obs Gerente
        sheet.getRange(row, 19).setValue(new Date().toISOString()); // Revisado em

        // Atualiza status dos docs reprovados
        if (data.status === 'bloqueado' && data.pendentesGerente) {
          updateDocStatus(docsSheet, data.id, data.pendentesGerente);
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);

    } else if (action === 'getDeals') {
      // Review page busca todos os deals
      const deals = getDealsFromSheet(sheet, docsSheet);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, deals }))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Deals');
    const docsSheet = ss.getSheetByName('Documentos');
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({ success: true, deals: [] })).setMimeType(ContentService.MimeType.JSON);
    const deals = getDealsFromSheet(sheet, docsSheet);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, deals }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function findDealRow(sheet, dealId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === dealId) return i + 1;
  }
  return -1;
}

function clearDealDocs(docsSheet, dealId) {
  const data = docsSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === dealId) docsSheet.deleteRow(i + 1);
  }
}

function updateDocStatus(docsSheet, dealId, pendentes) {
  const data = docsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === dealId && pendentes.includes(data[i][1])) {
      docsSheet.getRange(i + 1, 3).setValue('pendente_gerente');
    }
  }
}

function getDealsFromSheet(sheet, docsSheet) {
  const rows = sheet.getDataRange().getValues();
  const deals = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const dealId = String(row[1]);
    const docs = docsSheet ? getDealDocs(docsSheet, dealId) : [];
    const pendentes = docs.filter(d => d.status !== 'ok').map(d => d.doc);
    deals.push({
      id: dealId,
      dealId: dealId,
      vin: row[2], veiculo: row[3], vendedor: row[4],
      dataVenda: row[5],
      status: row[6],
      finance: row[7], lendbuzz: row[8], tag: row[9],
      parcelado: row[10], tradein: row[11], warranty: row[12],
      gap: row[13], rebuilt: row[14], obs: row[15],
      motivos: row[16] ? String(row[16]).split(', ') : [],
      obsGerente: row[17], reviewedAt: row[18],
      timestamp: row[0], docs, pendentes,
      history: row[19] ? JSON.parse(row[19]) : [],
      emplacamento: row[9],
      financiamento: row[7],
    });
  }
  return deals.reverse();
}

function getDealDocs(docsSheet, dealId) {
  const data = docsSheet.getDataRange().getValues();
  const docs = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === dealId) docs.push({ doc: data[i][1], status: data[i][2] });
  }
  return docs;
}
