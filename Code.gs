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

    const masterSheet = ensureSheet(ss, 'DEAL_MASTER', [
      'deal_id', 'stock', 'vin', 'vehicle', 'seller', 'customer_name', 'sale_date',
      'lender', 'deal_type', 'trade_in', 'gps_required', 'parcelado', 'out_of_state',
      'company_deal', 'review_status', 'commission_status', 'dmv_status',
      'parcelamento_status', 'title_status', 'omnia_status', 'folder_status',
      'envelope_status', 'created_at', 'updated_at'
    ]);
    const dealDocumentsSheet = ensureSheet(ss, 'DEAL_DOCUMENTS', [
      'deal_id', 'document_name', 'printed_status', 'dealer_center_status', 'status',
      'created_at', 'updated_at'
    ]);
    const historySheet = ensureSheet(ss, 'DEAL_HISTORY', [
      'deal_id', 'action', 'user', 'notes', 'timestamp'
    ]);
    const alertsSheet = ensureSheet(ss, 'DEAL_ALERTS', [
      'deal_id', 'alert_type', 'priority', 'status', 'created_at', 'resolved_at'
    ]);

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

      upsertDealMaster(masterSheet, data);
      syncDealDocuments(dealDocumentsSheet, data.id, data.docs || []);
      appendDealHistory(
        historySheet,
        data.id,
        existingRow > 0 ? 'submit_update' : 'submit',
        data.vendedor || '',
        existingRow > 0 ? 'Conferência atualizada pelo vendedor' : 'Conferência enviada pelo vendedor'
      );
      createInitialAlerts(alertsSheet, data);

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
        appendDealHistory(historySheet, data.id, 'correction', '', data.obs || 'Correção enviada pelo vendedor');
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

        updateDealMasterReviewStatus(masterSheet, data.id, data.status);
        appendDealHistory(historySheet, data.id, 'review', data.user || 'Giovanna', data.obsGerente || (data.motivos || []).join(', '));

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

function ensureSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
  headers.forEach((header, index) => {
    if (currentHeaders[index] !== header) {
      sheet.getRange(1, index + 1).setValue(header);
    }
  });
  sheet.setFrozenRows(1);
  return sheet;
}

function findDealRow(sheet, dealId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(dealId)) return i + 1;
  }
  return -1;
}

function findRowByFirstColumn(sheet, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(value)) return i + 1;
  }
  return -1;
}

function clearDealDocs(docsSheet, dealId) {
  const data = docsSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(dealId)) docsSheet.deleteRow(i + 1);
  }
}

function updateDocStatus(docsSheet, dealId, pendentes) {
  const data = docsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(dealId) && pendentes.includes(data[i][1])) {
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
      history: parseHistory(row[19]),
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
    if (String(data[i][0]) === String(dealId)) docs.push({ doc: data[i][1], status: data[i][2] });
  }
  return docs;
}

function upsertDealMaster(masterSheet, data) {
  const now = new Date().toISOString();
  const row = findRowByFirstColumn(masterSheet, data.id);
  const existing = row > 0 ? masterSheet.getRange(row, 1, 1, 24).getValues()[0] : [];
  const createdAt = existing[22] || now;
  const parcelado = normalizeYesNo(data.parcelado);

  const rowData = [
    data.id || '',
    data.stock || data.stockNumber || data.stock_number || '',
    data.vin || '',
    data.veiculo || data.vehicle || '',
    data.vendedor || data.seller || '',
    data.customerName || data.customer_name || data.cliente || '',
    data.dataVenda || data.sale_date || '',
    getLender(data),
    data.finance || data.deal_type || '',
    normalizeYesNo(data.tradein),
    parcelado === 'sim' ? 'sim' : normalizeYesNo(data.gps_required),
    parcelado,
    normalizeYesNo(data.out_of_state),
    normalizeYesNo(data.company_deal || data.companyDeal) || inferCompanyDeal(data.docs || []),
    existing[14] || 'pending',
    existing[15] || 'pending',
    existing[16] || 'pending',
    parcelado === 'sim' ? (existing[17] || 'pending') : (existing[17] || ''),
    existing[18] || '',
    existing[19] || '',
    existing[20] || '',
    existing[21] || '',
    createdAt,
    now
  ];

  if (row > 0) {
    masterSheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
  } else {
    masterSheet.appendRow(rowData);
  }
}

function updateDealMasterReviewStatus(masterSheet, dealId, status) {
  const row = findRowByFirstColumn(masterSheet, dealId);
  if (row < 1) return;
  const reviewStatus = status === 'aprovado' ? 'approved' : status === 'bloqueado' ? 'blocked' : status || 'pending';
  masterSheet.getRange(row, 15).setValue(reviewStatus);
  masterSheet.getRange(row, 24).setValue(new Date().toISOString());
}

function syncDealDocuments(dealDocumentsSheet, dealId, docs) {
  clearDealDocs(dealDocumentsSheet, dealId);
  const now = new Date().toISOString();
  const grouped = {};

  docs.forEach(doc => {
    const parsed = parseDocumentName(doc.doc);
    if (!grouped[parsed.name]) {
      grouped[parsed.name] = {
        printed_status: '',
        dealer_center_status: '',
        status: '',
      };
    }

    if (parsed.kind === 'printed') {
      grouped[parsed.name].printed_status = doc.status || '';
    } else if (parsed.kind === 'dealer_center') {
      grouped[parsed.name].dealer_center_status = doc.status || '';
    } else {
      grouped[parsed.name].status = doc.status || '';
    }
  });

  Object.keys(grouped).forEach(name => {
    const item = grouped[name];
    const status = item.status || aggregateDocumentStatus(item.printed_status, item.dealer_center_status);
    dealDocumentsSheet.appendRow([
      dealId,
      name,
      item.printed_status,
      item.dealer_center_status,
      status,
      now,
      now
    ]);
  });
}

function parseDocumentName(name) {
  const text = String(name || '');
  if (text.indexOf('— Impresso') > -1 || text.indexOf('- Impresso') > -1) {
    return { name: text.replace(/\s*[—-]\s*Impresso\s*$/, ''), kind: 'printed' };
  }
  if (text.indexOf('— DC') > -1 || text.indexOf('- DC') > -1) {
    return { name: text.replace(/\s*[—-]\s*DC\s*$/, ''), kind: 'dealer_center' };
  }
  return { name: text, kind: 'general' };
}

function aggregateDocumentStatus(printedStatus, dealerCenterStatus) {
  const statuses = [printedStatus, dealerCenterStatus].filter(Boolean);
  if (statuses.length === 0) return '';
  return statuses.every(status => status === 'ok') ? 'ok' : 'pendente';
}

function appendDealHistory(historySheet, dealId, action, user, notes) {
  historySheet.appendRow([
    dealId || '',
    action || '',
    user || '',
    notes || '',
    new Date().toISOString()
  ]);
}

function createInitialAlerts(alertsSheet, data) {
  if (normalizeYesNo(data.parcelado) === 'sim') {
    upsertOpenAlert(alertsSheet, data.id, 'Parcelamento não criado', 'high');
  }
  upsertOpenAlert(alertsSheet, data.id, 'DMV não enviado', 'medium');
}

function upsertOpenAlert(alertsSheet, dealId, alertType, priority) {
  const data = alertsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const sameDeal = String(data[i][0]) === String(dealId);
    const sameAlert = String(data[i][1]) === String(alertType);
    const stillOpen = !data[i][5] && String(data[i][3] || 'open') !== 'resolved';
    if (sameDeal && sameAlert && stillOpen) return;
  }
  alertsSheet.appendRow([
    dealId || '',
    alertType,
    priority || 'medium',
    'open',
    new Date().toISOString(),
    ''
  ]);
}

function normalizeYesNo(value) {
  const text = String(value || '').trim().toLowerCase();
  if (['sim', 'yes', 'true', '1'].includes(text)) return 'sim';
  if (['nao', 'não', 'no', 'false', '0'].includes(text)) return 'nao';
  return text;
}

function inferCompanyDeal(docs) {
  const companyDocs = docs.some(doc => /sunbiz|empresa|company|poa em nome da empresa|certificate of title/i.test(String(doc.doc || '')));
  return companyDocs ? 'sim' : '';
}

function getLender(data) {
  if (data.lender) return data.lender;
  if (normalizeYesNo(data.lendbuzz) === 'sim') return 'Lendbuzz';
  return '';
}

function parseHistory(value) {
  if (!value) return [];
  try {
    return JSON.parse(value);
  } catch (e) {
    return [];
  }
}
