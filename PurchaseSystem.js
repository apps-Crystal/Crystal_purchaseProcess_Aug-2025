/**
 * PurchaseSystem.gs
 * Core server-side functions:
 * - Header map helper
 * - COUNTERS based ID generation (nextSerial)
 * - PR create / approve / fetch
 * - PO create / approve / fetch
 * - Payment request / approve / post
 * - Dashboard data
 *
 * NOTE: This file expects the sheets created by setup script:
 * PR_Master, PR_Items, PO_Master, PO_Items, Payments, Item_Master, Vendor_Master, Approval_Matrix, COUNTERS, Audit_Log
 */

// ---- Helpers ----

function getHeaderMap(sheet) {
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h,i) => map[h] = i);
  console.log("Header Map:", map); // Add this line
  return map;
}

function safeGetData(sheet) {
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 2 || lc < 1) return [];
  const data = sheet.getRange(2,1,lr-1,lc).getValues();
  console.log("Data from safeGetData:", data); // Add this line
  return data;
}

function appendByHeader(sheetName, obj) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row = headers.map(h => (obj[h] !== undefined ? obj[h] : ''));
  sh.appendRow(row);
  return true;
}

// ---- COUNTERS / ID generation ----

function nextSerialCounter(key) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("COUNTERS");
  if (!sheet) throw new Error("COUNTERS sheet missing");
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = sheet.getDataRange().getValues();
    const map = {};
    for (let i=1;i<data.length;i++){
      if (!data[i] || !data[i][0]) continue;
      map[String(data[i][0])] = {row: i+1, val: Number(data[i][1]||0)};
    }
    if (!map[key]) {
      sheet.appendRow([key,1,new Date()]);
      return 1;
    } else {
      const row = map[key].row;
      const next = map[key].val + 1;
      sheet.getRange(row,2).setValue(next);
      sheet.getRange(row,3).setValue(new Date());
      return next;
    }
  } finally {
    try { lock.releaseLock(); } catch(e){}
  }
}

function formatSerial(prefix, siteOrYear, d, serial) {
  // prefix like PR or PO; siteOrYear may be site or year string
  return `${prefix}-${siteOrYear}-${d}-${String(serial).padStart(4,'0')}`;
}

// ---- Audit Log ----

function logAudit(entity, entityId, action, fromState, toState, remarks, payload) {
  appendByHeader("Audit_Log", {
    "Timestamp": new Date(),
    "Entity": entity,
    "Entity_ID": entityId,
    "Action": action,
    "From_State": fromState || '',
    "To_State": toState || '',
    "By": Session.getActiveUser().getEmail(),
    "Remarks": remarks || '',
    "Payload_JSON": JSON.stringify(payload || {})
  });
}

/* ---------- PR (Requisition) ---------- */

// createPR: accept an object with fields and items array
function createPR(payload) {
  // payload: { site, requestedBy, vendorId, purchaseCategory, paymentTerms, deliveryTerms, deliveryLocation, expectedDeliveryDate, items: [{Item_Code, Item_Name, Qty, UOM, Rate, GST_%, Purpose}] }
  const ss = SpreadsheetApp.getActive();
  const prSheet = ss.getSheetByName("PR_Master");
  const prItemsSheet = ss.getSheetByName("PR_Items");
  if (!prSheet || !prItemsSheet) throw new Error("PR_Master or PR_Items sheet missing");

  let vendorId = payload.vendorId || '';
  // If vendor is not registered, register now
  if (payload.vendorRegistered === "No" && payload.vendorDetails) {
    vendorId = registerVendor(payload.vendorDetails);
  }

  // generate PR_ID
  const site = payload.site || 'SITE';
  const yyyymm = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMM");
  const key = `PR:${site}:${yyyymm}`;
  const serial = nextSerialCounter(key);
  const prId = `PR-${site}-${yyyymm}-${String(serial).padStart(4,'0')}`;

  // compute totals
  let totalInclGST = 0;
  const items = payload.items || [];
  items.forEach((it, idx) => {
    const qty = Number(it.Qty||it.qty||0);
    const rate = Number(it.Rate||it.rate||0);
    const gst = Number(it['GST_%']||it.gst||0);
    const lineTotal = qty * rate * (1 + gst/100);
    totalInclGST += lineTotal;

    // append PR_Items row
    appendByHeader("PR_Items", {
      "PR_ID": prId,
      "Line_No": idx+1,
      "Item_Code": it.Item_Code || it.itemCode || '',
      "Item_Name": it.Item_Name || it.itemName || '',
      "Purpose": it.Purpose || it.purpose || '',
      "Qty": qty,
      "UOM": it.UOM || it.uom || '',
      "Rate": rate,
      "GST_%": gst,
      "Warranty_AMC": it.Warranty_AMC || '',
      "Line_Total": lineTotal
    });
  });

  // append PR_Master row
  const now = new Date();
  appendByHeader("PR_Master", {
    "PR_ID": prId,
    "Timestamp": now,
    "Date_of_Requisition": now,
    "Site": site,
    "Requested_By": payload.requestedBy || Session.getActiveUser().getEmail(),
    "Vendor_ID": vendorId,
    "Purchase_Category": payload.purchaseCategory || '',
    "Payment_Terms": payload.paymentTerms || '',
    "Delivery_Terms": payload.deliveryTerms || '',
    "Delivery_Location": payload.deliveryLocation || '',
    "Is_Vendor_Registered": payload.vendorRegistered || '',
    "Is_Customer_Reimbursable": payload.isCustomerReimbursable ? 'Yes' : 'No',
    "Total_Incl_GST": totalInclGST,
    "Status_Code": "PR_SUBMITTED",
    "Status_Label": "Submitted",
    "Last_Action_By": Session.getActiveUser().getEmail(),
    "Last_Action_At": now,
    "Approver_Remarks": "",
    "Approved_PR_Link": "",
    "PR_PDF_Link": "",
    "Approval_Link": "",
    "Expected_Delivery_Date": payload.expectedDeliveryDate || ""
  });

  logAudit("PR", prId, "CREATE", "", "PR_SUBMITTED", "", payload);
  return { success: true, prId: prId };
}

// approvePR / rejectPR
function approvePR(prId, remarks) {
  const ss = SpreadsheetApp.getActive();
  const prSheet = ss.getSheetByName("PR_Master");
  const headers = getHeaderMap(prSheet);
  const data = safeGetData(prSheet);
  const user = Session.getActiveUser().getEmail();
  const now = new Date();
  for (let i=0;i<data.length;i++){
    if (data[i][headers['PR_ID']] === prId) {
      const from = data[i][headers['Status_Code']];
      const to = 'PR_APPROVED';
      prSheet.getRange(i+2, headers['Status_Code']+1).setValue(to);
      prSheet.getRange(i+2, headers['Status_Label']+1).setValue('Approved');
      prSheet.getRange(i+2, headers['Last_Action_By']+1).setValue(user);
      prSheet.getRange(i+2, headers['Last_Action_At']+1).setValue(now);
      prSheet.getRange(i+2, headers['Approver_Remarks']+1).setValue(remarks || '');
      logAudit("PR", prId, "APPROVE", from, to, remarks || "");
      return { success: true, prId: prId };
    }
  }
  throw new Error("PR not found: " + prId);
}

function rejectPR(prId, remarks) {
  const ss = SpreadsheetApp.getActive();
  const prSheet = ss.getSheetByName("PR_Master");
  const headers = getHeaderMap(prSheet);
  const data = safeGetData(prSheet);
  const user = Session.getActiveUser().getEmail();
  const now = new Date();
  for (let i=0;i<data.length;i++){
    if (data[i][headers['PR_ID']] === prId) {
      const from = data[i][headers['Status_Code']];
      const to = 'PR_REJECTED';
      prSheet.getRange(i+2, headers['Status_Code']+1).setValue(to);
      prSheet.getRange(i+2, headers['Status_Label']+1).setValue('Rejected');
      prSheet.getRange(i+2, headers['Last_Action_By']+1).setValue(user);
      prSheet.getRange(i+2, headers['Last_Action_At']+1).setValue(now);
      prSheet.getRange(i+2, headers['Approver_Remarks']+1).setValue(remarks || '');
      logAudit("PR", prId, "REJECT", from, to, remarks || "");
      return { success: true, prId: prId };
    }
  }
  throw new Error("PR not found: " + prId);
}

/* ---------- PO ---------- */

// createPOFromPR: create a PO (posting Tally PO no / file) for an approved PR
function createPOFromPR(payload) {
  // payload: { prId, poNoTally, poDate, poPreparedBy, attachFileId, attachFileUrl }
  const ss = SpreadsheetApp.getActive();
  const prSheet = ss.getSheetByName("PR_Master");
  const poSheet = ss.getSheetByName("PO_Master");
  if (!prSheet || !poSheet) throw new Error("PR_Master or PO_Master missing");

  // verify PR state
  const prHeaders = getHeaderMap(prSheet);
  const prData = safeGetData(prSheet);
  let prRow = null;
  for (let i=0;i<prData.length;i++){
    if (prData[i][prHeaders['PR_ID']] === payload.prId) { prRow = {rowIndex: i+2, data: prData[i]}; break; }
  }
  if (!prRow) throw new Error("PR not found: " + payload.prId);
  if (prRow.data[prHeaders['Status_Code']] !== 'PR_APPROVED') throw new Error("PR must be APPROVED before creating PO");

  // generate PO_ID
  const site = prRow.data[prHeaders['Site']] || 'SITE';
  const yyyymm = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMM");
  const key = `PO:${site}:${yyyymm}`;
  const serial = nextSerialCounter(key);
  const poId = `PO-${site}-${yyyymm}-${String(serial).padStart(4,'0')}`;

  // copy PR totals
  const totalInclGST = prRow.data[prHeaders['Total_Incl_GST']] || 0;
  const now = new Date();

  appendByHeader("PO_Master", {
    "PO_ID": poId,
    "PR_ID": payload.prId,
    "Site": site,
    "Vendor_ID": prRow.data[prHeaders['Vendor_ID']] || '',
    "PO_No_Tally": payload.poNoTally || '',
    "PO_Date": payload.poDate || now,
    "PO_FileId": payload.attachFileId || '',
    "PO_File_URL": payload.attachFileUrl || '',
    "Total_Incl_GST": totalInclGST,
    "Status_Code": "PO_POSTED",
    "Status_Label": "Posted",
    "Last_Action_By": Session.getActiveUser().getEmail(),
    "Last_Action_At": now,
    "PO_Remarks": payload.remark || ''
  });

  // copy items from PR_Items to PO_Items
  const prItemsSheet = ss.getSheetByName("PR_Items");
  const poItemsSheet = ss.getSheetByName("PO_Items");
  const prItemsHeaders = getHeaderMap(prItemsSheet);
  const itemsData = safeGetData(prItemsSheet).filter(r => r[prItemsHeaders['PR_ID']] === payload.prId);
  itemsData.forEach((r, idx) => {
    appendByHeader("PO_Items", {
      "PO_ID": poId,
      "Line_No": idx+1,
      "Item_Code": r[prItemsHeaders['Item_Code']],
      "Item_Name": r[prItemsHeaders['Item_Name']],
      "Qty": r[prItemsHeaders['Qty']],
      "UOM": r[prItemsHeaders['UOM']],
      "Rate": r[prItemsHeaders['Rate']],
      "GST_%": r[prItemsHeaders['GST_%']],
      "Line_Total": r[prItemsHeaders['Line_Total']]
    });
  });

  // update PR status to indicate PO posted
  const prRowIndex = prRow.rowIndex;
  prSheet.getRange(prRowIndex, prHeaders['Status_Code']+1).setValue("PO_POSTED");
  prSheet.getRange(prRowIndex, prHeaders['Status_Label']+1).setValue("PO Posted");
  prSheet.getRange(prRowIndex, prHeaders['Last_Action_By']+1).setValue(Session.getActiveUser().getEmail());
  prSheet.getRange(prRowIndex, prHeaders['Last_Action_At']+1).setValue(now);

  logAudit("PO", poId, "CREATE_FROM_PR", "PR_APPROVED", "PO_POSTED", "", payload);
  return { success: true, poId: poId };
}

// processPOApproval: for approvers to approve/reject PO (keeps state)
function processPOApproval(poId, action, remarks) {
  const ss = SpreadsheetApp.getActive();
  const poSheet = ss.getSheetByName("PO_Master");
  if (!poSheet) throw new Error("PO_Master missing");
  const headers = getHeaderMap(poSheet);
  const data = safeGetData(poSheet);
  for (let i=0;i<data.length;i++){
    if (data[i][headers['PO_ID']] === poId) {
      const from = data[i][headers['Status_Code']];
      const to = action === 'Approved' ? 'PO_APPROVED' : 'PO_REJECTED';
      const label = action === 'Approved' ? 'Approved' : 'Rejected';
      poSheet.getRange(i+2, headers['Status_Code']+1).setValue(to);
      poSheet.getRange(i+2, headers['Status_Label']+1).setValue(label);
      poSheet.getRange(i+2, headers['Last_Action_By']+1).setValue(Session.getActiveUser().getEmail());
      poSheet.getRange(i+2, headers['Last_Action_At']+1).setValue(new Date());
      poSheet.getRange(i+2, headers['PO_Remarks']+1).setValue(remarks || '');
      logAudit("PO", poId, "APPROVAL", from, to, remarks || {});
      // also update parent PR Status_Label to reflect PO approval
      const prId = data[i][headers['PR_ID']];
      if (prId) {
        const prSheet = ss.getSheetByName("PR_Master");
        const prHeaders = getHeaderMap(prSheet);
        const prData = safeGetData(prSheet);
        for (let j=0;j<prData.length;j++){
          if (prData[j][prHeaders['PR_ID']] === prId) {
            prSheet.getRange(j+2, prHeaders['Status_Code']+1).setValue(to === 'PO_APPROVED' ? 'PO_APPROVED' : prData[j][prHeaders['Status_Code']]);
            prSheet.getRange(j+2, prHeaders['Status_Label']+1).setValue(to === 'PO_APPROVED' ? 'PO Approved' : prData[j][prHeaders['Status_Label']]);
            break;
          }
        }
      }
      return { success: true };
    }
  }
  throw new Error("PO not found: " + poId);
}

/* ---------- PAYMENTS ---------- */

// requestPayment: payment maker uploads voucher first (attachFileId/URL) and creates payment request
function requestPayment(payload) {
  // payload: { poId, trancheNo, amount, mode, utr, voucherFileId, voucherFileUrl, remarks }
  const ss = SpreadsheetApp.getActive();
  const poSheet = ss.getSheetByName("PO_Master");
  const paySheet = ss.getSheetByName("Payments");
  if (!poSheet || !paySheet) throw new Error("PO_Master or Payments sheet missing");

  const poHeaders = getHeaderMap(poSheet);
  const poData = safeGetData(poSheet);
  let poRow = null;
  for (let i=0;i<poData.length;i++){
    if (poData[i][poHeaders['PO_ID']] === payload.poId) { poRow = {rowIndex: i+2, data: poData[i]}; break; }
  }
  if (!poRow) throw new Error("PO not found: " + payload.poId);

  // ensure PO posted / approved? At least require PO_POSTED
  const poStatus = poRow.data[poHeaders['Status_Code']];
  if (poStatus !== 'PO_POSTED' && poStatus !== 'PO_APPROVED') throw new Error("PO must be posted before payment request");

  // generate PAY_ID
  const yyyymm = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMM");
  const key = `PAY:${yyyymm}`;
  const serial = nextSerialCounter(key);
  const payId = `PAY-${yyyymm}-${String(serial).padStart(4,'0')}`;
  const now = new Date();

  appendByHeader("Payments", {
    "PAY_ID": payId,
    "PO_ID": payload.poId,
    "Tranche_No": payload.trancheNo || 1,
    "Amount": Number(payload.amount || 0),
    "Payment_Voucher_FileId": payload.voucherFileId || '',
    "Payment_Voucher_URL": payload.voucherFileUrl || '',
    "Status_Code": "VOUCHER_UPLOADED",
    "Status_Label": "Voucher Uploaded",
    "Mode": payload.mode || '',
    "UTR": payload.utr || '',
    "Posted_Date": '',
    "Remarks": payload.remarks || '',
    "Last_Action_By": Session.getActiveUser().getEmail(),
    "Last_Action_At": now
  });

  logAudit("PAYMENT", payId, "REQUESTED", "", "VOUCHER_UPLOADED", "", payload);
  return { success: true, payId: payId };
}

// approvePayment by Director
function approvePayment(payId, action, remarks) {
  const ss = SpreadsheetApp.getActive();
  const paySheet = ss.getSheetByName("Payments");
  if (!paySheet) throw new Error("Payments sheet missing");
  const headers = getHeaderMap(paySheet);
  const data = safeGetData(paySheet);
  for (let i=0;i<data.length;i++){
    if (data[i][headers['PAY_ID']] === payId) {
      const from = data[i][headers['Status_Code']];
      const to = action === 'Approved' ? 'DIRECTOR_OK' : 'PAY_REJECTED';
      const label = action === 'Approved' ? 'Director Approved' : 'Rejected';
      paySheet.getRange(i+2, headers['Status_Code']+1).setValue(to);
      paySheet.getRange(i+2, headers['Status_Label']+1).setValue(label);
      paySheet.getRange(i+2, headers['Last_Action_By']+1).setValue(Session.getActiveUser().getEmail());
      paySheet.getRange(i+2, headers['Last_Action_At']+1).setValue(new Date());
      if (remarks) paySheet.getRange(i+2, headers['Remarks']+1).setValue(remarks);
      logAudit("PAYMENT", payId, "DIRECTOR_APPROVAL", from, to, remarks || "");
      return { success: true };
    }
  }
  throw new Error("PAY not found: " + payId);
}

// postPayment: after director ok, payment maker posts UTR/date and marks PAID
function postPayment(payId, postedDate, utr, remarks) {
  const ss = SpreadsheetApp.getActive();
  const paySheet = ss.getSheetByName("Payments");
  if (!paySheet) throw new Error("Payments sheet missing");
  const headers = getHeaderMap(paySheet);
  const data = safeGetData(paySheet);
  for (let i=0;i<data.length;i++){
    if (data[i][headers['PAY_ID']] === payId) {
      const status = data[i][headers['Status_Code']];
      if (status !== 'DIRECTOR_OK') throw new Error("Payment must be DIRECTOR_OK before posting");
      paySheet.getRange(i+2, headers['Status_Code']+1).setValue('PAY_POSTED');
      paySheet.getRange(i+2, headers['Status_Label']+1).setValue('Paid');
      paySheet.getRange(i+2, headers['Posted_Date']+1).setValue(postedDate || new Date());
      if (utr) paySheet.getRange(i+2, headers['UTR']+1).setValue(utr);
      if (remarks) paySheet.getRange(i+2, headers['Remarks']+1).setValue(remarks);
      paySheet.getRange(i+2, headers['Last_Action_By']+1).setValue(Session.getActiveUser().getEmail());
      paySheet.getRange(i+2, headers['Last_Action_At']+1).setValue(new Date());
      logAudit("PAYMENT", payId, "POSTED", status, "PAY_POSTED", remarks || "");
      return { success: true };
    }
  }
  throw new Error("PAY not found: " + payId);
}

/* ---------- Fetchers / Dashboard ---------- */

function getDashboardData() {
  // returns compact stats for front-end dashboard
  const ss = SpreadsheetApp.getActive();
  const prSheet = ss.getSheetByName("PR_Master");
  const poSheet = ss.getSheetByName("PO_Master");
  const payments = ss.getSheetByName("Payments");

  const result = {
    totalPR:0, pendingPR:0, approvedPR:0,
    totalPO:0, postedPO:0, approvedPO:0,
    totalPayments:0, paymentsPendingDirector:0, paymentsPaid:0,
    recentPRs: []
  };

  if (prSheet) {
    const prData = safeGetData(prSheet);
    const h = getHeaderMap(prSheet);
    result.totalPR = prData.length;
    prData.forEach(r => {
      const st = r[h['Status_Code']];
      if (st === 'PR_SUBMITTED') result.pendingPR++;
      if (st === 'PR_APPROVED') result.approvedPR++;
    });
    // last 5 PRs
    const last5 = prData.slice(-5).reverse();
    last5.forEach(r => {
      result.recentPRs.push({
        id: r[h['PR_ID']],
        site: r[h['Site']],
        total: r[h['Total_Incl_GST']],
        status: r[h['Status_Label']],
        date: r[h['Date_of_Requisition']]
      });
    });
  }

  if (poSheet) {
    const poData = safeGetData(poSheet);
    const h = getHeaderMap(poSheet);
    result.totalPO = poData.length;
    poData.forEach(r => {
      const st = r[h['Status_Code']];
      if (st === 'PO_POSTED') result.postedPO++;
      if (st === 'PO_APPROVED') result.approvedPO++;
    });
  }

  if (payments) {
    const pData = safeGetData(payments);
    const h = getHeaderMap(payments);
    result.totalPayments = pData.length;
    pData.forEach(r => {
      const st = r[h['Status_Code']];
      if (st === 'VOUCHER_UPLOADED') result.paymentsPendingDirector++;
      if (st === 'PAY_POSTED') result.paymentsPaid++;
    });
  }

  return result;
}

// ---------- Utility: getApprovedPRs for PO form ----------
function getApprovedPRsList() {
  const ss = SpreadsheetApp.getActive();
  const prSheet = ss.getSheetByName("PR_Master");
  if (!prSheet) return [];
  const data = safeGetData(prSheet);
  const h = getHeaderMap(prSheet);
  return data.filter(r => r[h['Status_Code']]==='PR_APPROVED').map(r => r[h['PR_ID']]);
}

// ---------- get requisition details for PO form ----------
function getRequisitionDetailsForPO(prId) {
  if (!prId) return null;
  const ss = SpreadsheetApp.getActive();
  const prSheet = ss.getSheetByName("PR_Master");
  const prItemsSheet = ss.getSheetByName("PR_Items");
  const vendorSheet = ss.getSheetByName("Vendor_Master");
  if (!prSheet || !prItemsSheet) return null;
  const prData = safeGetData(prSheet);
  const prH = getHeaderMap(prSheet);
  const itemsData = safeGetData(prItemsSheet);
  const itemsH = getHeaderMap(prItemsSheet);
  let prRow = null;
  for (let i=0;i<prData.length;i++) if (prData[i][prH['PR_ID']]===prId) { prRow = prData[i]; break; }
  if (!prRow) return null;
  const vendorId = prRow[prH['Vendor_ID']];
  let vendor = null;
  if (vendorId && vendorSheet) {
    const vData = safeGetData(vendorSheet);
    const vH = getHeaderMap(vendorSheet);
    for (let i=0;i<vData.length;i++) if (vData[i][vH['Vendor_ID']] === vendorId) {
      vendor = {};
      Object.keys(vH).forEach(k => vendor[k] = vData[i][vH[k]]);
      break;
    }
  }
  const items = itemsData.filter(r => r[itemsH['PR_ID']] === prId).map(r => ({
    itemName: r[itemsH['Item_Name']],
    purpose: r[itemsH['Purpose']],
    quantity: r[itemsH['Qty']],
    uom: r[itemsH['UOM']],
    rate: r[itemsH['Rate']],
    gst: r[itemsH['GST_%']],
    totalCost: r[itemsH['Line_Total']]
  }));
  return {
    paymentTerms: prRow[prH['Payment_Terms']],
    deliveryTerms: prRow[prH['Delivery_Terms']],
    expectedDeliveryDate: prRow[prH['Expected_Delivery_Date']],
    site: prRow[prH['Site']],
    isVendorRegistered: prRow[prH['Is_Vendor_Registered']],
    vendor: vendor,
    items: items
  };
}

// returns master lists used on forms
function getMasterDataForForm() {
  const ss = SpreadsheetApp.getActive();
  const masterSheet = ss.getSheetByName("Master_Data");
  if (!masterSheet) {
    console.error("Master_Data sheet not found");
    return {};
  }

  const data = masterSheet.getDataRange().getValues();
  const headers = data.shift(); // Get and remove header row

  const masterData = {};
  headers.forEach((header, index) => {
    if (header) {
      // Get all values for the column, filter out empty cells
      const columnValues = data.map(row => row[index]).filter(String);
      const uniqueColumnValues = [...new Set(columnValues)]; // Remove duplicates using a Set
      masterData[header] = uniqueColumnValues;
    }
  });

  console.log("Master Data:", masterData); // Add this line
  return masterData;
}

function createVendor(vendorPayload) {
  // Logic to add a new vendor to the Vendor_Master sheet
  const ss = SpreadsheetApp.getActive();
  const vendorSheet = ss.getSheetByName("Vendor_Master");
  if (!vendorSheet) throw new Error("Vendor_Master sheet missing");

  // Generate a new Vendor_ID
  const yyyymm = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMM");
  const key = `VND:${yyyymm}`;
  const serial = nextSerialCounter(key);
  const vendorId = `VND-${yyyymm}-${String(serial).padStart(4,'0')}`;
  
  // Create a new vendor object
  const newVendor = {
  "Vendor_ID": vendorId,
  "Company_Name": vendorPayload.vendorName,
  "Contact_Person": vendorPayload.contactPerson,
  "Contact_Number": vendorPayload.phoneNumber,
  "Email_ID": vendorPayload.email,
  "Bank_Name": vendorPayload.bankName,
  "Acc_Holder_Name": vendorPayload.accountHolderName,
  "Acc_Number": vendorPayload.accountNo,
  "Branch_Name": vendorPayload.branchName,
  "IFSC_CODE": vendorPayload.ifscCode,
  "GST_Number": vendorPayload.gstNo,
  "Providing_Sites": vendorPayload.providingSites,
  "Vendor_PAN": vendorPayload.panNo,
  "Vendor_Address": vendorPayload.address,
  "Active": "Yes",
  "Created_At": new Date(),
  "Created_By": Session.getActiveUser().getEmail()
  };

  appendByHeader("Vendor_Master", newVendor);
  logAudit("VENDOR", vendorId, "CREATE", "", "REGISTERED", "", newVendor);
  return vendorId;
}

/**
 * Returns all vendors as array of objects for the form dropdown/autofill.
 */
// function getVendorMasterList() {
//   try {
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const sheet = ss.getSheetByName("Vendor_Master");
//     if (!sheet) {
//       console.error("Vendor_Master sheet not found.");
//       return [];
//     }

//     const headers = getHeaderMap(sheet);
//     console.log("Headers Map:", headers);

//     const data = safeGetData(sheet);
//     console.log("Data from Sheet:", data); // Add this line

//     return data.map(row => {
//       const vendor = {};
//       Object.keys(headers).forEach(key => {
//         vendor[key] = row[headers[key]];
//       });
//       return vendor;
//     });
//   } catch (e) {
//     console.error("Error in getVendorMasterList: " + e.toString());
//     return null; // Return null in case of error
//   }
// }
// function getVendorMasterList() {
//   try {
//     const ss = SpreadsheetApp.getActive();
//     const sheet = ss.getSheetByName("Vendor_Master");
//     if(sheet){
//       console.log("Vendor_Master sheet found.");
//     }
//     if (!sheet) return [];
//     const data = sheet.getDataRange().getValues();
//     const headers = data[0];
//     return data.slice(1)
//       .filter(row => row[headers.indexOf("Active")] === "Yes")
//       .map(row => {
//         let obj = {};
//         headers.forEach((h, i) => obj[h] = row[i]);
//         return obj;
//       });
//   } catch (e) {
//     Logger.log("Error in getVendorMasterList: " + e);
//     return [];
//   }
// }
function getVendorMasterList() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName("Vendor_Master");
    if (!sheet) {
      console.error("Vendor_Master sheet not found.");
      return [];
    }
    const data = sheet.getDataRange().getValues();
    if (!data.length) {
      console.error("Vendor_Master sheet is empty.");
      return [];
    }
    const headers = data[0];
    const activeIdx = headers.indexOf("Active");
    if (activeIdx === -1) {
      console.error('"Active" column not found in Vendor_Master.');
      return [];
    }
    return data.slice(1)
      .filter(row => row[activeIdx] === "Yes")
      .map(row => {
        let obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
  } catch (e) {
    console.error("Error in getVendorMasterList: " + e.toString());
    return [];
  }
}

/**
 * Registers a new vendor in Vendor_Master and returns the new Vendor_ID.
 * Expects vendorDetails object from the form.
 */
function registerVendor(vendorDetails) {
  try {
    const ss = SpreadsheetApp.getActive();
    const vendorSheet = ss.getSheetByName("Vendor_Master");
    if (!vendorSheet) {
      throw new Error("Vendor_Master sheet not found.");
    }

    // Generate Vendor_ID
    const yyyymm = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMM");
    const key = `VENDOR:${yyyymm}`;
    const serial = nextSerialCounter(key);
    const vendorId = `V-${yyyymm}-${String(serial).padStart(4,'0')}`;

    const newVendorData = {
      ...vendorDetails,
      "Vendor_ID": vendorId,
      "Active": "Yes",
      "Created_At": new Date(),
      "Created_By": Session.getActiveUser().getEmail()
    };

    appendByHeader("Vendor_Master", newVendorData);

    logAudit("VENDOR", vendorId, "CREATE", "", "ACTIVE", "New vendor registered via form.", newVendorData);

    return { success: true, vendorId: vendorId };
  } catch (e) {
    console.error("Error in registerVendor: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

