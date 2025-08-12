/**
 * Run this once to create the workbook structure:
 * - PR_Master, PR_Items, PO_Master, PO_Items, Payments,
 *   Item_Master, Vendor_Master, Approval_Matrix, COUNTERS, Audit_Log, Master_Data
 * - Headers, frozen header row, and basic dropdown validations (using Master_Data lists)
 *
 * After run: edit Master_Data values (Sites, UOMs, Categories, etc.) to your real org values.
 */

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActive();
  // Define sheet names and headers
  const sheets = {
    "PR_Master": [
      "PR_ID","Timestamp","Date_of_Requisition","Site","Requested_By","Vendor_ID",
      "Purchase_Category","Payment_Terms","Delivery_Terms","Delivery_Location",
      "Is_Vendor_Registered","Is_Customer_Reimbursable","Total_Incl_GST",
      "Status_Code","Status_Label","Last_Action_By","Last_Action_At","Approver_Remarks",
      "Approved_PR_Link","PR_PDF_Link","Approval_Link","Expected_Delivery_Date"
    ],
    "PR_Items": [
      "PR_ID","Line_No","Item_Code","Item_Name","Purpose","Qty","UOM","Rate","GST_%","Warranty_AMC","Line_Total"
    ],
    "PO_Master": [
      "PO_ID","PR_ID","Site","Vendor_ID","PO_No_Tally","PO_Date","PO_FileId","PO_File_URL",
      "Total_Incl_GST","Status_Code","Status_Label","Last_Action_By","Last_Action_At","PO_Remarks"
    ],
    "PO_Items": [
      "PO_ID","Line_No","Item_Code","Item_Name","Qty","UOM","Rate","GST_%","Line_Total"
    ],
    "Payments": [
      "PAY_ID","PO_ID","Tranche_No","Amount","Payment_Voucher_FileId","Payment_Voucher_URL",
      "Status_Code","Status_Label","Mode","UTR","Posted_Date","Remarks","Last_Action_By","Last_Action_At"
    ],
    "Item_Master": [
      "Item_Code","Item_Name","Category","UOM","Default_Vendor","Active"
    ],
    "Vendor_Master": [
      "Vendor_ID","Company_Name","Contact_Person","Contact_Number","Email_ID","Bank_Name",
      "Acc_Holder_Name","Acc_Number","Branch_Name","IFSC_CODE","GST_Number","Providing_Sites",
      "Vendor_PAN","Vendor_Address","GST_Certificate_FileId","PanCard_FileId","Cancelled_Cheque_FileId",
      "Active","Created_At","Created_By"
    ],
    "Approval_Matrix": [
      "Role","Min_Amount","Max_Amount","Levels_Required","Applies_To","Fallback_Approver"
    ],
    "COUNTERS": [
      "Key","LastSerial","UpdatedAt"
    ],
    "Audit_Log": [
      "Timestamp","Entity","Entity_ID","Action","From_State","To_State","By","Remarks","Payload_JSON"
    ],
    // Master_Data holds dropdown lists used across sheets
    "Master_Data": [
      "Sites","UOMs","Purchase_Categories","Payment_Terms","Delivery_Terms",
      "PR_Statuses","PO_Statuses","PAY_Statuses","Payment_Modes","Yes_No"
    ]
  };

  // Create or reset sheets and headers
  for (const [name, headers] of Object.entries(sheets)) {
    ensureSheetWithHeaders_(ss, name, headers);
  }

  // Put sample values in Master_Data below each header (row 2+)
  populateMasterDataDefaults_(ss.getSheetByName("Master_Data"));

  // Apply data validations that reference Master_Data lists
  applyValidations_(ss);

  // Freeze header rows, set reasonable column widths
  for (const name of Object.keys(sheets)) {
    const sh = ss.getSheetByName(name);
    if (sh) {
      sh.setFrozenRows(1);
      sh.setTabColor(name === "COUNTERS" ? "#f4cccc" : null);
      // set columns to auto-resize a bit
      try { sh.autoResizeColumns(1, Math.min(20, sh.getLastColumn())); } catch(e) {}
    }
  }

  // Add example COUNTERS rows (if not present)
  const counterSheet = ss.getSheetByName("COUNTERS");
  const existing = counterSheet.getRange(2,1,counterSheet.getLastRow()-1 || 0,1).getValues().flat().filter(String);
  const defaults = ["PR:<SITE>:<YYYYMM>","PO:<SITE>:<YYYYMM>","VENDOR:<YYYY>","PAY:<YYYYMM>"];
  for (let key of defaults) {
    if (existing.indexOf(key) === -1) {
      counterSheet.appendRow([key,0,new Date()]);
    }
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("Spreadsheet setup complete.\n\nNow open 'Master_Data' sheet and replace sample values (Sites, UOMs, etc.) with your actual lists.");
}

/* ----------------- Helpers ----------------- */

function ensureSheetWithHeaders_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  } else {
    // preserve sheet, but clear contents except headers row will be overwritten below
    sh.clearContents();
  }
  // set headers
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  // format header row (bold + background)
  sh.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#f2f2f2");
}

function populateMasterDataDefaults_(masterSheet) {
  // headers are in row 1; write sample list entries starting row 2 in each column
  const sample = {
    "Sites": ["SiteA","SiteB","HeadOffice"],
    "UOMs": ["Nos","Kg","Mtr","Litre"],
    "Purchase_Categories": ["Asset","Consumable","Service"],
    "Payment_Terms": ["Advance","Net 15","Net 30","On Delivery"],
    "Delivery_Terms": ["Door Delivery","Ex-Works","FOB"],
    "PR_Statuses": ["PR_SUBMITTED","PR_APPROVED","PR_REJECTED"],
    "PO_Statuses": ["PO_POSTED"],
    "PAY_Statuses": ["VOUCHER_UPLOADED","DIRECTOR_OK","PAY_POSTED","PAY_REJECTED"],
    "Payment_Modes": ["NEFT","RTGS","UPI","Cheque","Cash"],
    "Yes_No": ["Yes","No"]
  };

  const headers = masterSheet.getRange(1,1,1,masterSheet.getLastColumn()).getValues()[0];
  for (let c = 0; c < headers.length; c++) {
    const h = headers[c];
    const values = sample[h] || [];
    const writeRange = masterSheet.getRange(2, c+1, Math.max(1, values.length), 1);
    if (values.length) {
      masterSheet.getRange(2, c+1, values.length, 1).setValues(values.map(v => [v]));
    } else {
      // leave empty cell so user can fill
      masterSheet.getRange(2, c+1).setValue("");
    }
  }
  // widen columns a bit
  try { masterSheet.autoResizeColumns(1, masterSheet.getLastColumn()); } catch(e){}
}

function applyValidations_(ss) {
  const md = ss.getSheetByName("Master_Data");
  if (!md) return;

  // helper to get range of items under a given header in Master_Data (col header in row1)
  function listRange(header) {
    const headers = md.getRange(1,1,1,md.getLastColumn()).getValues()[0];
    const idx = headers.indexOf(header);
    if (idx === -1) return null;
    const col = idx + 1;
    // find last non-empty row in that column
    const values = md.getRange(2, col, md.getMaxRows()-1, 1).getValues().map(r=>r[0]);
    let last = 1;
    for (let i=0;i<values.length;i++){
      if (values[i] !== "" && values[i] !== null && values[i] !== undefined) last = i+2;
    }
    // always return at least row2..row10 to support validations even if few values
    const endRow = Math.max(last, 10);
    return md.getRange(2, col, endRow-1, 1);
  }

  const validations = [
    {sheet: "PR_Master", colName: "Site", listHeader: "Sites"},
    {sheet: "PR_Master", colName: "Purchase_Category", listHeader: "Purchase_Categories"},
    {sheet: "PR_Master", colName: "Payment_Terms", listHeader: "Payment_Terms"},
    {sheet: "PR_Master", colName: "Delivery_Terms", listHeader: "Delivery_Terms"},
    {sheet: "PR_Master", colName: "Is_Vendor_Registered", listHeader: "Yes_No"},
    {sheet: "PR_Master", colName: "Is_Customer_Reimbursable", listHeader: "Yes_No"},
    {sheet: "PR_Master", colName: "Status_Code", listHeader: "PR_Statuses"},
    {sheet: "PR_Items", colName: "UOM", listHeader: "UOMs"},
    {sheet: "PR_Items", colName: "Item_Code", listHeader: "Item_Master"}, // Item_Master used differently; allow free text too
    {sheet: "PO_Master", colName: "Site", listHeader: "Sites"},
    {sheet: "PO_Master", colName: "Status_Code", listHeader: "PO_Statuses"},
    {sheet: "Payments", colName: "Status_Code", listHeader: "PAY_Statuses"},
    {sheet: "Payments", colName: "Mode", listHeader: "Payment_Modes"},
    {sheet: "Item_Master", colName: "Category", listHeader: "Purchase_Categories"},
    {sheet: "Item_Master", colName: "UOM", listHeader: "UOMs"},
    {sheet: "Vendor_Master", colName: "Providing_Sites", listHeader: "Sites"},
    {sheet: "PR_Master", colName: "Requested_By", listHeader: null} // kept as free text (email)
  ];

  // A special range for Item_Master codes (we'll use the Item_Master sheet's Item_Code column)
  const itemMasterSheet = ss.getSheetByName("Item_Master");
  const itemCodeRange = itemMasterSheet.getRange(2,1,Math.max(10, itemMasterSheet.getMaxRows()-1),1);

  validations.forEach(v => {
    const sh = ss.getSheetByName(v.sheet);
    if (!sh) return;
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const idx = headers.indexOf(v.colName);
    if (idx === -1) return;
    const col = idx + 1;
    const lastRow = Math.max(1000, sh.getMaxRows());
    const range = sh.getRange(2, col, lastRow-1, 1);
    let rule = null;
    if (v.listHeader === "Item_Master") {
      rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(itemCodeRange, true)
        .setAllowInvalid(false)
        .build();
    } else if (v.listHeader) {
      const listR = listRange(v.listHeader);
      if (listR) {
        rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(listR, true)
          .setAllowInvalid(false)
          .build();
      }
    } else {
      // no validation (free text)
      rule = null;
    }
    try {
      if (rule) range.setDataValidation(rule);
      else range.clearDataValidations();
    } catch (e) {
      // ignore if column out of bounds
      Logger.log("Validation skipped for " + v.sheet + "." + v.colName + " : " + e.message);
    }
  });

  // Additionally, set number format for amount fields
  const pr = ss.getSheetByName("PR_Master");
  setNumberFormatIfExists_(pr, "Total_Incl_GST", "₹#,##0.00");
  const po = ss.getSheetByName("PO_Master");
  setNumberFormatIfExists_(po, "Total_Incl_GST", "₹#,##0.00");
  const payments = ss.getSheetByName("Payments");
  setNumberFormatIfExists_(payments, "Amount", "₹#,##0.00");
}

function setNumberFormatIfExists_(sheet, colName, format) {
  if (!sheet) return;
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idx = headers.indexOf(colName);
  if (idx === -1) return;
  const col = idx+1;
  sheet.getRange(2, col, sheet.getMaxRows()-1, 1).setNumberFormat(format);
}

/* Optional helper: generate ID with COUNTERS (example usage from your code later)
   This function is provided for reference and is safe to call from other scripts.
*/
function nextSerialCounter(key) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("COUNTERS");
  if (!sheet) throw new Error("COUNTERS sheet required");
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let map = {};
    for (let i=1;i<data.length;i++){
      const r = data[i];
      if (!r || !r[0]) continue;
      map[String(r[0])] = {row: i+1, val: Number(r[1]||0)};
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
