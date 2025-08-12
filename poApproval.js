// poApproval.gs
function poApprovalPage() {
  const t = HtmlService.createTemplateFromFile('poApprovl');
  const pending = getPendingPOsForUI();
  t.pending = pending;
  return t.evaluate().getContent();
}

function getPendingPOsForUI() {
  const ss = SpreadsheetApp.getActive();
  const poSheet = ss.getSheetByName("PO_Master");
  if (!poSheet) return [];
  const poData = safeGetData(poSheet);
  const ph = getHeaderMap(poSheet);

  const pendingPOs = poData
    .filter(row => row[ph['Status_Code']] === 'PO_POSTED')
    .map(row => ({
      poId: row[ph['PO_ID']],
      prId: row[ph['PR_ID']] || 'N/A', // Ensure prId is not null
      site: row[ph['Site']],
      totalValue: row[ph['Total_Incl_GST']],
      date: row[ph['PO_Date']],
      attach: row[ph['PO_File_URL']]
    }));

  console.log(`Pending POs: ${JSON.stringify(pendingPOs)}`);
  return pendingPOs;
}

function decide(poId, action) {
  if (!poId || poId.trim() === '') {
    alert("Invalid PO ID.");
    return;
  }
  const remarks = action === 'Rejected' ? prompt("Remarks (required)") : prompt("Any remarks?");
  if (action === 'Rejected' && !remarks) return alert("Remarks required");
  google.script.run
    .withSuccessHandler(function () {
      document.getElementById('po-' + poId).remove();
    })
    .withFailureHandler(function (e) {
      alert(e.message);
    })
    .processPOApproval(poId, action, remarks);
}

function processPOApproval(poId, action, remarks) {
  console.log(`Processing PO Approval: PO ID = ${poId}, Action = ${action}, Remarks = ${remarks}`);
  const ss = SpreadsheetApp.getActive();
  const poSheet = ss.getSheetByName("PO_Master");
  if (!poSheet) throw new Error("PO_Master sheet not found.");

  const poData = safeGetData(poSheet);
  const ph = getHeaderMap(poSheet);

  const poRow = poData.find(row => row[ph['PO_ID']] === poId);
  if (!poRow) {
    throw new Error(`PO with ID ${poId} not found.`);
  }

  // Update PO status and remarks (implementation depends on your requirements)
  poRow[ph['Status_Code']] = action === 'Approved' ? 'PO_APPROVED' : 'PO_REJECTED';
  poRow[ph['Remarks']] = remarks || '';

  poSheet.getRange(poRow.rowIndex + 1, 1, 1, poRow.length).setValues([poRow]);
  console.log(`PO updated successfully: ${JSON.stringify(poRow)}`);
}
