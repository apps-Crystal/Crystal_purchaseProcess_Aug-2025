// requisitionApproval.gs
function requisitionApprovalPage() {
  // get pending PRs and render template
  const t = HtmlService.createTemplateFromFile('requisitionApprovl');
  t.pending = getPendingPRsForUI(); // returns array
  return t.evaluate().getContent();
}

function getPendingPRsForUI() {
  const ss = SpreadsheetApp.getActive();
  const prSheet = ss.getSheetByName("PR_Master");
  const prData = safeGetData(prSheet);
  const h = getHeaderMap(prSheet);
  return prData.filter(r => r[h['Status_Code']] === 'PR_SUBMITTED').map(r => ({
    prId: r[h['PR_ID']],
    site: r[h['Site']],
    requestedBy: r[h['Requested_By']],
    total: r[h['Total_Incl_GST']],
    date: r[h['Date_of_Requisition']],
    vendorId: r[h['Vendor_ID']]
  }));
}

// server side handler to approve/reject from UI
function handlePRApproval(prId, action, remarks) {
  if (action === 'Approved') return approvePR(prId, remarks);
  else return rejectPR(prId, remarks);
}
