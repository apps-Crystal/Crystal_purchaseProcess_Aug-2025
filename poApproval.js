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
  const prSheet = ss.getSheetByName("PR_Master");
  if (!poSheet) return [];
  const poData = safeGetData(poSheet);
  const ph = getHeaderMap(poSheet);
  const list = [];
  for (let i=0;i<poData.length;i++){
    const row = poData[i];
    if (row[ph['Status_Code']] === 'PO_POSTED') {
      list.push({
        poId: row[ph['PO_ID']],
        prId: row[ph['PR_ID']],
        site: row[ph['Site']],
        totalValue: row[ph['Total_Incl_GST']],
        date: row[ph['PO_Date']],
        attach: row[ph['PO_File_URL']]
      });
    }
  }
  return list;
}
