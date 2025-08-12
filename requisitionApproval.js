// requisitionApproval.gs
function requisitionApprovalPage() {
  // get
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

function getPRDetails(prId) {
  console.log(`Fetching details for PR ID: ${prId}`);
  const ss = SpreadsheetApp.getActive();
  const prSheet = ss.getSheetByName("PR_Master");
  if (!prSheet) throw new Error("PR_Master sheet not found.");

  const prData = safeGetData(prSheet);
  console.log(`PR Data: ${JSON.stringify(prData)}`);
  const h = getHeaderMap(prSheet);

  const prRow = prData.find(r => r[h['PR_ID']] === prId);
  console.log(`PR Row: ${JSON.stringify(prRow)}`);
  if (!prRow) {
    throw new Error(`PR with ID ${prId} not found.`);
  }

  // Extract PR details
  const prDetails = {
    prId: prRow[h['PR_ID']],
    site: prRow[h['Site']],
    requestedBy: prRow[h['Requested_By']],
    date: prRow[h['Date_of_Requisition']],
    expectedDelivery: prRow[h['Expected_Delivery']],
    status: prRow[h['Status_Label']],
    vendorCompany: prRow[h['Vendor_Company']],
    vendorContact: prRow[h['Vendor_Contact']],
    vendorEmail: prRow[h['Vendor_Email']],
    vendorPhone: prRow[h['Vendor_Phone']],
    remarks: prRow[h['Approver_Remarks']],
    items: [], // Placeholder for items
    total: prRow[h['Total_Incl_GST']]
  };

  // Fetch items (assuming items are stored in another sheet or format)
  const itemSheet = ss.getSheetByName("PR_Items");
  if (itemSheet) {
    const itemData = safeGetData(itemSheet);
    const itemHeader = getHeaderMap(itemSheet);
    prDetails.items = itemData
      .filter(item => item[itemHeader['PR_ID']] === prId)
      .map(item => ({
        name: item[itemHeader['Item_Name']],
        qty: item[itemHeader['Quantity']],
        unit: item[itemHeader['Unit']],
        price: item[itemHeader['Unit_Price']],
        total: item[itemHeader['Total']]
      }));
  }

  return prDetails;
}
