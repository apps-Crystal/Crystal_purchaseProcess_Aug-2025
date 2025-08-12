/**
 * @OnlyCurrentDoc
 */

/**
 * Fetches dashboard data from the PR_Master sheet and returns it as a JSON object.
 * This function is intended to be called from client-side JavaScript.
 *
 * @return {object} An object containing dashboard KPIs and recent requisitions data.
 */
function getDashboardData() {
  const sheetName = "PR_Master";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    // Handle the case where the sheet is not found
    return {
      totalRequisitions: 0,
      pendingApproval: 0,
      approved: 0,
      totalValue: 0,
      pendingGRN: 0,
      recentRequisitions: []
    };
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Assuming the first row contains headers
  const headers = values.shift();

  // Find the index of key columns
  const idIndex = headers.indexOf("PR_ID");
  const statusIndex = headers.indexOf("Status_Label");
  const totalIndex = headers.indexOf("Total_Incl_GST");
  const siteIndex = headers.indexOf("Site");
  const dateIndex = headers.indexOf("Date_of_Requisition");

  // Calculate KPIs
  let totalRequisitions = 0;
  let pendingApproval = 0;
  let approved = 0;
  let totalValue = 0;
  let pendingGRN = 0;

  // Recent requisitions (limit to the last 5 for this example)
  const recentRequisitions = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    totalRequisitions++;

    const status = row[statusIndex];
    const total = parseFloat(row[totalIndex]);

    if (status === 'Pending Approval') {
      pendingApproval++;
    } else if (status === 'Approved') {
      approved++;
    } else if (status === 'PO Created') {
      // Assuming 'PO Created' from the screenshot implies a pending GRN
      pendingGRN++;
    }

    if (!isNaN(total)) {
      totalValue += total;
    }
  }

  // Get recent requisitions from the bottom of the sheet
  const recentData = values.slice(-5);
  recentData.forEach(row => {
    recentRequisitions.push({
      id: row[idIndex],
      site: row[siteIndex],
      amount: parseFloat(row[totalIndex]),
      status: row[statusIndex],
      date: row[dateIndex]
    });
  });

  return {
    totalRequisitions: totalRequisitions,
    pendingApproval: pendingApproval,
    approved: approved,
    totalValue: totalValue,
    pendingGRN: pendingGRN,
    recentRequisitions: recentRequisitions.reverse() // Reverse to show most recent first
  };
}
