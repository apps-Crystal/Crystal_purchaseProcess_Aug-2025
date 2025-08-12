/**
 * Code.gs
 * - doGet routing for pages
 * - include(filename) helper for HTML templates
 */

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'home';
  const user = {
    email: Session.getActiveUser().getEmail(),
    name: Session.getActiveUser().getEmail().split('@')[0]
  };

  let pageContent = '';
  let pageTitle = 'Purchase Requisition System';

  switch (page) {
    case 'home':
      pageTitle = 'Dashboard';
      pageContent = HtmlService.createTemplateFromFile('home').evaluate().getContent();
      break;

    case 'requisitionForm':
      pageTitle = 'New Requisition';
      // Call the function from requisitionForm.js which returns the page's HTML content.
      // The function itself builds the HTML string.
      pageContent = requisitionFormPage(user).getContent();
      break;

    case 'requisitionApproval':
      pageTitle = 'PR Approvals';
      pageContent = requisitionApprovalPage();
      break;

    case 'poForm':
      pageTitle = 'Create PO';
      pageContent = poFormPage(user);
      break;

    case 'poApproval':
      pageTitle = 'PO Approvals';
      pageContent = poApprovalPage();
      break;

    case 'paymentSubmission':
      pageTitle = 'Payment Submission';
      pageContent = paymentSubmissionPage();
      break;

    case 'addVendor':
      pageTitle = 'Register New Vendor';
      pageContent = addVendorPage(user);
      break;

    default:
      pageTitle = 'Dashboard';
      pageContent = HtmlService.createTemplateFromFile('home').evaluate().getContent();
  }

  const layout = HtmlService.createTemplateFromFile('layout');
  layout.pageContent = pageContent;
  layout.activePage = page;
  layout.pageTitle = pageTitle;
  layout.user = user;

  return layout.evaluate().setTitle(pageTitle).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
