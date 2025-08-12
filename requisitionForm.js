// requisitionForm.gs
// function requisitionFormPage(user) {
//   // We'll serve an HTML template file called 'requisitionForm'
//   const t = HtmlService.createTemplateFromFile('requisitionFrm');
//   t.user = user;
//   return t.evaluate().getContent().setMimeType(ContentService.MimeType.HTML).getContent();
// }

function requisitionFormPage(user) {
  const t = HtmlService.createTemplateFromFile('requisitionFrm');
  t.user = user;
  return t.evaluate().getContent();
}