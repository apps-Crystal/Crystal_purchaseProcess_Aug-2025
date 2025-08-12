// poForm.gs
function poFormPage(user) {
  const t = HtmlService.createTemplateFromFile('poFrm');
  t.user = user;
  return t.evaluate().getContent();
}
