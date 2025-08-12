// paymentSubmission.gs
function paymentSubmissionPage() {
  const t = HtmlService.createTemplateFromFile('paymentSubmissin');
  return t.evaluate().getContent();
}
