function addVendorPage(user) {
  const t = HtmlService.createTemplateFromFile('vendorForm');
  t.user = user;
  return t.evaluate().getContent();
}