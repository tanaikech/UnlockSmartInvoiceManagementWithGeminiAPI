/*
This report describes an invoice processing application built with Google Apps Script. It leverages Gemini, a large language model, to automatically parse invoices received as email attachments and automates the process using time-driven triggers.

Repository: https://github.com/tanaikech/PUnlockSmartInvoiceManagementWithGeminiAPI
*/

// When this function is run, the installed time-driven trigger is deleted.
function stopTrigger() {
  new InvoiceManager().deleteTimeDrivenTriggers();
  Browser.msgBox("Trigger was removed.");
}

// This is a main function.
// When this function is run, this application is launched.
function main() {
  new InvoiceManager().run();
}
