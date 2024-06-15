/**
 * Class object for Unlock Smart Invoice Management: Gemini, Gmail, and Google Apps Script Integration
 */
class InvoiceManager {

  /**
   *
   * @param {Object} object API key or access token for using Gemini API.
   * @param {String} object.apiKey API key.
   * @param {String} object.accessToken Access token.
   * @param {String} object.model Model. Default is "models/gemini-1.5-pro-latest".
   * @param {String} object.version Version of API. Default is "v1beta".
   */
  constructor(object = null) {
    /** @private */
    this.object = object;

    /** @private */
    this.apiKey = null;

    /** @private */
    this.useAccessToken = false;

    /** @private */
    this.model = "models/gemini-1.5-flash-latest";

    /** @private */
    this.version = "v1beta";

    /** @private */
    this.labelName = null;

    /** @private */
    this.cycleMinTimeDrivenTrigger = 10;

    /** @private */
    this.extraTime = this.cycleMinTimeDrivenTrigger * 2; // In the current stage, this value is set as 2 times this.cycleMinTimeDrivenTrigger. For example, when the script is run by the time-driven trigger with a cycle of 10 minutes, this script retrieved the emails from 30 minutes before. By this, even when the script was finished by an error, you have 2 chances for retrying.

    /** @private */
    this.mainFunctionName = "main";

    /** @private */
    this.notifyModificationpointsToSender = false;

    /** @private */
    this.configurationSheetName = "configuration";

    /** @private */
    this.logSheetName = "log";

    /** @private */
    this.accessToken = null;

    /** @private */
    this.dashboardSheet = null;

    /** @private */
    this.logSheet = null;

    /** @private */
    this.keys = ["apiKey", "useAccessToken", "model", "version", "labelName", "cycleMinTimeDrivenTrigger", "extraTime", "mainFunctionName", "notifyModificationpointsToSender"];

    /** @private */
    this.now = new Date();

    /** @private */
    this.waitTime = 5; // seconds

    /** @private */
    this.rowColors = { doneRows: "#d9ead3", invalidRows: "#f4cccc", unrelatedRows: "#d9d9d9" };

  }

  /**
   * ### Description
   * Main method.
   *
   * @return {void}
   */
  async run() {
    this.setTimeDrivenTriggers_();
    this.getSheets_();
    if (this.object && Object.keys(this.object).length > 0) {
      this.keys.forEach(k => {
        if (this.object[k]) {
          this[k] = this.object[k];
        }
      });
    } else {
      this.getInitParams_();
    }
    let messages = this.getEmailsWithInvoices_();
    const coloredRows = {
      doneRows: [],
      invalidRows: [],
      unrelatedRows: [],
    };
    const values = [];
    for (let i = 0; i < messages.length; i++) {
      const { threadId, messageId, pdfFiles, searchUrl, sender, subject, messageObj } = messages[i];
      const pdfFilesLen = pdfFiles.length;
      for (let j = 0; j < pdfFilesLen; j++) {
        const blob = pdfFiles[j];
        const o = await this.parseInvoiceByGemini_(blob);
        const prefix = [this.now, threadId, messageId, searchUrl, sender, subject];
        if (o.check.invoice == true && o.check.invalidCheck == false) {
          values.push([...prefix, true, true, null, JSON.stringify(o), null]);
          coloredRows.doneRows.push(values.length - 1);
        } else if (o.check.invoice == true && o.check.invalidCheck == true) {
          if (this.notifyModificationpointsToSender == true) {
            const msg = `This message is automatically sent from a script for checking your invoice by Gemini.\nNow, Gemini suggested the modification points in your invoice. Please confirm the following modification points and send the modified invoice again.\n\nModification points:\n${o.check.invalidPoints}`;
            messageObj.reply(msg);
          }
          values.push([...prefix, true, false, o.check.invalidPoints, JSON.stringify(o), null]);
          coloredRows.invalidRows.push(values.length - 1);
        } else if (o.check.invoice == false) {
          values.push([...prefix, false, null, null, null, null]);
          coloredRows.unrelatedRows.push(values.length - 1);
        } else {
          values.push([...prefix, null, null, null, null, JSON.stringify(o)]);
        }
        if (pdfFilesLen >= 2) {
          Utilities.sleep(this.waitTime * 1000);
        }
      }
    }
    const valuesLen = values.length;
    let msg = "";
    if (valuesLen > 0) {
      const offset = this.logSheet.getLastRow() + 1;
      this.logSheet.getRange(offset, 1, values.length, values[0].length).setValues(values);
      ["doneRows", "invalidRows", "unrelatedRows"].forEach(k => {
        if (coloredRows[k].length > 0) {
          this.logSheet.getRangeList(coloredRows[k].map(e => `${e + offset}:${e + offset}`)).setBackground(this.rowColors[k]);
        }
      });
      msg = `${valuesLen} emails were processed.`;
      this.logSheet.activate();
    } else {
      msg = "No emails were processed.";
    }
    this.showLog_(msg);
  }

  /**
   * ### Description
   * Get work sheets.
   *
   * @private
   */
  getSheets_() {
    // In the current stage, the sheet names of "configuration" and "log" are fixed.
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    this.dashboardSheet = ss.getSheetByName(this.configurationSheetName) || ss.insertSheet(this.configurationSheetName);
    this.logSheet = ss.getSheetByName(this.logSheetName) || ss.insertSheet(this.logSheetName);
    if (!this.dashboardSheet || !this.logSheet) {
      this.showError_("Sheet names are changed from the default names of 'dashboard' and 'log'. Please confirm them.");
    }
  }

  /**
   * ### Description
   * Delete time-driven triggers.
   */
  deleteTimeDrivenTriggers() {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() == this.mainFunctionName) {
        ScriptApp.deleteTrigger(t);
      }
    });
  }

  /**
   * ### Description
   * Set time-driven triggers.
   *
   * @private
   */
  setTimeDrivenTriggers_() {
    this.deleteTimeDrivenTriggers();
    ScriptApp.newTrigger(this.mainFunctionName).timeBased().everyMinutes(this.cycleMinTimeDrivenTrigger).create();
  }

  /**
   * ### Description
   * Get user's values.
   *
   * @private
   */
  getInitParams_() {
    // Dashboard sheet has 2 header rows.
    const v = this.dashboardSheet.getDataRange().getValues();
    if (v.join("") == "") {
      const defaultValues = [
        ["Configuration", "", ""],
        ["Names of values", "Your values", "Descriptions"],
        ["apiKey", "", "This API key is used for requesting Gemini API."],
        ["useAccessToken", false, "Default is FALSE. If you want to use your access token, please set this as TRUE. At that time, The access token is retrieved by ScriptApp.getOAuthToken(). When you use this value as TRUE, the API key is not used."],
        ["model", "models/gemini-1.5-flash-latest", "This value is the model name to use for generating content. Default is \"models/gemini-1.5-flash-latest\"."],
        ["version", "v1beta", "This value is the version of Gemini API. Default is v1beta."],
        ["labelName", "", "This label name on Gmail is used for searching the emails of invoices. If you have no label, please set empty. By this, this application retrieves the emails by searching a word \"invoice\" in the email."],
        ["cycleMinTimeDrivenTrigger", 10, "Unit is minutes. The default is 10 minutes. This value is used for executing the script for managing the invoices of emails by the time-driven trigger. Please select one of 5, 10, 15, or 30 from the dropdown list."],
        ["mainFunctionName", "main", "This value is the name of main function. Default is \"main\"."],
        ["notifyModificationpointsToSender", false, "Default is false. When this value is true, when the invoice has modification points, an email including them is automatically sent as a reply mail."]
      ];
      this.dashboardSheet.getRange(1, 1, defaultValues.length, defaultValues[0].length).setValues(defaultValues);
    }
    const [, , ...values] = v;
    values.forEach(([a, b]) => {
      const ta = a.trim();
      const tb = typeof b == "string" ? b.trim() : b;
      if (this.keys.includes(ta)) {
        this[ta] = tb ?? this[ta];
      }
    });
    if (this.useAccessToken === true) {
      this.accessToken = ScriptApp.getOAuthToken();
    }
  }

  /**
   * ### Description
   * Get log from the log sheet.
   *
   * @returns {Array} Log.
   * @private
   */
  getLog_() {
    let [head, ...values] = this.logSheet.getDataRange().getValues();
    if (head.join("") == "") {
      head = ["date", "threadId", "messageId", "searchUrl", "sender", "subject", "hasInvoice", "isValidInvoice", "modificationPoints", "parsedInvoice", "notes"];
      this.logSheet.getRange(1, 1, 1, head.length).setValues([head]);
    }
    return values.map(r => head.reduce((o, h, j) => (o[h] = r[j], o), {}));
  }

  /**
   * ### Description
   * Get emails including the invoices as PDF files from Gmail.
   *
   * @returns {Array} Retrieved messages including PDF files.
   * @private
   */
  getEmailsWithInvoices_() {
    const processedMessageIds = this.getLog_().map(({ messageId }) => messageId);
    const now = this.now.getTime();
    const after = (now - ((this.cycleMinTimeDrivenTrigger + this.extraTime) * 60 * 1000)).toString();
    let searchQuery = `after:${after.slice(0, after.length - 3)} has:attachment`;
    if (this.labelName != "") {
      searchQuery += ` label:invoices`;
    } else {
      searchQuery += ` label:INBOX`;
    }
    const threads = GmailApp.search(searchQuery);
    const messages = threads.reduce((ar, t) => {
      const threadId = t.getId();
      t.getMessages().forEach(m => {
        const files = m.getAttachments();
        if (files.length > 0) {
          const pdfFiles = files.filter(f => f.getContentType() == MimeType.PDF).map(a => Utilities.newBlob(a.getBytes(), a.getContentType(), a.getName));
          if (pdfFiles.length > 0) {
            const messageId = m.getId();
            if (!processedMessageIds.includes(messageId)) {
              const sender = m.getFrom();
              const subject = m.getSubject();
              const searchUrl = `https://mail.google.com/mail/#search/rfc822msgid:${encodeURIComponent(m.getHeader("Message-ID"))}`;
              ar.push({ threadId, messageId, pdfFiles, searchUrl, sender, subject, messageObj: m });
            }
          }
        }
      });
      return ar;
    }, []);
    return messages;
  }

  /**
   * ### Description
   * Generate content from PDF blob of invoice.
   * This method generates content by Gemini API with my Google Apps Script library [GeminiWithFiles](https://github.com/tanaikech/GeminiWithFiles).
   *
   * @param {Blob} blob PDF blob of invoice.
   * @returns {object} Generated content as a JSON object.
   * @private
   */
  async parseInvoiceByGemini_(blob) {
    try {
      const jsonSchema = {
        description: "About the invoices of the following files, check carefully, and create an array including an object that parses the following images of the invoices by pointing out the detailed improvement points in the invoice. Confirm by calculating 3 times whether the total amount of the invoice is correct. Furthermore, confirm whether the name, address, phone number, and the required fields are written in the invoice.",
        type: "object",
        properties: {
          check: {
            description: "Point out the improvement points in the invoice. Return the detailed imprivement points like details of invalid, insufficient, wrong, and miscalculated parts. Here, ignore the calculation of tax.",
            type: "object",
            properties: {
              invoice: {
                description: "If the file is an invoice, it's true. If the file is not an invoice, it's false.",
                type: "boolean",
              },
              invalidCheck: {
                description: "Details of invalid, insufficient, wrong, and miscalculated points of the invoice. When no issue was found, this should be false. When issues were found, this should be true.",
                type: "boolean"
              },
              invalidPoints: {
                description: "Details of invalid, insufficient, wrong, and miscalculated points of the invoice. When no issue was found, this should be no value.",
                type: "string"
              }
            },
            required: ["invoice", "invalidCheck"],
            additionalProperties: false,
          },
          parse: {
            description: "Create an object parsed the invoice.",
            type: "object",
            properties: {
              name: { description: "Name given as 'Filename'", type: "string" },
              invoiceTitle: { description: "Title of invoice", type: "string" },
              invoiceDate: { description: "Date of invoice", type: "string" },
              invoiceNumber: { description: "Number of the invoice", type: "string" },
              invoiceDestinationName: { description: "Name of destination of invoice", type: "string" },
              invoiceDestinationAddress: { description: "Address of the destination of invoice", type: "string" },
              totalCost: { description: "Total cost of all costs", type: "string" },
              table: {
                description: "Table of the invoice. This is a 2-dimensional array. Add the first header row to the table in the 2-dimensional array. The column should be 'title or description of item', 'number of items', 'unit cost', 'total cost'",
                type: "array",
              },
            },
            required: [
              "name",
              "invoiceTitle",
              "invoiceDate",
              "invoiceNumber",
              "invoiceDestinationName",
              "invoiceDestinationAddress",
              "totalCost",
              "table",
            ],
            additionalProperties: false,
          }
        },
        required: [
          "check",
          "parse",
        ],
        additionalProperties: false,
      };
      const tempObj = { model: this.model, version: this.version, response_mime_type: "application/json" };
      if (this.accessToken) {
        tempObj.accessToken = this.accessToken;
      } else if (this.apiKey) {
        tempObj.apiKey = this.apiKey;
      } else {
        showError_("Please set your API key for using Gemini API.");
      }
      const g = new GeminiWithFiles.geminiWithFiles(tempObj);
      const fileList = await g.setBlobs([blob], true).uploadFiles();
      const res = g.withUploadedFilesByGenerateContent(fileList).generateContent({ jsonSchema });
      g.deleteFiles(fileList.map(({ name }) => name));
      return res;
    } catch ({ stack }) {
      this.showError_(stack);
    }
  }

  /**
   * ### Description
   * Show message as a log.
   *
   * @param {string} msg Message.
   * 
   * @private
   */
  showLog_(msg) {
    console.log(msg);
    Browser.msgBox(msg);
  }

  /**
   * ### Description
   * Show error message.
   *
   * @param {string} msg Error message.
   * 
   * @private
   */
  showError_(msg) {
    console.log(msg);
    Browser.msgBox(msg);
    throw new Error(msg);
  }
}
