// ============================================================
// BACKEND FETCH HELPER
// ============================================================

function fetchFromBackend(path) {
  try {
    var res = UrlFetchApp.fetch(BACKEND_URL + path, {
      muteHttpExceptions: true,
      headers: { Accept: "application/json" },
    });
    if (res.getResponseCode() === 200) {
      return JSON.parse(res.getContentText());
    }
    console.error("Backend " + res.getResponseCode() + " for " + path);
  } catch (e) {
    console.error("fetchFromBackend error " + path + ": " + e);
  }
  return null;
}

// ============================================================
// INGEST — plain body (used by scheduled processEmails)
// ============================================================

function ingestMessage(message, ruleId) {
  return ingestMessageWithBody(message, ruleId, message.getPlainBody());
}

// ============================================================
// INGEST — explicit body (used by handleSearchSubmit)
// Allows passing HTML or plain body depending on user preference.
// ============================================================

function ingestMessageWithBody(message, ruleId, body) {
  try {
    var res = UrlFetchApp.fetch(BACKEND_URL + "/ingest", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        ruleId:    ruleId,
        messageId: message.getId(),
        subject:   message.getSubject(),
        sender:    message.getFrom(),
        rawBody:   (body || "").substring(0, 5000),
      }),
      muteHttpExceptions: true,
    });
    return JSON.parse(res.getContentText());
  } catch (e) {
    console.error("ingestMessageWithBody error: " + e);
    return null;
  }
}

// ============================================================
// INGEST RICH — body + PDF/Excel attachments
// ============================================================

function ingestMessageRich(message, ruleId) {
  try {
    var attachments = [];
    message.getAttachments().forEach(function (att) {
      var mime = att.getContentType();
      if (
        mime === "application/pdf" ||
        mime === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
        mime === "application/vnd.ms-excel"
      ) {
        var obj = {
          name:     att.getName(),
          mimeType: mime,
          data:     Utilities.base64Encode(att.getBytes()),
          isExcel:  mime !== "application/pdf",
        };
        if (obj.isExcel) {
          var txt = extractExcelContent(obj.data, obj.name);
          if (txt) obj.extractedText = txt;
        } else {
          Utilities.sleep(3000);
          var pdf = extractPdfContent(obj.data, obj.name);
          if (pdf) obj.extractedText = pdf;
        }
        attachments.push(obj);
      }
    });

    var res = UrlFetchApp.fetch(BACKEND_URL + "/ingest-rich", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        ruleId:      ruleId,
        messageId:   message.getId(),
        subject:     message.getSubject(),
        sender:      message.getFrom(),
        rawBody:     message.getPlainBody().substring(0, 5000),
        attachments: attachments,
      }),
      muteHttpExceptions: true,
    });
    return JSON.parse(res.getContentText());
  } catch (e) {
    console.error("ingestMessageRich error: " + e);
    return null;
  }
}

// ============================================================
// TEXT ANCHOR PARSER (Apps Script side — used in Direct mode)
// ============================================================

function parseTextAnchor(body, anchor, type, boundary) {
  if (!anchor || !body) return "";
  var idx = body.indexOf(anchor);
  if (idx === -1) return "";

  if (type === "after") {
    var rest = body.substring(idx + anchor.length).replace(/^\s+/, "");
    return extractByBoundary(rest, boundary);
  } else if (type === "before") {
    var before = body.substring(0, idx).replace(/\s+$/, "");
    var lines  = before.split("\n");
    return extractByBoundary(lines[lines.length - 1].trim(), boundary);
  }
  return "";
}

function extractByBoundary(text, boundary) {
  if (boundary === "word")      return (text.match(/^\S+/) || [""])[0];
  if (boundary === "line")      return text.split("\n")[0].trim();
  if (boundary === "paragraph") return text.split(/\n\n/)[0].trim();
  return text.split("\n")[0].trim();
}

// ============================================================
// GMAIL LABEL HELPERS
// ============================================================

function hasLabel(message, labelName) {
  return message
    .getThread()
    .getLabels()
    .some(function (l) { return l.getName() === labelName; });
}

function markAsProcessed(message, ruleId) {
  var labelName = "Parsed/Rule_" + ruleId;
  var label =
    GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
  message.getThread().addLabel(label);
}

function resetProcessedLabels() {
  GmailApp.getUserLabels().forEach(function (label) {
    if (label.getName().indexOf("Parsed/Rule_") === 0) {
      label.getThreads().forEach(function (t) { t.removeLabel(label); });
      console.log("Reset label: " + label.getName());
    }
  });
}

// ============================================================
// EXCEL EXTRACTION
// ============================================================

function extractExcelContent(base64Data, fileName) {
  try {
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      fileName,
    );
    var tempFile  = DriveApp.createFile(blob);
    Utilities.sleep(2000);
    var sheetFile = Drive.Files.copy(
      { mimeType: "application/vnd.google-apps.spreadsheet" },
      tempFile.getId(),
    );
    var sheet       = SpreadsheetApp.openById(sheetFile.id).getSheets()[0];
    var textContent = "";
    var lastRow     = sheet.getLastRow();
    var lastCol     = sheet.getLastColumn();
    if (lastRow > 0 && lastCol > 0) {
      sheet.getRange(1, 1, lastRow, lastCol).getValues().forEach(function (row) {
        var r = row.filter(function (c) { return c !== ""; }).join(" | ");
        if (r) textContent += r + "\n";
      });
    }
    DriveApp.getFileById(sheetFile.id).setTrashed(true);
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    console.log("Excel extracted: " + fileName + " (" + textContent.length + " chars)");
    return textContent.trim();
  } catch (e) {
    console.error("Excel extraction failed for " + fileName + ": " + e);
    return "";
  }
}

// ============================================================
// PDF EXTRACTION (OCR via Drive)
// ============================================================

function extractPdfContent(base64Data, fileName) {
  try {
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      "application/pdf",
      fileName,
    );
    var tempFile = DriveApp.createFile(blob);
    Utilities.sleep(2000);
    var docFile  = Drive.Files.copy(
      { mimeType: "application/vnd.google-apps.document", title: fileName + "_ocr" },
      tempFile.getId(),
      { ocr: true, ocrLanguage: "en" },
    );
    var text = DocumentApp.openById(docFile.id).getBody().getText();
    DriveApp.getFileById(docFile.id).setTrashed(true);
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    console.log("PDF extracted: " + fileName + " (" + text.length + " chars)");
    return text.trim();
  } catch (e) {
    console.error("PDF extraction failed for " + fileName + ": " + e);
    return "";
  }
}

// ============================================================
// SCHEDULED SYNC — runs via time-based trigger (setSchedule)
// Processes all active rules, skips already-labelled messages.
// ============================================================

function processEmails() {
  var rules = fetchFromBackend("/rules");
  if (!rules || rules.length === 0) {
    console.log("processEmails: no rules found.");
    return;
  }

  rules.forEach(function (rule) {
    if (!rule.isActive) return;
    console.log("processEmails: rule [" + rule.id + "] " + rule.criteriaQuery);
    var threads = GmailApp.search(rule.criteriaQuery, 0, MAX_THREADS);
    console.log("  Found " + threads.length + " threads.");

    threads.forEach(function (thread) {
      thread.getMessages().forEach(function (message) {
        if (hasLabel(message, "Parsed/Rule_" + rule.id)) return;
        try {
          var result = ingestMessage(message, rule.id);
          if (result && result.status === "success") {
            markAsProcessed(message, rule.id);
          }
        } catch (err) {
          console.error("processEmails ingest error: " + err);
        }
      });
    });
  });
}

// ============================================================
// AUTHORIZATION
// ============================================================

function triggerAuthorization() {
  try {
    var temp = SpreadsheetApp.create("Auth Test — Delete Me");
    DriveApp.getFileById(temp.getId()).setTrashed(true);
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText("✅ Authorization active!"),
      )
      .build();
  } catch (e) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "⚠️ Authorization needed — check Gmail for a permission banner.",
        ),
      )
      .build();
  }
}

// ============================================================
// DEBUG (run from Apps Script editor, not from card)
// ============================================================

function debugSync() {
  var rules = fetchFromBackend("/rules");
  if (!rules) {
    console.log("No rules found.");
    return;
  }
  rules.forEach(function (rule) {
    console.log("Rule [" + rule.id + "]: " + rule.name + " | " + rule.criteriaQuery);
    var threads = GmailApp.search(rule.criteriaQuery, 0, 20);
    console.log("  Threads matched: " + threads.length);
    threads.forEach(function (thread) {
      thread.getMessages().forEach(function (message) {
        var done = hasLabel(message, "Parsed/Rule_" + rule.id);
        console.log("  " + (done ? "✅" : "🔲") + " " + message.getSubject());
      });
    });
  });
}