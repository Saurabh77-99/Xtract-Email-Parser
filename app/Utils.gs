// ============================================================
// FETCH
// ============================================================

function fetchFromBackend(path) {
  try {
    var res = UrlFetchApp.fetch(BACKEND_URL + path, {
      muteHttpExceptions: true,
      headers: { Accept: "application/json" },
    });
    if (res.getResponseCode() === 200) return JSON.parse(res.getContentText());
    console.error("Backend " + res.getResponseCode() + " for " + path);
  } catch (e) {
    console.error("Fetch error " + path + ": " + e);
  }
  return null;
}

// ============================================================
// TEXT ANCHOR PARSER (GAS side — for Direct mode)
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
    var lines = before.split("\n");
    return extractByBoundary(lines[lines.length - 1].trim(), boundary);
  }
  return "";
}

function extractByBoundary(text, boundary) {
  if (boundary === "word") {
    var m = text.match(/^\S+/);
    return m ? m[0] : "";
  }
  if (boundary === "line") return text.split("\n")[0].trim();
  if (boundary === "paragraph") return text.split(/\n\n/)[0].trim();
  return text.split("\n")[0].trim();
}

// ============================================================
// LABELS
// ============================================================

function hasLabel(message, labelName) {
  return message
    .getThread()
    .getLabels()
    .some(function (l) {
      return l.getName() === labelName;
    });
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
      label.getThreads().forEach(function (t) {
        t.removeLabel(label);
      });
      console.log("Reset: " + label.getName());
    }
  });
}

// ============================================================
// INGEST — simple body only
// ============================================================

function ingestMessage(message, ruleId) {
  try {
    var res = UrlFetchApp.fetch(BACKEND_URL + "/ingest", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        ruleId: ruleId,
        messageId: message.getId(),
        subject: message.getSubject(),
        sender: message.getFrom(),
        rawBody: message.getPlainBody().substring(0, 5000),
      }),
      muteHttpExceptions: true,
    });
    return JSON.parse(res.getContentText());
  } catch (e) {
    console.error("ingestMessage: " + e);
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
        mime ===
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
        mime === "application/vnd.ms-excel"
      ) {
        var obj = {
          name: att.getName(),
          mimeType: mime,
          data: Utilities.base64Encode(att.getBytes()),
          isExcel: mime !== "application/pdf",
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
        ruleId: ruleId,
        messageId: message.getId(),
        subject: message.getSubject(),
        sender: message.getFrom(),
        rawBody: message.getPlainBody().substring(0, 5000),
        attachments: attachments,
      }),
      muteHttpExceptions: true,
    });
    return JSON.parse(res.getContentText());
  } catch (e) {
    console.error("ingestMessageRich: " + e);
    return null;
  }
}

// ============================================================
// EXCEL + PDF EXTRACTION
// ============================================================

function extractExcelContent(base64Data, fileName) {
  try {
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      fileName,
    );
    var tempFile = DriveApp.createFile(blob);
    Utilities.sleep(2000);
    var sheetFile = Drive.Files.copy(
      { mimeType: "application/vnd.google-apps.spreadsheet" },
      tempFile.getId(),
    );
    var sheet = SpreadsheetApp.openById(sheetFile.id).getSheets()[0];
    var textContent = "";
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow > 0 && lastCol > 0) {
      sheet
        .getRange(1, 1, lastRow, lastCol)
        .getValues()
        .forEach(function (row) {
          var r = row
            .filter(function (c) {
              return c !== "";
            })
            .join(" | ");
          if (r) textContent += r + "\n";
        });
    }
    DriveApp.getFileById(sheetFile.id).setTrashed(true);
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    console.log(
      "Excel extracted: " + fileName + " (" + textContent.length + " chars)",
    );
    return textContent.trim();
  } catch (e) {
    console.error("Excel extraction failed for " + fileName + ": " + e);
    return "";
  }
}

function extractPdfContent(base64Data, fileName) {
  try {
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      "application/pdf",
      fileName,
    );
    var tempFile = DriveApp.createFile(blob);
    Utilities.sleep(2000);
    var docFile = Drive.Files.copy(
      {
        mimeType: "application/vnd.google-apps.document",
        title: fileName + "_ocr",
      },
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
// SYNC
// ============================================================

function processEmails() {
  var rules = fetchFromBackend("/rules");
  if (!rules) return;
  rules.forEach(function (rule) {
    if (!rule.isActive) return;
    var threads = GmailApp.search(rule.criteriaQuery);
    threads.forEach(function (thread) {
      thread.getMessages().forEach(function (message) {
        if (hasLabel(message, "Parsed/Rule_" + rule.id)) return;
        var result = ingestMessage(message, rule.id);
        if (result && result.status === "success")
          markAsProcessed(message, rule.id);
      });
    });
  });
}

// ============================================================
// BACKGROUND JOB
// ============================================================

function runBackgroundSync() {
  var props = PropertiesService.getUserProperties();
  var jobId = props.getProperty("latest_job_id");
  if (!jobId) return;

  var ruleId = parseInt(props.getProperty(jobId + "_ruleId"));
  var query = props.getProperty(jobId + "_query");
  if (!ruleId || !query) return;

  console.log("Background sync: " + jobId + " | query: " + query);
  var threads = GmailApp.search(query);
  var count = 0;

  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (message) {
      if (hasLabel(message, "Parsed/Rule_" + ruleId)) return;
      var result = ingestMessage(message, ruleId);
      if (result && result.status === "success") {
        markAsProcessed(message, ruleId);
        count++;
      }
    });
  });

  props.setProperty(jobId + "_count", count.toString());
  props.setProperty(jobId + "_status", "done");
  console.log("Background sync done: " + count);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "runBackgroundSync")
      ScriptApp.deleteTrigger(t);
  });
}

function getJobStatus(jobId) {
  var props = PropertiesService.getUserProperties();
  return {
    status: props.getProperty(jobId + "_status") || "running",
    count: props.getProperty(jobId + "_count") || "0",
    ruleId: props.getProperty(jobId + "_ruleId") || "",
  };
}

function clearJobProperties(jobId) {
  var props = PropertiesService.getUserProperties();
  ["_ruleId", "_query", "_status", "_count"].forEach(function (k) {
    props.deleteProperty(jobId + k);
  });
}

// ============================================================
// AUTH
// ============================================================

function triggerAuthorization() {
  try {
    var temp = SpreadsheetApp.create("Auth Test");
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
          "⚠️ Check Gmail for auth banner.",
        ),
      )
      .build();
  }
}

// ============================================================
// DEBUG
// ============================================================

function debugSync() {
  var rules = fetchFromBackend("/rules");
  if (!rules) {
    console.log("No rules.");
    return;
  }
  rules.forEach(function (rule) {
    console.log(
      "Rule [" + rule.id + "]: " + rule.name + " | " + rule.criteriaQuery,
    );
    var threads = GmailApp.search(rule.criteriaQuery);
    console.log("Threads: " + threads.length);
    threads.forEach(function (thread) {
      thread.getMessages().forEach(function (message) {
        var done = hasLabel(message, "Parsed/Rule_" + rule.id);
        console.log("  " + (done ? "✅" : "🔲") + " " + message.getSubject());
      });
    });
  });
}
