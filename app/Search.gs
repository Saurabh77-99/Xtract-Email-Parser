// ============================================================
// SEARCH VIEW
// ============================================================

function renderSearchView() {
  var builder = CardService.newCardBuilder();
  builder.setHeader(
    CardService.newCardHeader()
      .setTitle("Email Parser")
      .setImageUrl(
        "https://www.gstatic.com/images/branding/product/1x/gmail_512dp.png",
      ),
  );
  builder.addSection(createTopNav("search"));

  // Filters
  builder.addSection(
    CardService.newCardSection()
      .setHeader("Filter emails by subject keywords")
      .addWidget(
        CardService.newTextInput()
          .setFieldName("f_subject")
          .setTitle("Subject")
          .setHint("e.g., Google Workspace Invoice"),
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName("f_sender")
          .setTitle("Sender")
          .setHint("e.g., payments-noreply@google.com"),
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName("f_start")
          .setTitle("Start (MM/DD/YYYY)")
          .setHint("e.g., 01/01/2025"),
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName("f_end")
          .setTitle("End (MM/DD/YYYY)")
          .setHint("e.g., 12/31/2025"),
      )
      .addWidget(
        CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.CHECK_BOX)
          .setFieldName("f_thread")
          .addItem(
            "Search within each thread (all messages, not just latest)",
            "yes",
            false,
          ),
      )
      .addWidget(
        CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.CHECK_BOX)
          .setFieldName("f_html")
          .addItem(
            "Search email as HTML (only use when plain body not present)",
            "yes",
            false,
          ),
      ),
  );

  // Extract Data — Field 1
  builder.addSection(
    CardService.newCardSection()
      .setHeader("Extract Data — Field 1")
      .addWidget(
        CardService.newTextInput()
          .setFieldName("col_1")
          .setTitle("Column name")
          .setHint("e.g., Amount"),
      )
      .addWidget(
        CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.DROPDOWN)
          .setFieldName("type_1")
          .setTitle("Method")
          .addItem("Text After", "after", true)
          .addItem("Text Before", "before", false)
          .addItem("Regex", "regex", false)
          .addItem("Entire Email", "entire", false),
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName("anchor_1")
          .setTitle("Text / Pattern")
          .setHint("e.g., Total: or Invoice No:"),
      )
      .addWidget(
        CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.DROPDOWN)
          .setFieldName("end_1")
          .setTitle("Until")
          .addItem("End of Word", "word", true)
          .addItem("End of Line", "line", false)
          .addItem("End of Paragraph", "paragraph", false),
      ),
  );

  // Extract Data — Field 2
  builder.addSection(
    CardService.newCardSection()
      .setHeader("Extract Data — Field 2 (optional)")
      .addWidget(
        CardService.newTextInput()
          .setFieldName("col_2")
          .setTitle("Column name")
          .setHint("e.g., Date"),
      )
      .addWidget(
        CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.DROPDOWN)
          .setFieldName("type_2")
          .setTitle("Method")
          .addItem("Text After", "after", true)
          .addItem("Text Before", "before", false)
          .addItem("Regex", "regex", false)
          .addItem("Entire Email", "entire", false),
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName("anchor_2")
          .setTitle("Text / Pattern")
          .setHint("e.g., Date:"),
      )
      .addWidget(
        CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.DROPDOWN)
          .setFieldName("end_2")
          .setTitle("Until")
          .addItem("End of Word", "word", true)
          .addItem("End of Line", "line", false)
          .addItem("End of Paragraph", "paragraph", false),
      ),
  );

  // Extract Data — Field 3
  builder.addSection(
    CardService.newCardSection()
      .setHeader("Extract Data — Field 3 (optional)")
      .addWidget(
        CardService.newTextInput()
          .setFieldName("col_3")
          .setTitle("Column name")
          .setHint("e.g., Invoice No"),
      )
      .addWidget(
        CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.DROPDOWN)
          .setFieldName("type_3")
          .setTitle("Method")
          .addItem("Text After", "after", true)
          .addItem("Text Before", "before", false)
          .addItem("Regex", "regex", false)
          .addItem("Entire Email", "entire", false),
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName("anchor_3")
          .setTitle("Text / Pattern")
          .setHint("e.g., Invoice No:"),
      )
      .addWidget(
        CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.DROPDOWN)
          .setFieldName("end_3")
          .setTitle("Until")
          .addItem("End of Word", "word", true)
          .addItem("End of Line", "line", false)
          .addItem("End of Paragraph", "paragraph", false),
      ),
  );

  // Output mode + Submit
  builder.addSection(
    CardService.newCardSection()
      .addWidget(
        CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.DROPDOWN)
          .setFieldName("output_mode")
          .setTitle("Output mode")
          .addItem("Store & export (backend + Google Sheets)", "sheet", true)
          .addItem(
            "Write directly to Sheet (no backend storage)",
            "direct",
            false,
          ),
      )
      .addWidget(
        CardService.newTextParagraph().setText(
          "💡 Tip: Use date filters to avoid timeouts on large mailboxes. Max " +
            MAX_THREADS +
            " threads per run.",
        ),
      )
      .addWidget(
        CardService.newTextButton()
          .setText("Submit")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(
            CardService.newAction().setFunctionName("handleSearchSubmit"),
          ),
      ),
  );

  return builder.build();
}

// ============================================================
// SUBMIT — runs synchronously
// ============================================================

function handleSearchSubmit(e) {
  var fi = e.formInput;
  var subject = fi.f_subject || "";
  var sender = fi.f_sender || "";
  var startDate = fi.f_start || "";
  var endDate = fi.f_end || "";
  var outputMode = fi.output_mode || "sheet";
  var searchThread = fi.f_thread === "yes";
  var searchHtml = fi.f_html === "yes";

  var query = "";
  if (subject) query += 'subject:"' + subject + '" ';
  if (sender) query += "from:" + sender + " ";
  if (startDate) query += "after:" + toGmailDate(startDate) + " ";
  if (endDate) query += "before:" + toGmailDate(endDate) + " ";
  query = query.trim() || "in:inbox";

  var targetFields = {};
  for (var i = 1; i <= 3; i++) {
    var col = fi["col_" + i] || "";
    var type = fi["type_" + i] || "after";
    var anchor = fi["anchor_" + i] || "";
    var end = fi["end_" + i] || "word";

    if (col && anchor) {
      var key = col.toLowerCase().replace(/\s+/g, "_");
      if (type === "entire") {
        targetFields[key] = { type: "entire", anchor: "", end: "paragraph" };
      } else if (type === "regex") {
        targetFields[key] = anchor;
      } else {
        targetFields[key] = { type: type, anchor: anchor, end: end };
      }
    }
  }

  if (Object.keys(targetFields).length === 0) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "⚠️ Add at least one extract field with a column name and text/pattern.",
        ),
      )
      .build();
  }

  if (outputMode === "direct") {
    return handleDirectMode(query, targetFields, searchThread, searchHtml);
  }

  var ruleName =
    "Search: " + (subject || sender || new Date().toLocaleDateString());

  var saveRes = UrlFetchApp.fetch(BACKEND_URL + "/rules", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      name: ruleName,
      criteriaQuery: query,
      targetFields: JSON.stringify(targetFields),
      extractionMode: "text_anchor",
    }),
    muteHttpExceptions: true,
  });

  if (saveRes.getResponseCode() !== 200) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "❌ Backend error (" + saveRes.getResponseCode() + ").",
        ),
      )
      .build();
  }

  var allRules = fetchFromBackend("/rules");
  var newRule = null;
  if (allRules) {
    allRules.forEach(function (r) {
      if (r.name === ruleName && (!newRule || r.id > newRule.id)) newRule = r;
    });
  }
  if (!newRule) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "❌ Could not retrieve new rule from backend.",
        ),
      )
      .build();
  }

  var threads = GmailApp.search(query, 0, MAX_THREADS);
  var count = 0;
  var skipped = 0;

  for (var t = 0; t < threads.length; t++) {
    var msgs = searchThread
      ? threads[t].getMessages()
      : [threads[t].getMessages()[threads[t].getMessages().length - 1]];

    for (var m = 0; m < msgs.length; m++) {
      var message = msgs[m];
      if (hasLabel(message, "Parsed/Rule_" + newRule.id)) {
        skipped++;
        continue;
      }
      try {
        var body = searchHtml ? message.getBody() : message.getPlainBody();
        var result = ingestMessageWithBody(message, newRule.id, body);
        if (result && result.status === "success") {
          markAsProcessed(message, newRule.id);
          count++;
        }
      } catch (err) {
        console.error(
          "Ingest error on message " + message.getId() + ": " + err,
        );
      }
    }
  }

  var notice =
    count > 0
      ? "✅ " + count + " email" + (count === 1 ? "" : "s") + " imported!"
      : skipped > 0
        ? "ℹ️ All matching emails already imported (skipped " + skipped + ")."
        : "⚠️ No emails matched the filter.";

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(notice))
    .setNavigation(
      CardService.newNavigation().pushCard(
        renderResultCard(newRule.id, ruleName, count, skipped),
      ),
    )
    .build();
}

// ============================================================
// RESULT CARD
// ============================================================

function renderResultCard(ruleId, jobName, count, skipped) {
  var builder = CardService.newCardBuilder();
  builder.setHeader(
    CardService.newCardHeader()
      .setTitle("Email Parser")
      .setSubtitle("Import Complete"),
  );
  builder.addSection(createTopNav("search"));

  var section = CardService.newCardSection().setHeader(
    jobName || "Import Result",
  );

  if (count > 0) {
    section.addWidget(
      CardService.newTextParagraph().setText(
        "✅ " +
          count +
          " email" +
          (count === 1 ? "" : "s") +
          " successfully imported." +
          (skipped > 0 ? "\n(" + skipped + " already imported, skipped.)" : ""),
      ),
    );
    section.addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText("📊 EXPORT TO SHEETS")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName("handleExportToSheets")
                .setParameters({ ruleId: String(ruleId) }),
            ),
        )
        .addButton(
          CardService.newTextButton()
            .setText("📄 EXPORT TO PDF")
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName("handleExportToPdf")
                .setParameters({ ruleId: String(ruleId) }),
            ),
        ),
    );
  } else if (skipped > 0) {
    section.addWidget(
      CardService.newTextParagraph().setText(
        "ℹ️ " +
          skipped +
          " email" +
          (skipped === 1 ? "" : "s") +
          " already imported previously.\n\nYou can still export them below.",
      ),
    );
    section.addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText("📊 EXPORT TO SHEETS")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName("handleExportToSheets")
                .setParameters({ ruleId: String(ruleId) }),
            ),
        )
        .addButton(
          CardService.newTextButton()
            .setText("📄 EXPORT TO PDF")
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName("handleExportToPdf")
                .setParameters({ ruleId: String(ruleId) }),
            ),
        ),
    );
  } else {
    section.addWidget(
      CardService.newTextParagraph().setText(
        "⚠️ No emails matched the filter.\n\nTry adjusting subject, sender, or date range.",
      ),
    );
  }

  section.addWidget(CardService.newDivider());
  section.addWidget(
    CardService.newTextButton()
      .setText("← NEW SEARCH")
      .setOnClickAction(
        CardService.newAction().setFunctionName("handleNavToSearch"),
      ),
  );

  builder.addSection(section);
  return builder.build();
}

function handleNavToSearch(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(renderSearchView()))
    .build();
}

// ============================================================
// EXPORT TO SHEETS
// ============================================================

function handleExportToSheets(e) {
  var ruleId = parseInt(e.parameters.ruleId);
  var summaryRes = UrlFetchApp.fetch(BACKEND_URL + "/summary/" + ruleId, {
    muteHttpExceptions: true,
  });

  if (summaryRes.getResponseCode() !== 200) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "❌ Could not fetch data from backend.",
        ),
      )
      .build();
  }

  var summary = JSON.parse(summaryRes.getContentText());

  if (!summary.data || summary.data.length === 0) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText("⚠️ No extracted data found."),
      )
      .build();
  }

  var allKeys = [];
  summary.data.forEach(function (item) {
    Object.keys(item.fields || {}).forEach(function (k) {
      if (
        !k.startsWith("attachment_") &&
        !k.startsWith("content_") &&
        allKeys.indexOf(k) === -1
      )
        allKeys.push(k);
    });
  });

  var header = ["Subject", "Sender", "Date"].concat(allKeys);
  var rows = summary.data.map(function (item) {
    var row = [item.subject || "", item.sender || "", item.createdAt || ""];
    allKeys.forEach(function (k) {
      row.push((item.fields || {})[k] || "");
    });
    return row;
  });

  var ss = SpreadsheetApp.create(
    "Email Parser Export — " + new Date().toLocaleDateString(),
  );
  var sheet = ss.getSheets()[0];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet
    .getRange(1, 1, 1, header.length)
    .setFontWeight("bold")
    .setBackground("#4a86e8")
    .setFontColor("#ffffff");
  if (rows.length > 0)
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  sheet.autoResizeColumns(1, header.length);

  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText(
        "✅ " + rows.length + " records exported to Sheets!",
      ),
    )
    .setOpenLink(CardService.newOpenLink().setUrl(ss.getUrl()))
    .build();
}

// ============================================================
// EXPORT TO PDF — reliable HTTP download method
// ============================================================

function handleExportToPdf(e) {
  var ruleId = parseInt(e.parameters.ruleId);
  var summaryRes = UrlFetchApp.fetch(BACKEND_URL + "/summary/" + ruleId, {
    muteHttpExceptions: true,
  });

  if (summaryRes.getResponseCode() !== 200) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "❌ Could not fetch data from backend.",
        ),
      )
      .build();
  }

  var summary = JSON.parse(summaryRes.getContentText());

  if (!summary.data || summary.data.length === 0) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText("⚠️ No data to export."),
      )
      .build();
  }

  var userEmail = Session.getActiveUser().getEmail();
  var docTitle = "Email Parser Report — " + new Date().toLocaleDateString();
  var doc = DocumentApp.create(docTitle);
  var body = doc.getBody();

  body
    .appendParagraph("EMAIL PARSER REPORT")
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("Generated: " + new Date().toLocaleString());
  body.appendParagraph("User: " + userEmail);
  body.appendParagraph("Total Emails: " + summary.data.length);
  body.appendHorizontalRule();

  summary.data.forEach(function (item, idx) {
    body
      .appendParagraph(idx + 1 + ". " + (item.subject || "No Subject"))
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("From: " + (item.sender || ""));
    body.appendParagraph("Date: " + (item.createdAt || ""));
    Object.keys(item.fields || {}).forEach(function (k) {
      if (!k.startsWith("attachment_") && !k.startsWith("content_")) {
        body.appendParagraph("• " + k + ": " + item.fields[k]);
      }
    });
    body.appendHorizontalRule();
  });

  doc.saveAndClose();
  var docId = doc.getId();

  Utilities.sleep(2000);

  try {
    // Download PDF via HTTP — 100% reliable in add-ons (no Drive API weirdness)
    var url =
      "https://docs.google.com/feeds/download/documents/export/Export?id=" +
      docId +
      "&exportFormat=pdf";
    var token = ScriptApp.getOAuthToken();
    var pdfBlob = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + token },
    })
      .getBlob()
      .setName(docTitle + ".pdf");

    var pdfFile = DriveApp.createFile(pdfBlob);
    DriveApp.getFileById(docId).setTrashed(true);

    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText("📄 PDF created successfully!"),
      )
      .setOpenLink(CardService.newOpenLink().setUrl(pdfFile.getUrl()))
      .build();
  } catch (err) {
    DriveApp.getFileById(docId).setTrashed(true);
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "⚠️ PDF failed. Try Export to Sheets.",
        ),
      )
      .build();
  }
}

// ============================================================
// DIRECT MODE
// ============================================================

function handleDirectMode(query, targetFields, searchThread, searchHtml) {
  var threads = GmailApp.search(query, 0, MAX_THREADS);

  if (threads.length === 0) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "⚠️ No emails found for this query.",
        ),
      )
      .build();
  }

  var fieldKeys = Object.keys(targetFields);
  var headers = ["Subject", "Sender", "Date"].concat(fieldKeys);
  var rows = [];

  for (var t = 0; t < threads.length; t++) {
    var msgs = searchThread
      ? threads[t].getMessages()
      : [threads[t].getMessages()[threads[t].getMessages().length - 1]];

    for (var m = 0; m < msgs.length; m++) {
      var message = msgs[m];
      var body = searchHtml ? message.getBody() : message.getPlainBody();
      var row = [
        message.getSubject(),
        message.getFrom(),
        message.getDate().toLocaleDateString(),
      ];

      fieldKeys.forEach(function (k) {
        var def = targetFields[k];
        var val = "";
        if (typeof def === "string") {
          try {
            var rx = new RegExp(def, "i");
            var match = body.match(rx);
            if (match) val = match[1] || match[0];
          } catch (ex) {}
        } else if (def && def.type === "entire") {
          val = body.substring(0, 500);
        } else if (def) {
          val = parseTextAnchor(body, def.anchor, def.type, def.end);
        }
        row.push(val);
      });
      rows.push(row);
    }
  }

  if (rows.length === 0) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "⚠️ No email content could be extracted.",
        ),
      )
      .build();
  }

  var ss = SpreadsheetApp.create(
    "Email Parser — " + new Date().toLocaleDateString(),
  );
  var sheet = ss.getSheets()[0];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#4a86e8")
    .setFontColor("#ffffff");
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);

  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText(
        "✅ " +
          rows.length +
          " email" +
          (rows.length === 1 ? "" : "s") +
          " written to Sheet!",
      ),
    )
    .setOpenLink(CardService.newOpenLink().setUrl(ss.getUrl()))
    .build();
}

// ============================================================
// HELPERS
// ============================================================

function toGmailDate(dateStr) {
  var parts = dateStr.split("/");
  if (parts.length === 3) return parts[2] + "/" + parts[0] + "/" + parts[1];
  return dateStr;
}
