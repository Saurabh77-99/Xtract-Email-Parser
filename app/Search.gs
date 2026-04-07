// ============================================================
// SEARCH VIEW — Extractor style
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
          .addItem("Search within each thread", "yes", false),
      )
      .addWidget(
        CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.CHECK_BOX)
          .setFieldName("f_html")
          .addItem(
            "Search email as HTML (only use when plainbody not present)",
            "yes",
            false,
          ),
      ),
  );

  // Extract Data — Field 1
  builder.addSection(
    CardService.newCardSection()
      .setHeader("Extract Data")
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

  // Field 2
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

  // Field 3
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
          .addItem("Write to new Google Sheet", "sheet", true)
          .addItem("Write directly (no backend storage)", "direct", false),
      )
      .addWidget(
        CardService.newTextParagraph().setText(
          "By submitting, the requested data will overwrite the current sheet from the top left cell.",
        ),
      )
      .addWidget(
        CardService.newTextButton()
          .setText("submit")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(
            CardService.newAction().setFunctionName("handleSearchSubmit"),
          ),
      ),
  );

  return builder.build();
}

// ============================================================
// SUBMIT
// ============================================================

function handleSearchSubmit(e) {
  // if (!checkDailyLimit()) {
  //   return CardService.newActionResponseBuilder()
  //     .setNotification(
  //       CardService.newNotification().setText(
  //         "❌ Daily import limit reached (" +
  //           PLAN_DAILY_LIMIT +
  //           "/day on Free plan).",
  //       ),
  //     )
  //     .build();
  // }
  // if (!checkMonthlyLimit()) {
  //   return CardService.newActionResponseBuilder()
  //     .setNotification(
  //       CardService.newNotification().setText(
  //         "❌ Monthly import limit reached (" + PLAN_MONTHLY_LIMIT + "/month).",
  //       ),
  //     )
  //     .build();
  // }

  var fi = e.formInput;
  var subject = fi.f_subject || "";
  var sender = fi.f_sender || "";
  var startDate = fi.f_start || "";
  var endDate = fi.f_end || "";
  var outputMode = fi.output_mode || "sheet";

  // Build query
  var query = "";
  if (subject) query += 'subject:"' + subject + '" ';
  if (sender) query += "from:" + sender + " ";
  if (startDate) query += "after:" + toGmailDate(startDate) + " ";
  if (endDate) query += "before:" + toGmailDate(endDate) + " ";
  query = query.trim() || "in:inbox";

  // Build targetFields from up to 3 field definitions
  var targetFields = {};
  for (var i = 1; i <= 3; i++) {
    var col = fi["col_" + i];
    var type = fi["type_" + i];
    var anchor = fi["anchor_" + i];
    var end = fi["end_" + i];
    if (col && anchor) {
      var key = col.toLowerCase().replace(/\s+/g, "_");
      if (type === "entire") {
        targetFields[key] = { type: "entire", anchor: "", end: "paragraph" };
      } else if (type === "regex") {
        targetFields[key] = anchor;
      } else {
        targetFields[key] = {
          type: type || "after",
          anchor: anchor,
          end: end || "word",
        };
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

  // DIRECT MODE — no backend, write straight to Sheets
  if (outputMode === "direct") {
    return handleDirectMode(query, targetFields, fi);
  }

  // BACKEND MODE — save rule + background sync
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
          "❌ Backend error. Check settings.",
        ),
      )
      .build();
  }

  var allRules = fetchFromBackend("/rules");
  var newRule = null;
  if (allRules)
    allRules.forEach(function (r) {
      if (r.name === ruleName && (!newRule || r.id > newRule.id)) newRule = r;
    });
  if (!newRule) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText("❌ Could not retrieve rule."),
      )
      .build();
  }

  var jobId = "job_" + new Date().getTime();
  var props = PropertiesService.getUserProperties();
  props.setProperties({
    [jobId + "_ruleId"]: newRule.id.toString(),
    [jobId + "_query"]: query,
    [jobId + "_status"]: "running",
    [jobId + "_count"]: "0",
  });
  props.setProperty("latest_job_id", jobId);

  ScriptApp.newTrigger("runBackgroundSync").timeBased().after(500).create();

  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText("🚀 Import started!"),
    )
    .setNavigation(
      CardService.newNavigation().pushCard(renderJobCard(jobId, ruleName)),
    )
    .build();
}

// ============================================================
// DIRECT MODE — no backend, writes to Sheet inline
// ============================================================

function handleDirectMode(query, targetFields, fi) {
  var threads = GmailApp.search(query, 0, 10);
  var rows = [];
  var headers = ["Subject", "Sender", "Date"];

  var fieldKeys = Object.keys(targetFields);
  fieldKeys.forEach(function (k) {
    headers.push(k);
  });

  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (message) {
      var body = message.getPlainBody();
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
            var m = body.match(new RegExp(def, "i"));
            if (m) val = m[1] || m[0];
          } catch (e) {}
        } else if (def && def.type === "entire") {
          val = body.substring(0, 500);
        } else if (def) {
          val = parseTextAnchor(body, def.anchor, def.type, def.end);
        }
        row.push(val);
      });
      rows.push(row);
    });
  });

  if (rows.length === 0) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText("⚠️ No emails found."),
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

  incrementImportCount(rows.length);

  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText(
        "✅ " + rows.length + " emails written to Sheet!",
      ),
    )
    .setOpenLink(CardService.newOpenLink().setUrl(ss.getUrl()))
    .build();
}

// ============================================================
// JOB STATUS CARD
// ============================================================

function renderJobCard(jobId, jobName) {
  var builder = CardService.newCardBuilder();
  builder.setHeader(
    CardService.newCardHeader()
      .setTitle("Email Parser")
      .setSubtitle("Importing..."),
  );
  builder.addSection(createTopNav("search"));

  var job = getJobStatus(jobId);
  var isDone = job.status === "done";
  var statusText = isDone
    ? "✅ Done! " + job.count + " emails imported."
    : "⏳ Processing emails... click Check Status in about 1 minute.";

  var section = CardService.newCardSection()
    .setHeader(jobName || "Import")
    .addWidget(CardService.newTextParagraph().setText(statusText));

  if (isDone && parseInt(job.count) > 0) {
    section.addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText("📊 EXPORT TO SHEETS")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName("handleExportJob")
                .setParameters({ jobId: jobId }),
            ),
        )
        .addButton(
          CardService.newTextButton()
            .setText("📄 EXPORT TO PDF")
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName("handleExportJobPdf")
                .setParameters({ jobId: jobId }),
            ),
        ),
    );
  } else if (!isDone) {
    section.addWidget(
      CardService.newTextButton()
        .setText("🔄 CHECK STATUS")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(
          CardService.newAction()
            .setFunctionName("handleCheckJobStatus")
            .setParameters({ jobId: jobId, jobName: jobName || "" }),
        ),
    );
  } else {
    section.addWidget(
      CardService.newTextParagraph().setText(
        "⚠️ No emails matched the filter.",
      ),
    );
  }

  builder.addSection(section);
  return builder.build();
}

function handleCheckJobStatus(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(
      CardService.newNavigation().updateCard(
        renderJobCard(e.parameters.jobId, e.parameters.jobName),
      ),
    )
    .build();
}

function handleExportJob(e) {
  var job = getJobStatus(e.parameters.jobId);
  var summaryRes = UrlFetchApp.fetch(BACKEND_URL + "/summary/" + job.ruleId, {
    muteHttpExceptions: true,
  });
  var summary = JSON.parse(summaryRes.getContentText());

  if (!summary.data || summary.data.length === 0) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText("⚠️ No data extracted."),
      )
      .build();
  }

  var ss = SpreadsheetApp.create(
    "Email Parser Export — " + new Date().toLocaleDateString(),
  );
  var sheet = ss.getSheets()[0];

  var allKeys = [];
  summary.data.forEach(function (item) {
    Object.keys(item.fields).forEach(function (k) {
      if (
        !k.startsWith("attachment_") &&
        !k.startsWith("content_") &&
        allKeys.indexOf(k) === -1
      )
        allKeys.push(k);
    });
  });

  var header = ["Subject", "Sender", "Date"].concat(allKeys);
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet
    .getRange(1, 1, 1, header.length)
    .setFontWeight("bold")
    .setBackground("#4a86e8")
    .setFontColor("#ffffff");

  var rows = summary.data.map(function (item) {
    var row = [item.subject || "", item.sender || "", item.createdAt || ""];
    allKeys.forEach(function (k) {
      row.push(item.fields[k] || "");
    });
    return row;
  });
  if (rows.length > 0)
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  sheet.autoResizeColumns(1, header.length);

  incrementImportCount(rows.length);
  clearJobProperties(e.parameters.jobId);

  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText(
        "✅ " + rows.length + " records exported!",
      ),
    )
    .setOpenLink(CardService.newOpenLink().setUrl(ss.getUrl()))
    .build();
}

function handleExportJobPdf(e) {
  var job = getJobStatus(e.parameters.jobId);
  var summaryRes = UrlFetchApp.fetch(BACKEND_URL + "/summary/" + job.ruleId, {
    muteHttpExceptions: true,
  });
  var summary = JSON.parse(summaryRes.getContentText());

  if (!summary.data || summary.data.length === 0) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("⚠️ No data."))
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
    Object.keys(item.fields).forEach(function (k) {
      if (!k.startsWith("attachment_") && !k.startsWith("content_")) {
        body.appendParagraph("• " + k + ": " + item.fields[k]);
      }
    });
    body.appendHorizontalRule();
  });

  doc.saveAndClose();
  var pdf = DriveApp.getFileById(doc.getId())
    .getBlob()
    .getAs("application/pdf");
  pdf.setName(docTitle + ".pdf");
  var pdfFile = DriveApp.createFile(pdf);

  clearJobProperties(e.parameters.jobId);

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("📄 PDF created!"))
    .setOpenLink(CardService.newOpenLink().setUrl(pdfFile.getUrl()))
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
