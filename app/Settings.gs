// ============================================================
// SETTINGS VIEW
// ============================================================

function renderSettingsView() {
  var builder = CardService.newCardBuilder();
  builder.setHeader(
    CardService.newCardHeader()
      .setTitle("Email Parser")
      .setImageUrl(
        "https://www.gstatic.com/images/branding/product/1x/gmail_512dp.png",
      ),
  );
  builder.addSection(createTopNav("settings"));

  var saved = getSavedSearches();

  builder.addSection(
    CardService.newCardSection()
      .setHeader("Save current search as template")
      .addWidget(
        CardService.newTextInput()
          .setFieldName("search_name")
          .setTitle("Name")
          .setHint("e.g., Monthly Google Invoices"),
      )
      .addWidget(
        CardService.newTextButton()
          .setText("Save")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(
            CardService.newAction().setFunctionName("handleSaveSearch"),
          ),
      ),
  );

  if (saved.length > 0) {
    var listSection = CardService.newCardSection().setHeader("Saved Searches");
    saved.forEach(function (s) {
      listSection.addWidget(
        CardService.newDecoratedText()
          .setText(s.name)
          .setBottomLabel(s.query || "")
          .setStartIcon(
            CardService.newIconImage().setIcon(CardService.Icon.BOOKMARK),
          )
          .setButton(
            CardService.newTextButton()
              .setText("DELETE")
              .setOnClickAction(
                CardService.newAction()
                  .setFunctionName("handleDeleteSaved")
                  .setParameters({ searchId: s.id }),
              ),
          ),
      );
    });
    builder.addSection(listSection);
  }

  // Add-on triggers minimum interval is 1 hour — everyMinutes() not allowed
  builder.addSection(
    CardService.newCardSection()
      .setHeader("Scheduled Parsing")
      .addWidget(
        CardService.newTextParagraph().setText(
          "Automatically parse emails matching your saved rules. " +
            "Emails already parsed are skipped via Gmail labels.",
        ),
      )
      .addWidget(
        CardService.newButtonSet()
          .addButton(
            CardService.newTextButton()
              .setText("Every 1 hour")
              .setOnClickAction(
                CardService.newAction()
                  .setFunctionName("setSchedule")
                  .setParameters({ freq: "1hour" }),
              ),
          )
          .addButton(
            CardService.newTextButton()
              .setText("Every 4 hours")
              .setOnClickAction(
                CardService.newAction()
                  .setFunctionName("setSchedule")
                  .setParameters({ freq: "4hours" }),
              ),
          )
          .addButton(
            CardService.newTextButton()
              .setText("Daily at 8 AM")
              .setOnClickAction(
                CardService.newAction()
                  .setFunctionName("setSchedule")
                  .setParameters({ freq: "daily" }),
              ),
          ),
      )
      .addWidget(
        CardService.newTextButton()
          .setText("⛔ Stop Schedule")
          .setOnClickAction(
            CardService.newAction().setFunctionName("clearSchedule"),
          ),
      ),
  );

  builder.addSection(
    CardService.newCardSection()
      .setHeader("Permissions")
      .addWidget(
        CardService.newTextButton()
          .setText("🔑 RE-AUTHORIZE")
          .setOnClickAction(
            CardService.newAction().setFunctionName("triggerAuthorization"),
          ),
      ),
  );

  return builder.build();
}

// ============================================================
// SAVE / DELETE SEARCH
// ============================================================

function handleSaveSearch(e) {
  var name = (e.formInput.search_name || "").trim();
  if (!name) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "⚠️ Please enter a name for this search.",
        ),
      )
      .build();
  }

  var allRules = fetchFromBackend("/rules");
  var latest =
    allRules && allRules.length > 0 ? allRules[allRules.length - 1] : null;
  if (!latest) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "⚠️ Submit a search first before saving it as a template.",
        ),
      )
      .build();
  }

  var saved = getSavedSearches();
  var duplicate = saved.some(function (s) {
    return s.name === name;
  });
  if (duplicate) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          '⚠️ A saved search named "' + name + '" already exists.',
        ),
      )
      .build();
  }

  saved.push({
    id: "s_" + new Date().getTime(),
    name: name,
    query: latest.criteriaQuery,
    ruleId: latest.id,
  });
  PropertiesService.getUserProperties().setProperty(
    "saved_searches",
    JSON.stringify(saved),
  );

  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText("✅ Saved: " + name),
    )
    .setNavigation(
      CardService.newNavigation().updateCard(renderSettingsView()),
    )
    .build();
}

function handleDeleteSaved(e) {
  var saved = getSavedSearches().filter(function (s) {
    return s.id !== e.parameters.searchId;
  });
  PropertiesService.getUserProperties().setProperty(
    "saved_searches",
    JSON.stringify(saved),
  );
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("Deleted."))
    .setNavigation(
      CardService.newNavigation().updateCard(renderSettingsView()),
    )
    .build();
}

function getSavedSearches() {
  var raw = PropertiesService.getUserProperties().getProperty("saved_searches");
  try {
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

// ============================================================
// SCHEDULE — Add-on triggers minimum is 1 hour
// everyMinutes() is NOT allowed in Gmail Add-ons
// ============================================================

function setSchedule(e) {
  var freq = e.parameters.freq;

  // Clear any existing processEmails triggers first
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "processEmails") ScriptApp.deleteTrigger(t);
  });

  if (freq === "1hour") {
    ScriptApp.newTrigger("processEmails")
      .timeBased()
      .everyHours(1)
      .create();
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "✅ Auto-sync set to every 1 hour.",
        ),
      )
      .build();
  } else if (freq === "4hours") {
    ScriptApp.newTrigger("processEmails")
      .timeBased()
      .everyHours(4)
      .create();
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "✅ Auto-sync set to every 4 hours.",
        ),
      )
      .build();
  } else if (freq === "daily") {
    ScriptApp.newTrigger("processEmails")
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "✅ Auto-sync set to daily at 8 AM.",
        ),
      )
      .build();
  }

  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText("⚠️ Unknown schedule type."),
    )
    .build();
}

function clearSchedule() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "processEmails") ScriptApp.deleteTrigger(t);
  });
  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText("⛔ Auto-sync stopped."),
    )
    .build();
}