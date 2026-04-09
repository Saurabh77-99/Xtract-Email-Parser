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
          .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.BOOKMARK))
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

  // Fixed: Apps Script only allows everyMinutes(1, 5, 10, 15, 30)
  // Daily uses everyDays(1).atHour(0) instead of everyMinutes(1440)
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
              .setText("Every 30 min")
              .setOnClickAction(
                CardService.newAction()
                  .setFunctionName("setSchedule")
                  .setParameters({ freq: "30min" }),
              ),
          )
          .addButton(
            CardService.newTextButton()
              .setText("Daily at midnight")
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
      .setNotification(CardService.newNotification().setText("⚠️ Please enter a name for this search."))
      .build();
  }

  var allRules = fetchFromBackend("/rules");
  var latest = allRules && allRules.length > 0 ? allRules[allRules.length - 1] : null;
  if (!latest) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("⚠️ Submit a search first before saving it as a template."))
      .build();
  }

  var saved = getSavedSearches();
  var duplicate = saved.some(function (s) { return s.name === name; });
  if (duplicate) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("⚠️ A saved search named \"" + name + "\" already exists."))
      .build();
  }

  saved.push({ id: "s_" + new Date().getTime(), name: name, query: latest.criteriaQuery, ruleId: latest.id });
  PropertiesService.getUserProperties().setProperty("saved_searches", JSON.stringify(saved));

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("✅ Saved: " + name))
    .setNavigation(CardService.newNavigation().updateCard(renderSettingsView()))
    .build();
}

function handleDeleteSaved(e) {
  var saved = getSavedSearches().filter(function (s) { return s.id !== e.parameters.searchId; });
  PropertiesService.getUserProperties().setProperty("saved_searches", JSON.stringify(saved));
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("Deleted."))
    .setNavigation(CardService.newNavigation().updateCard(renderSettingsView()))
    .build();
}

function getSavedSearches() {
  var raw = PropertiesService.getUserProperties().getProperty("saved_searches");
  try { return raw ? JSON.parse(raw) : []; } catch (e) { return []; }
}

// ============================================================
// SCHEDULE — uses only valid Apps Script intervals
// ============================================================

function setSchedule(e) {
  var freq = e.parameters.freq;

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "processEmails") ScriptApp.deleteTrigger(t);
  });

  if (freq === "30min") {
    ScriptApp.newTrigger("processEmails").timeBased().everyMinutes(30).create();
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("✅ Auto-sync set to every 30 minutes."))
      .build();
  } else if (freq === "daily") {
    ScriptApp.newTrigger("processEmails").timeBased().everyDays(1).atHour(0).create();
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("✅ Auto-sync set to daily at midnight."))
      .build();
  }

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("⚠️ Unknown schedule type."))
    .build();
}

function clearSchedule() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "processEmails") ScriptApp.deleteTrigger(t);
  });
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("⛔ Auto-sync stopped."))
    .build();
}