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

  // Save current search
  var saved = getSavedSearches();
  var saveSection = CardService.newCardSection()
    .setHeader("Save current search")
    .addWidget(
      CardService.newTextInput()
        .setFieldName("search_name")
        .setTitle("Name")
        .setHint("Name your search"),
    )
    .addWidget(
      CardService.newTextButton()
        .setText("submit")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(
          CardService.newAction().setFunctionName("handleSaveSearch"),
        ),
    );

  if (saved.length >= PLAN_SAVED_LIMIT) {
    saveSection.addWidget(
      CardService.newTextParagraph().setText(
        "⚠️ Saved search limit reached (" +
          PLAN_SAVED_LIMIT +
          " on Free plan).",
      ),
    );
  }
  builder.addSection(saveSection);

  // Saved searches list
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

  // Schedule
  var schedSection = CardService.newCardSection()
    .setHeader("Scheduled Parsing")
    .addWidget(
      CardService.newTextParagraph().setText(
        "Schedules run at midnight and parse emails from the previous 24 hours. Business plans run hourly, Premium every 4 hours.",
      ),
    )
    .addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText("Hourly")
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName("setSchedule")
                .setParameters({ minutes: "60" }),
            ),
        )
        .addButton(
          CardService.newTextButton()
            .setText("Every 4h")
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName("setSchedule")
                .setParameters({ minutes: "240" }),
            ),
        ),
    )
    .addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText("Daily")
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName("setSchedule")
                .setParameters({ minutes: "1440" }),
            ),
        )
        .addButton(
          CardService.newTextButton()
            .setText("Stop")
            .setOnClickAction(
              CardService.newAction().setFunctionName("clearSchedule"),
            ),
        ),
    );
  builder.addSection(schedSection);

  // Re-auth
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

function handleSaveSearch(e) {
  var name = e.formInput.search_name;
  if (!name) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText("⚠️ Please enter a name."),
      )
      .build();
  }
  var saved = getSavedSearches();
  if (saved.length >= PLAN_SAVED_LIMIT) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText("❌ Saved search limit reached."),
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
          "⚠️ Submit a search first before saving.",
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
    .setNotification(CardService.newNotification().setText("✅ Saved: " + name))
    .setNavigation(CardService.newNavigation().updateCard(renderSettingsView()))
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
    .setNavigation(CardService.newNavigation().updateCard(renderSettingsView()))
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

function setSchedule(e) {
  var minutes = parseInt(e.parameters.minutes);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "processEmails") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("processEmails")
    .timeBased()
    .everyMinutes(minutes)
    .create();
  var label =
    minutes === 60 ? "hourly" : minutes === 240 ? "every 4 hours" : "daily";
  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText("✅ Auto-sync set to " + label),
    )
    .build();
}

function clearSchedule() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "processEmails") ScriptApp.deleteTrigger(t);
  });
  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText("🚫 Auto-sync stopped."),
    )
    .build();
}
