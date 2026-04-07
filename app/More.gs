function renderMoreView() {
  var builder = CardService.newCardBuilder();
  builder.setHeader(
    CardService.newCardHeader()
      .setTitle("Email Parser")
      .setImageUrl(
        "https://www.gstatic.com/images/branding/product/1x/gmail_512dp.png",
      ),
  );
  builder.addSection(createTopNav("more"));

  // // Plan + limits
  // var counts = getImportCounts();
  // builder.addSection(
  //   CardService.newCardSection()
  //     .setHeader("Manage License")
  //     .addWidget(
  //       CardService.newKeyValue().setTopLabel("Plan").setContent("Free"),
  //     )
  //     .addWidget(
  //       CardService.newKeyValue()
  //         .setTopLabel("Search Fields Limit")
  //         .setContent(String(PLAN_FIELD_LIMIT)),
  //     )
  //     .addWidget(
  //       CardService.newKeyValue()
  //         .setTopLabel("Saved Search Limit")
  //         .setContent(String(PLAN_SAVED_LIMIT)),
  //     )
  //     .addWidget(
  //       CardService.newKeyValue()
  //         .setTopLabel("Import Limit (Day)")
  //         .setContent(counts.daily + " out of " + PLAN_DAILY_LIMIT),
  //     )
  //     .addWidget(
  //       CardService.newKeyValue()
  //         .setTopLabel("Import Limit (Month)")
  //         .setContent(counts.monthly + " out of " + PLAN_MONTHLY_LIMIT),
  //     )
  //     .addWidget(
  //       CardService.newTextButton()
  //         .setText("Reset Counts")
  //         .setOnClickAction(
  //           CardService.newAction().setFunctionName("handleResetCounts"),
  //         ),
  //     ),
  // );

  // Search tips
  builder.addSection(
    CardService.newCardSection()
      .setHeader("Search tips")
      .addWidget(
        CardService.newTextParagraph().setText(
          "• When returning a large amount of emails, break them down into chunks by using the date filter to avoid execution run time errors.\n\n" +
            "• To remove certain keywords from search add a minus operator in front of the keyword like: desired search -keyword\n\n" +
            "• To search by labels add an operator: function in the subject line followed by label:[your label]\n\n" +
            "• You can use pipes | to filter based on variations. Wrap entire filter terms in parenthesis: (Invoice|Bill)\n\n" +
            "• Searching HTML email is not recommended since HTML tags can lead to unpredictable parsing. Only use when plainbody content is not present.",
        ),
      ),
  );

  // Scheduled parsing tips
  builder.addSection(
    CardService.newCardSection()
      .setHeader("Scheduled parsing")
      .addWidget(
        CardService.newTextParagraph().setText(
          "• Schedules run at midnight based on the timezone of the spreadsheet and parse emails 24 hours prior.\n\n" +
            "• Schedules for Business Plans run every hour.\n\n" +
            "• Schedules for Premium Plans run every four hours.\n\n" +
            "• If spreadsheets have different timezones, schedule execution time can be affected.",
        ),
      ),
  );

  return builder.build();
}

// function handleResetCounts() {
//   var props = PropertiesService.getUserProperties();
//   props.deleteProperty("daily_count");
//   props.deleteProperty("daily_date");
//   props.deleteProperty("monthly_count");
//   props.deleteProperty("monthly_month");
//   return CardService.newActionResponseBuilder()
//     .setNotification(
//       CardService.newNotification().setText("✅ Import counts reset."),
//     )
//     .setNavigation(CardService.newNavigation().updateCard(renderMoreView()))
//     .build();
// }

// function getImportCounts() {
//   var props = PropertiesService.getUserProperties();
//   var today = new Date().toDateString();
//   var thisMonth = new Date().getMonth() + "-" + new Date().getFullYear();
//   var daily =
//     props.getProperty("daily_date") === today
//       ? parseInt(props.getProperty("daily_count") || "0")
//       : 0;
//   var monthly =
//     props.getProperty("monthly_month") === thisMonth
//       ? parseInt(props.getProperty("monthly_count") || "0")
//       : 0;
//   return { daily: daily, monthly: monthly };
// }

// function checkDailyLimit() {
//   return getImportCounts().daily < PLAN_DAILY_LIMIT;
// }
// function checkMonthlyLimit() {
//   return getImportCounts().monthly < PLAN_MONTHLY_LIMIT;
// }

// function incrementImportCount(count) {
//   var props = PropertiesService.getUserProperties();
//   var today = new Date().toDateString();
//   var thisMonth = new Date().getMonth() + "-" + new Date().getFullYear();
//   var counts = getImportCounts();
//   props.setProperties({
//     daily_date: today,
//     daily_count: String(counts.daily + count),
//     monthly_month: thisMonth,
//     monthly_count: String(counts.monthly + count),
//   });
// }
