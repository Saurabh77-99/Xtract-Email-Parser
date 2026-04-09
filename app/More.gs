// ============================================================
// MORE VIEW
// ============================================================

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

  // ── Search tips ────────────────────────────────────────────
  builder.addSection(
    CardService.newCardSection()
      .setHeader("Search tips")
      .addWidget(
        CardService.newTextParagraph().setText(
          "• When returning a large amount of emails, break them down into chunks " +
          "using the date filter to avoid execution timeouts.\n\n" +
          "• To exclude certain keywords, add a minus operator: desired search -keyword\n\n" +
          "• To search by label, use: label:[your label]\n\n" +
          "• Use pipes | to match variations — wrap in parentheses: (Invoice|Bill)\n\n" +
          "• Avoid searching HTML emails unless plain body is not present — HTML tags " +
          "can cause unpredictable parsing results.",
        ),
      ),
  );

  // ── Scheduled parsing tips ─────────────────────────────────
  builder.addSection(
    CardService.newCardSection()
      .setHeader("Scheduled parsing")
      .addWidget(
        CardService.newTextParagraph().setText(
          "• Schedules automatically parse emails matching your active rules.\n\n" +
          "• Already-parsed emails are skipped via Gmail labels (Parsed/Rule_N).\n\n" +
          "• If spreadsheets have different timezones, schedule execution time may vary.\n\n" +
          "• You can set hourly, every 4-hour, or daily schedules in Settings.",
        ),
      ),
  );

  // ── Extraction tips ────────────────────────────────────────
  builder.addSection(
    CardService.newCardSection()
      .setHeader("Extraction tips")
      .addWidget(
        CardService.newTextParagraph().setText(
          "• Text After: finds your anchor text and captures what comes after it.\n\n" +
          "• Text Before: captures the word/line just before your anchor text.\n\n" +
          "• Regex: use a capture group () to extract a specific part, e.g. Amount: ([\\d.]+)\n\n" +
          "• Entire Email: pastes the first 500 characters of the email body into the cell.\n\n" +
          "• End of Word captures up to the next space — good for amounts and codes.\n\n" +
          "• End of Line captures to the next newline — good for full names or addresses.",
        ),
      ),
  );

  return builder.build();
}