const BACKEND_URL = "https://extract-email-parser.vercel.app/";
// const PLAN_DAILY_LIMIT = 10;
// const PLAN_MONTHLY_LIMIT = 50;
// const PLAN_FIELD_LIMIT = 3;
// const PLAN_SAVED_LIMIT = 3;

function buildHomePage() {
  return renderSearchView();
}
function onGmailMessageOpen() {
  return renderSearchView();
}

function createTopNav(active) {
  var grid = CardService.newGrid()
    .setNumColumns(3)
    .addItem(
      CardService.newGridItem()
        .setSubtitle(active === "search" ? "► SEARCH" : "SEARCH")
        .setTextAlignment(CardService.HorizontalAlignment.CENTER)
        .setIdentifier("nav_search"),
    )
    .addItem(
      CardService.newGridItem()
        .setSubtitle(active === "settings" ? "► SETTINGS" : "SETTINGS")
        .setTextAlignment(CardService.HorizontalAlignment.CENTER)
        .setIdentifier("nav_settings"),
    )
    .addItem(
      CardService.newGridItem()
        .setSubtitle(active === "more" ? "► MORE" : "MORE")
        .setTextAlignment(CardService.HorizontalAlignment.CENTER)
        .setIdentifier("nav_more"),
    )
    .setOnClickAction(
      CardService.newAction().setFunctionName("handleNavClick"),
    );

  return CardService.newCardSection()
    .setCollapsible(false)
    .addWidget(grid)
    .addWidget(CardService.newDivider());
}

function handleNavClick(e) {
  var id = e.parameters.grid_item_identifier;
  var nav = CardService.newNavigation();
  if (id === "nav_search")
    return CardService.newActionResponseBuilder()
      .setNavigation(nav.updateCard(renderSearchView()))
      .build();
  if (id === "nav_settings")
    return CardService.newActionResponseBuilder()
      .setNavigation(nav.updateCard(renderSettingsView()))
      .build();
  if (id === "nav_more")
    return CardService.newActionResponseBuilder()
      .setNavigation(nav.updateCard(renderMoreView()))
      .build();
}
