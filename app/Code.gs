// ============================================================
// CONSTANTS
// ============================================================

const BACKEND_URL = "https://extract-email-parser.vercel.app";
const MAX_THREADS = 50; // Safety cap to stay within 30s execution limit

// ============================================================
// ENTRY POINTS
// ============================================================

function buildHomePage() {
  return renderSearchView();
}

function onGmailMessageOpen() {
  return renderSearchView();
}

// ============================================================
// TOP NAV
// ============================================================

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

function handleNavToSearch() {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(renderSearchView()))
    .build();
}