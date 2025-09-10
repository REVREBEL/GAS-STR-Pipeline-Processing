/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  CREATE USER MENU
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/**
 * Adds a custom menu to the Google Sheets UI and maps menu items to functions from the starReportProcessing library.
 */

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // Create the custom menu
  const menu = ui.createMenu('[ DATA PIPELINE MENU ]');

  // Add menu items mapped to starReportProcessing functions
  menu.addItem("Load Pipeline Settings", "saveFolderIdsFromNamedRanges");
  menu.addItem("Format + Rename Files", "renameAllStarReports");
  menu.addItem("Create/Validate Folder Structure", "buildProcessedStarFolderTree");
  menu.addSeparator();

  menu.addItem("CLEAN Data", "starReportProcessing.cleanUpData");

  menu.addSeparator();
  menu.addItem("GET Response Data", "starReportProcessing.copyResponseData");
  menu.addItem("CLEAR Data Sheets", "starReportProcessing.clearSheetsExceptSpecified");

  menu.addSeparator();
  menu.addItem("PROCESS Star Dashboard File", "moveStarDashboards");
  menu.addItem("GET Folder List", "callGetFolderList");

  menu.addSeparator();
  menu.addItem("CLEAR Saved Values", "clearAllScriptProperties");

  menu.addSeparator()
  menu.addItem("── GET OR ADD SETTINGS ──", "doNothing") // Fake header
  menu.addItem("DISPLAY Stored FoldersIDs", "getAllStoredFolderIds")
  menu.addItem("SET FolderIDs", "setAllFolderIds")
  menu.addItem("CLEAR All Settings", "clearAllScriptProperties")
  menu.addSeparator()

  menu.addSeparator()
  menu.addItem("── WRITE DATA TO SHEET ──", "doNothing") // Fake header
  menu.addItem("GET Star FoldersIDs", "listStarDataSubfoldersInFolder3Levels")
  menu.addItem("GET Demand FoldersIDs", "listDemandDataSubfoldersInFolder")
  menu.addItem("GET Pace FoldersIDs", "listPaceDataSubfoldersInFolder3Levels")
  menu.addSeparator()

  menu.addSeparator();
  menu.addItem("HELP", "help_");

  // Add the menu to the UI
  menu.addToUi();
}




/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  CLEAR SCRIPT PROPERTIES
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/**
 * Clears all script properties.
 */
function clearAllScriptProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();

  // Delete all properties
  scriptProperties.deleteAllProperties();

  Logger.log('All script properties have been cleared.');
}






/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  HELP MENU
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

function help_() {
  const ref1 = "Data Pipline";
  const ref2 = "Toolkit for Reporting Automation";
  RevRebelGlobalMessagesLibrary.showHelp(ref1, ref2);
}

function addProjectMetadata() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addDeveloperMetadata('Aparium Hotel Group | Budget Worksheet', SpreadsheetApp.DeveloperMetadataVisibility.PROJECT);
}


