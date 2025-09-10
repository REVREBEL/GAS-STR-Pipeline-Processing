/**
 * Test A: Build in library, show in container.
 * - Proves the template exists & compiles (library),
 *   AND that container UI can show dialogs.
 */
function test_BuildInLibrary_ShowHere() {
  var html = RevRebelGlobalMessagesLibrary.buildUserMessageDialog('Test A', 'Built in library, shown by container.');
  SpreadsheetApp.getUi().showModalDialog(html, 'Test A');
}

/**
 * Test B: Ask the library to do its own UI call.
 * - Proves library has script.container.ui scope AND can access container UI when called from here.
 * - If this fails but Test A succeeds, it’s a scope/version issue for the library.
 */
function test_LibraryCallsItsOwnUI() {
  var result = RevRebelGlobalMessagesLibrary.libraryCallsShowUserMessageForTest();
  if (!result || !result.ok) {
    SpreadsheetApp.getActive().toast('Library UI call failed: ' + (result && result.error), 'Test B', 8);
    Logger.log(result);
  }
}

/**
 * Test C: Pure template self-test (no UI) returned from library.
 * - Useful to see detailed error text if the template file isn’t found or has script errors.
 */
function test_LibraryTemplateSelfTest() {
  var info = RevRebelGlobalMessagesLibrary.libraryTemplateSelfTest();
  SpreadsheetApp.getActive().toast(JSON.stringify(info), 'Test C', 8);
  Logger.log(info);
}
