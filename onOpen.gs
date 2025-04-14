// MASTER onOpen function - creates menus and initializes sheets
function onOpen() {
  // Create the Ad Customizer menu (from adCustomizer.gs)
  try {
    SpreadsheetApp.getUi()
      .createMenu('Ad Customizer')
      .addItem('Open Tool', 'showPromoDialog')
      .addToUi();
  } catch (e) {
    console.error("Error creating Ad Customizer menu:", e);
  }
  
  // Create the AdPulse Verification menu
  try {
    SpreadsheetApp.getUi()
      .createMenu('AdPulse Verification')
      .addItem('Run Verification', 'runBudgetVerification')
      .addItem('Verify Current Sheet', 'testVerifyCurrentSheet')
      .addToUi();
  } catch (e) {
    console.error("Error creating AdPulse Verification menu:", e);
  }
  
  // Initialize the AdPulse sheet if needed (from adPulseBudgetVerification.gs)
  try {
    initializeAdPulseSheet();
  } catch (e) {
    console.error("Error initializing AdPulse sheet:", e);
  }
}
