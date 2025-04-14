// Custom menu setup
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ad Customizer')
    .addItem('Open Tool', 'showPromoDialog')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('adCustomizerPage')
    .setTitle('Ad Promo Customizer');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Cache management for facilities list
function getFacilities() {
  const cache = CacheService.getScriptCache();
  const cacheKey = getCacheDateKey();
  
  // Try to get facilities from cache
  let facilities = null;
  const cachedData = cache.get(cacheKey);
  
  if (cachedData) {
    try {
      facilities = JSON.parse(cachedData);
      console.log('Retrieved facilities from cache');
      return facilities;
    } catch (e) {
      console.error('Error parsing cached facilities:', e);
    }
  }
  
  // If not in cache or error, get from sheet
  facilities = getFacilitiesFromSheet();
  
  // Cache the results for 6 hours
  try {
    cache.put(cacheKey, JSON.stringify(facilities), 21600);
    console.log('Cached facilities list');
  } catch (e) {
    console.error('Error caching facilities:', e);
  }
  
  return facilities;
}

// Generate cache key based on current date
function getCacheDateKey() {
  const today = new Date();
  return 'facilities_' + Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy_MM_dd");
}

// Get facilities from sheet
function getFacilitiesFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get current month sheet name
  const today = new Date();
  const monthYearSheetName = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM yyyy") + " Budgets";
  
  const budgetSheet = ss.getSheetByName(monthYearSheetName);
  if (!budgetSheet) {
    throw new Error(`Sheet "${monthYearSheetName}" not found`);
  }
  
  // Get facility data
  const data = budgetSheet.getRange("B:S").getValues();
  
  // Filter for Search campaigns where column S (index 17) is TRUE
  const facilities = data
    .filter(row => row[0] === 'Search' && row[17] === true)
    .map(row => row[1]) // Column C (index 1)
    .filter(facility => facility !== '');
    
  return facilities;
}

// Function to refresh facilities list and clear cache
function refreshFacilitiesList() {
  // Clear the cache first
  const cache = CacheService.getScriptCache();
  cache.remove(getCacheDateKey());
  console.log('Facilities cache cleared');
  
  // Get fresh data from sheet
  return getFacilitiesFromSheet();
}

// Validate promotion text length
function validatePromoText(promoText) {
  const finalText = promoText + '*';
  if (finalText.length > 30) {
    throw new Error('Promotion text exceeds 30 characters (including *)');
  }
  return finalText;
}

// Main function to save promotion to sheet
function savePromoToSheet(promoData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Google Import - DO NOT TOUCH');
    if (!sheet) {
      throw new Error('Google Import sheet not found');
    }
    
    // Get all data
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange('A1:C' + lastRow).getValues();
    
    // Find all rows that match the facility name in column A
    const matchingRows = [];
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === promoData.facility) {
        matchingRows.push(i + 1);
      }
    }
    
    if (matchingRows.length === 0) {
      throw new Error('Selected facility not found in Google Import sheet');
    }
    
    // Generate promo text
    let promoText = '';
    let moveInText = '';
    
    // Handle Move-In Special if enabled
    if (promoData.includeMoveInSpecial) {
      const moveInAmount = Number(promoData.moveInAmount);
      if (!moveInAmount || isNaN(moveInAmount) || moveInAmount <= 0 || moveInAmount > 100) {
        throw new Error('Invalid move-in amount (must be between $1 and $100)');
      }
      moveInText = `$${moveInAmount} Move-In Special`;
    }
    
    // Generate standard or custom promo text
    if (promoData.type === 'standard') {
      switch(promoData.standardPromo) {
        case 'first_month_dollar':
          promoText = '$1 First Month\'s Rent';
          break;
        case 'move_in':
          promoText = 'Move-In Specials';
          break;
        case 'online_rates':
          promoText = 'See Our Online Rates Today';
          break;
        case 'online_specials':
          promoText = 'View Online Specials';
          break;
        case 'web_exclusive':
          promoText = 'Web-Exclusive Discounts';
          break;
        case 'first_month_free':
          promoText = 'First Month Free';
          break;
        case 'military':
          promoText = 'Military Special';
          break;
        case 'first_responder':
          promoText = 'First Responder Discounts';
          break;
        default:
          throw new Error('Invalid standard promotion type');
      }
    } else {
      // Custom promo validation
      if (promoData.discountType === 'dollar') {
        const amount = Number(promoData.amount);
        if (!amount || isNaN(amount) || amount <= 0 || amount > 100) {
          throw new Error('Invalid dollar amount (must be between $1 and $100)');
        }
        
        if (promoData.months === 'lifetime') {
          promoText = `$${amount} Off Lifetime Of Rental`;
        } else {
          promoText = `$${amount} Off ${promoData.months} Month${promoData.months > 1 ? 's' : ''}`;
        }
      } else if (promoData.discountType === 'percentage') {
        const percentage = Number(promoData.percentage);
        if (!percentage || isNaN(percentage) || percentage <= 0 || percentage > 100) {
          throw new Error('Invalid percentage (must be between 1% and 100%)');
        }
        
        if (promoData.months === 'lifetime') {
          promoText = `${percentage}% Off Lifetime Of Rental`;
        } else {
          promoText = `${percentage}% Off ${promoData.months} Month${promoData.months > 1 ? 's' : ''}`;
        }
      }
    }
    
    // Combine promo texts - only use + if both parts exist
    let finalPromoText = '';
    if (moveInText && promoText) {
      // Use shortened version when combined with another promotion, but keep dollar sign
      const moveInAmount = Number(promoData.moveInAmount);
      finalPromoText = `${moveInAmount} Move-In+${promoText}`;
    } else if (moveInText) {
      finalPromoText = moveInText; // Use full "$XX Move-In Special" when standalone
    } else {
      finalPromoText = promoText;
    }
    
    // Add lock promotion if selected
    if (promoData.includeLock) {
      finalPromoText += ' +Free Lock';
    }

    // Validate and add asterisk
    finalPromoText = validatePromoText(finalPromoText);
    
    // Update all matching rows
    matchingRows.forEach(row => {
      sheet.getRange(row, 3).setValue(finalPromoText);
    });
    
    return true;
    
  } catch (error) {
    throw new Error('Failed to save promotion: ' + error.message);
  }
}

// Main dialog UI
function showPromoDialog() {
  // Get facilities list for the template
  const facilities = getFacilities();
  
  // Create and show the dialog
  const template = HtmlService.createTemplateFromFile('adCustomizerPage');
  template.facilities = facilities;
  const html = template.evaluate()
    .setWidth(450)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Promo Editor');
}
