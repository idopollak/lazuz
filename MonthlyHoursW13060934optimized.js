/**
 * MONTHLY HOURS ANALYSIS SCRIPT - OPTIMIZED VERSION
 * 
 * This Google Apps Script analyzes booking hours data for facility management,
 * providing validation and detailed usage reports for club/court bookings.
 * 
 * CORE FUNCTIONALITY:
 * 1. Data validation between primary and split booking tabs
 * 2. Creation of detailed hours/days cross-tabulation reports with usage analysis
 * 3. Calculation of capacity utilization and overall efficiency metrics
 * 
 * OPTIMIZATION FEATURES:
 * - Cached club schedule data for improved performance
 * - Optional data return from validation function
 * - Batch operations for formatting
 * - Configurable debug logging
 * - Smart data range detection
 * 
 * REPORT FEATURES:
 * - Standardized column widths (B-AF set to 70 pixels)
 * - Clear labeling with "Hours/Days" header
 * - Overall utilization percentage in summary
 * - Color-coded capacity indicators
 * 
 * REQUIRED SPREADSHEET STRUCTURE:
 * - ClubInfo tab with club configuration
 * - Monthly booking tabs in format [MMYY]e, [MMYY]i, [MMYY]na
 * - Split booking tabs in format [MMYY]e_2, [MMYY]i_2, [MMYY]na_2
 */

// Global configuration
var DEBUG_MODE = false; // Set to true for detailed logging
var CLUB_INFO_CACHE = null; // Cache for club information

/**
 * Monthly Hours Analysis Script - Google Sheets Functions
 * 
 * This script analyzes booking hours data and creates detailed usage reports.
 * 
 * MAIN FUNCTIONALITY:
 * 
 * 1. Data Validation (monthHours):
 *    - Validates data consistency between primary and secondary tabs
 *    - Compares column H totals for e/e_2, i/i_2, and na/na_2 tab pairs
 *    - Optionally returns data for reuse
 * 
 * 2. Hours Analysis Table (createHoursDaysTable):
 *    - Creates cross-tabulation of hours by time slot and day
 *    - Reads pre-split hourly data from [MMYY]e_2 tabs
 *    - Calculates available hours based on club schedule
 *    - Shows usage percentages and capacity utilization
 *    - Applies conditional formatting for visual analysis
 * 
 * DATA STRUCTURE (e_2 tabs):
 * - Column D: Date (e.g., "29-May-25")
 * - Column F: Start time (e.g., "9:00 am")
 * - Column G: End time (e.g., "10:00 am")
 * - Column H: Hours (pre-calculated, e.g., 1, 0.5)
 * 
 * CLUB CONFIGURATION (ClubInfo tab):
 * - B4: Club name
 * - D5: Max hours per hour (number of courts)
 * - A9:B25: Opening hours by day
 *   - Column A: Day names
 *   - Column B: Hours (format: HH:MM-HH:MM)
 * 
 * The e_2 tabs now contain pre-split hourly data, eliminating the need
 * for complex proportional distribution calculations.
 */

// Commented out onOpen function for future use
// function onOpen() {
//   var ui = SpreadsheetApp.getUi();
//   ui.createMenu('Month Hours')
//     .addItem('Run Month Hours', 'monthHours')
//     .addItem('Create Hours-Days Table', 'createHoursDaysTable')
//     .addToUi();
// }

/**
 * Debug logging function - only logs when DEBUG_MODE is true
 * @param {string} message - Message to log
 * @param {*} data - Optional data to log
 */
function debugLog(message, data) {
  if (DEBUG_MODE) {
    if (data !== undefined) {
      console.log(message, data);
    } else {
      console.log(message);
    }
  }
}

/**
 * Builds a cache of day of week information for a given month
 * @param {number} monthNum - Month number (1-12)
 * @param {number} year - Full year (e.g., 2025)
 * @return {Object} Cache object with day numbers as keys and day info as values
 */
function buildDayOfWeekCache(monthNum, year) {
  debugLog("Building day of week cache for month " + monthNum + "/" + year);
  
  var dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  var cache = {};
  var daysInMonth = new Date(year, monthNum, 0).getDate();
  
  for (var day = 1; day <= daysInMonth; day++) {
    var date = new Date(year, monthNum - 1, day);
    var dayIndex = date.getDay();
    
    cache[day] = {
      dayName: dayNames[dayIndex],
      dayIndex: dayIndex
    };
  }
  
  debugLog("Day of week cache built for " + daysInMonth + " days");
  return cache;
}

/**
 * monthHours Function
 * Validates data consistency between primary and secondary tabs
 * @param {boolean} returnData - Optional. If true, returns validation data instead of just displaying
 * @return {Object|undefined} If returnData is true, returns validation results object
 */
function monthHours(returnData) {
  console.log("Starting monthHours function");
  
  // Get UI instance
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Prompt for 4-digit month
  var response = ui.prompt('Month Input', 'Please enter a month (4 digits):', ui.ButtonSet.OK_CANCEL);
  
  // Check if user cancelled
  if (response.getSelectedButton() != ui.Button.OK) {
    debugLog("User cancelled the operation");
    return;
  }
  
  var monthDigits = response.getResponseText();
  debugLog("User input: " + monthDigits);
  
  // Validate input is exactly 4 digits
  if (!/^\d{4}$/.test(monthDigits)) {
    ui.alert('Invalid Input', 'Please enter exactly 4 digits.', ui.ButtonSet.OK);
    console.log("Invalid input: not 4 digits");
    return;
  }
  
  // Define tab pairs to check
  var tabPairs = [
    { primary: monthDigits + 'e', secondary: monthDigits + 'e_2', name: 'e' },
    { primary: monthDigits + 'i', secondary: monthDigits + 'i_2', name: 'i' },
    { primary: monthDigits + 'na', secondary: monthDigits + 'na_2', name: 'na' }
  ];
  
  var results = [];
  var validationData = {
    monthDigits: monthDigits,
    tabData: {},
    allValid: true
  };
  
  // Process each tab pair
  for (var i = 0; i < tabPairs.length; i++) {
    var pair = tabPairs[i];
    debugLog("Processing pair: " + pair.name);
    
    // Get primary tab
    var primarySheet = spreadsheet.getSheetByName(pair.primary);
    if (!primarySheet) {
      debugLog("Primary tab not found: " + pair.primary);
      results.push(pair.name + " comparison: Primary tab '" + pair.primary + "' not found");
      validationData.allValid = false;
      continue;
    }
    debugLog("Found primary tab: " + pair.primary);
    
    // Get secondary tab
    var secondarySheet = spreadsheet.getSheetByName(pair.secondary);
    if (!secondarySheet) {
      debugLog("Secondary tab not found: " + pair.secondary);
      results.push(pair.name + " comparison: Secondary tab '" + pair.secondary + "' not found");
      validationData.allValid = false;
      continue;
    }
    debugLog("Found secondary tab: " + pair.secondary);
    
    // Sum column H for primary tab
    var primarySum = sumColumnH(primarySheet, pair.primary);
    debugLog("Sum for " + pair.primary + " column H: " + primarySum);
    
    // Sum column H for secondary tab  
    var secondarySum = sumColumnH(secondarySheet, pair.secondary);
    debugLog("Sum for " + pair.secondary + " column H: " + secondarySum);
    
    // Compare sums
    var isEqual = primarySum === secondarySum;
    debugLog("Comparison result for " + pair.name + ": " + (isEqual ? "EQUAL" : "NOT EQUAL"));
    
    results.push(pair.name + " comparison: " + primarySum + " vs " + secondarySum + " - " + (isEqual ? "EQUAL" : "NOT EQUAL"));
    
    // Store data if requested
    if (returnData) {
      validationData.tabData[pair.name] = {
        primarySheet: primarySheet,
        secondarySheet: secondarySheet,
        primarySum: primarySum,
        secondarySum: secondarySum,
        isValid: isEqual
      };
      if (!isEqual) {
        validationData.allValid = false;
      }
    }
  }
  
  // Display results
  var resultMessage = "Results for month " + monthDigits + ":\n\n" + results.join("\n");
  ui.alert('Month Hours Results', resultMessage, ui.ButtonSet.OK);
  console.log("Completed monthHours function");
  
  // Return data if requested
  if (returnData) {
    return validationData;
  }
}

/**
 * Helper function to sum column H values in a sheet
 * @param {Sheet} sheet - The sheet to process
 * @param {string} sheetName - Name of the sheet (for logging)
 * @return {number} Sum of numeric values in column H
 */
function sumColumnH(sheet, sheetName) {
  debugLog("Summing column H for sheet: " + sheetName);
  
  // Use getDataRange for more efficient processing
  var dataRange = sheet.getDataRange();
  var lastRow = dataRange.getLastRow();
  var lastCol = dataRange.getLastColumn();
  
  debugLog("Data range in " + sheetName + ": " + lastRow + " rows, " + lastCol + " columns");
  
  if (lastRow === 0 || lastCol < 8) {
    debugLog("No data in column H of sheet " + sheetName);
    return 0;
  }
  
  // Get column H data (column 8) - only if it exists
  var values = sheet.getRange(1, 8, lastRow, 1).getValues();
  debugLog("Retrieved " + values.length + " rows from column H");
  
  var sum = 0;
  var numericCount = 0;
  
  // Sum numeric values
  for (var i = 0; i < values.length; i++) {
    var value = values[i][0];
    if (typeof value === 'number' && !isNaN(value)) {
      sum += value;
      numericCount++;
    }
  }
  
  debugLog("Found " + numericCount + " numeric values in column H of " + sheetName);
  debugLog("Total sum: " + sum);
  
  return sum;
}

/**
 * Helper function to convert 24-hour time to AM/PM format
 * @param {number} hour - Hour in 24-hour format (0-23)
 * @return {string} Time in AM/PM format
 */
function formatTimeAMPM(hour) {
  if (hour === 0) return "12AM";
  if (hour === 24) return "Midnight";
  if (hour === 12) return "12PM";
  if (hour < 12) return hour + "AM";
  return (hour - 12) + "PM";
}

/**
 * Helper function to parse hour from various time formats
 * @param {string|number|Date} timeValue - Time value from sheet
 * @return {number} Hour in 24-hour format (0-23)
 */
function parseHourFromTime(timeValue) {
  if (!timeValue && timeValue !== 0) return 0;
  
  // If it's already a Date object
  if (timeValue instanceof Date) {
    return timeValue.getHours();
  }
  
  // If it's a number (decimal time)
  if (typeof timeValue === 'number') {
    // Handle decimal time (0.5 = 12:00 PM)
    var hour = Math.floor(timeValue * 24);
    return hour;
  }
  
  // Convert to string
  var timeStr = timeValue.toString().trim();
  
  // Try to parse AM/PM format (e.g., "10AM", "10PM", "11:00 AM")
  var ampmMatch = timeStr.match(/(\d{1,2})(?::(\d{2}))?\s*(AM|PM)/i);
  if (ampmMatch) {
    var hour = parseInt(ampmMatch[1]);
    var isPM = ampmMatch[3].toUpperCase() === 'PM';
    
    if (hour === 12 && !isPM) hour = 0;  // 12AM = 0
    else if (hour !== 12 && isPM) hour += 12;  // Add 12 for PM (except 12PM)
    
    return hour;
  }
  
  // Try to parse 24-hour format (e.g., "14:00")
  var hourMatch = timeStr.match(/(\d{1,2}):/);
  if (hourMatch) {
    return parseInt(hourMatch[1]);
  }
  
  // Try to parse just a number
  var num = parseInt(timeStr);
  if (!isNaN(num)) {
    return num;
  }
  
  debugLog("Could not parse time: " + timeStr);
  return 0;
}

/**
 * Gets club info from cache or loads it if not cached
 * @return {Object} Club info including name and max hours
 */
function getClubInfo() {
  if (CLUB_INFO_CACHE) {
    debugLog("Using cached club info");
    return CLUB_INFO_CACHE;
  }
  
  debugLog("Loading club info from sheet");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var clubInfoSheet = spreadsheet.getSheetByName('ClubInfo');
  
  if (!clubInfoSheet) {
    throw new Error('ClubInfo tab not found');
  }
  
  CLUB_INFO_CACHE = {
    sheet: clubInfoSheet,
    clubName: clubInfoSheet.getRange('B4').getValue(),
    maxHoursPerHour: clubInfoSheet.getRange('D5').getValue()
  };
  
  debugLog("Club info cached", CLUB_INFO_CACHE);
  return CLUB_INFO_CACHE;
}

/**
 * calculateAvailableHours Function
 * 
 * Calculates the maximum available hours for each time slot based on:
 * - Club opening hours (from ClubInfo A9:B25)
 * - Number of courts available (maxHoursPerHour from ClubInfo D5)
 * - Day of the week for each date in the month
 * 
 * @param {number} monthNum - Month number (1-12)
 * @param {number} year - Full year (e.g., 2025)
 * @return {Object} Object containing availableGrid and metadata
 */
function calculateAvailableHours(monthNum, year) {
  debugLog("=== INSIDE CALCULATEAVAILABLEHOURS ===");
  debugLog("Calculating available hours for month " + monthNum + "/" + year);
  
  // Get club info from cache
  var clubInfo = getClubInfo();
  var clubInfoSheet = clubInfo.sheet;
  var maxHoursPerHour = clubInfo.maxHoursPerHour;
  
  debugLog("Max hours per hour (courts): " + maxHoursPerHour);
  
  // Get opening hours - check both possible formats
  debugLog("=== READING OPENING HOURS IN CALCULATEAVAILABLEHOURS ===");
  var openingHoursData = clubInfoSheet.getRange('A9:B25').getValues();
  debugLog("Read " + openingHoursData.length + " rows from A9:B25");
  
  var dayOpeningHours = {};
  var dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  var earliestHour = 24;
  var latestHour = 0;
  
  // Check if we have data in A9:B25 format
  var hasTraditionalFormat = false;
  var validRowCount = 0;
  for (var i = 0; i < openingHoursData.length; i++) {
    if (openingHoursData[i][0] && openingHoursData[i][1]) {
      validRowCount++;
      if (!hasTraditionalFormat) {
        hasTraditionalFormat = true;
        debugLog("  First valid data found at row " + i + ", setting hasTraditionalFormat = true");
      }
    }
  }
  
  debugLog("hasTraditionalFormat = " + hasTraditionalFormat + ", total valid rows = " + validRowCount);
  
  if (hasTraditionalFormat) {
    // Use traditional A9:B25 format
    debugLog("=== PARSING TRADITIONAL FORMAT IN CALCULATEAVAILABLEHOURS ===");
    debugLog("Using traditional opening hours format from A9:B25");
    
    for (var i = 0; i < openingHoursData.length; i++) {
      var dayName = openingHoursData[i][0];
      var hoursString = openingHoursData[i][1];
      
      if (dayName && hoursString) {
        debugLog("Parsing row " + (i+9) + ": Day='" + dayName + "', Hours='" + hoursString + "'");
        
        // Convert to string if needed
        var hoursStr = hoursString.toString();
        
        // Parse HH:MM-HH:MM format
        var parts = hoursStr.split('-');
        if (parts.length === 2) {
          var openTime = parts[0].trim();
          var closeTime = parts[1].trim();
          
          var openHour = parseInt(openTime.split(':')[0]);
          var closeHour = parseInt(closeTime.split(':')[0]);
          
          debugLog("  Parsed times: open=" + openHour + ", close=" + closeHour);
          
          // Handle midnight closing time (24:00 becomes 24, not 0)
          if (closeTime === "24:00" || (closeHour === 0 && closeTime.indexOf("00:00") >= 0)) {
            closeHour = 24;
          }
          
          // Validate parsed hours
          if (!isNaN(openHour) && !isNaN(closeHour)) {
            dayOpeningHours[dayName] = {
              open: openHour,
              close: closeHour
            };
            debugLog("  Successfully stored: " + dayName + " = " + JSON.stringify(dayOpeningHours[dayName]));
            
            if (openHour < earliestHour) earliestHour = openHour;
            if (closeHour > latestHour) latestHour = closeHour;
          } else {
            debugLog("  ERROR: Invalid hours for " + dayName);
          }
        }
      }
    }
    
    debugLog("=== AFTER PARSING IN CALCULATEAVAILABLEHOURS ===");
    debugLog("dayOpeningHours: " + JSON.stringify(dayOpeningHours));
    debugLog("dayOpeningHours keys: " + Object.keys(dayOpeningHours).join(", "));
  } else {
    // Use grid format from D5:K7
    debugLog("Using grid opening hours format from D5:K7");
    var openingHoursRange = clubInfoSheet.getRange('D5:K7').getValues();
    debugLog("Opening hours range data:");
    debugLog("Row 0 (headers): " + openingHoursRange[0].join(", "));
    debugLog("Row 1 (open times): " + openingHoursRange[1].join(", "));
    debugLog("Row 2 (close times): " + openingHoursRange[2].join(", "));
    
    // Column E (index 1) through K (index 7) contain the days
    for (var i = 0; i < 7; i++) {
      var dayName = dayNames[i];
      var openTime = openingHoursRange[1][i + 1];  // Row 2 (index 1) is Open times
      var closeTime = openingHoursRange[2][i + 1]; // Row 3 (index 2) is Close times
      
      debugLog("Raw opening hours for " + dayName + ": open=" + openTime + ", close=" + closeTime + " (type: " + typeof openTime + ")");
      
      // Parse times
      var openHour = parseHourFromTime(openTime);
      var closeHour = parseHourFromTime(closeTime);
      
      // Handle midnight closing
      if (closeHour === 0) {
        closeHour = 24;
        debugLog(dayName + " closes at midnight, setting closeHour to 24");
      }
      
      dayOpeningHours[dayName] = {
        open: openHour,
        close: closeHour
      };
      
      if (openHour < earliestHour) earliestHour = openHour;
      if (closeHour > latestHour) latestHour = closeHour;
      
      debugLog(dayName + ": " + openHour + " - " + closeHour);
    }
  }
  
  debugLog("=== END OF OPENING HOURS PARSING IN CALCULATEAVAILABLEHOURS ===");
  debugLog("Earliest opening hour: " + earliestHour);
  debugLog("Latest closing hour: " + latestHour);
  debugLog("Day opening hours parsed: " + JSON.stringify(dayOpeningHours));
  debugLog("Number of days with hours: " + Object.keys(dayOpeningHours).length);
  
  // Safety check - if no valid hours were found, use defaults
  if (earliestHour === 24 || latestHour === 0 || Object.keys(dayOpeningHours).length === 0) {
    console.log("ERROR: No valid opening hours found! Using default hours 8AM-11PM");
    earliestHour = 8;
    latestHour = 23; // 11 PM
    
    // Set default hours for all days
    for (var i = 0; i < dayNames.length; i++) {
      dayOpeningHours[dayNames[i]] = {
        open: 8,
        close: 23
      };
    }
    debugLog("After defaults, dayOpeningHours: " + JSON.stringify(dayOpeningHours));
  }
  
  // Build day of week cache for the month
  var dayOfWeekCache = buildDayOfWeekCache(monthNum, year);
  
  // Calculate days in month
  var daysInMonth = new Date(year, monthNum, 0).getDate();
  
  // Build hour rows
  var timeRows = [];
  var lastHourRow = latestHour === 24 ? 23 : latestHour - 1;
  
  debugLog("Building time rows from " + earliestHour + " to " + lastHourRow);
  
  // Safety check to ensure we have valid hours
  if (earliestHour > lastHourRow) {
    console.log("ERROR: Invalid hour range! Using defaults 8-23");
    earliestHour = 8;
    lastHourRow = 23;
  }
  
  for (var hour = earliestHour; hour <= lastHourRow; hour++) {
    timeRows.push(hour);
  }
  
  // Initialize available hours grid
  var availableGrid = [];
  for (var i = 0; i < timeRows.length; i++) {
    var row = [];
    for (var d = 0; d < daysInMonth; d++) {
      row.push(0);
    }
    availableGrid.push(row);
  }
  
  // Calculate available hours for each slot
  for (var hourIndex = 0; hourIndex < timeRows.length; hourIndex++) {
    var hour = timeRows[hourIndex];
    
    for (var day = 1; day <= daysInMonth; day++) {
      var dayIndex = day - 1;
      
      // Get day of week from cache
      var dayInfo = dayOfWeekCache[day];
      var dayName = dayInfo.dayName;
      
      // Check if club is open during this hour
      var openingInfo = dayOpeningHours[dayName];
      
      if (openingInfo && hour >= openingInfo.open && hour < openingInfo.close) {
        // Club is open - set available hours to max
        availableGrid[hourIndex][dayIndex] = maxHoursPerHour;
      } else {
        // Club is closed - 0 available hours
        availableGrid[hourIndex][dayIndex] = 0;
      }
    }
  }
  
  debugLog("Day of week cache built for " + daysInMonth + " days");
  
  // Prepare return object
  var returnData = {
    availableGrid: availableGrid,
    timeRows: timeRows,
    daysInMonth: daysInMonth,
    monthNum: monthNum,
    year: year,
    maxHoursPerHour: maxHoursPerHour,
    earliestHour: earliestHour,
    latestHour: latestHour,
    dayOpeningHours: dayOpeningHours,
    dayOfWeekCache: dayOfWeekCache
  };
  
  debugLog("=== ABOUT TO RETURN FROM CALCULATEAVAILABLEHOURS ===");
  debugLog("Return data keys: " + Object.keys(returnData).join(", "));
  debugLog("dayOpeningHours in return: " + JSON.stringify(returnData.dayOpeningHours));
  debugLog("dayOpeningHours keys in return: " + Object.keys(returnData.dayOpeningHours).join(", "));
  
  // Return results with metadata
  return returnData;
}

/**
 * Applies batch formatting to a range
 * @param {Sheet} sheet - The sheet to format
 * @param {Array} formatSpecs - Array of format specifications
 */
function batchFormat(sheet, formatSpecs) {
  for (var i = 0; i < formatSpecs.length; i++) {
    var spec = formatSpecs[i];
    var range = sheet.getRange(spec.range);
    
    if (spec.values) range.setValues(spec.values);
    if (spec.fontWeight) range.setFontWeight(spec.fontWeight);
    if (spec.fontSize) range.setFontSize(spec.fontSize);
    if (spec.fontColor) range.setFontColor(spec.fontColor);
    if (spec.background) range.setBackground(spec.background);
    if (spec.border) {
      range.setBorder(
        spec.border.top || false,
        spec.border.left || false,
        spec.border.bottom || false,
        spec.border.right || false,
        spec.border.vertical || false,
        spec.border.horizontal || false,
        spec.border.color || 'black',
        spec.border.style || SpreadsheetApp.BorderStyle.SOLID
      );
    }
    if (spec.horizontalAlignment) range.setHorizontalAlignment(spec.horizontalAlignment);
    if (spec.numberFormat) range.setNumberFormat(spec.numberFormat);
    if (spec.formula) range.setFormula(spec.formula);
  }
}

/**
 * createHoursDaysTable Function
 * 
 * Creates a cross-tabulation table showing booking patterns with integrated
 * availability calculation. Reads pre-split hourly data from e_2 tabs.
 */
function createHoursDaysTable() {
  console.log("Starting createHoursDaysTable function");
  
  // Get UI instance
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Prompt for 4-digit month
  var response = ui.prompt('Create Hours-Days Table', 'Please enter a month (4 digits):', ui.ButtonSet.OK_CANCEL);
  
  // Check if user cancelled
  if (response.getSelectedButton() != ui.Button.OK) {
    debugLog("User cancelled the operation");
    return;
  }
  
  var monthDigits = response.getResponseText();
  debugLog("User input: " + monthDigits);
  
  // Validate input is exactly 4 digits
  if (!/^\d{4}$/.test(monthDigits)) {
    ui.alert('Invalid Input', 'Please enter exactly 4 digits.', ui.ButtonSet.OK);
    console.log("Invalid input: not 4 digits");
    return;
  }
  
  // Parse month and year from 4 digits (MMYY format)
  var monthNum = parseInt(monthDigits.substring(0, 2));
  var year = parseInt('20' + monthDigits.substring(2, 4));
  debugLog("Parsed month: " + monthNum + ", year: " + year);
  
  // Convert month number to month name
  var monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var monthName = monthNames[monthNum - 1];
  var yearShort = year.toString().substring(2);
  
  // Get club info from cache
  var clubInfo = getClubInfo();
  var clubName = clubInfo.clubName;
  var maxHoursPerHour = clubInfo.maxHoursPerHour;
  
  debugLog("Club name: " + clubName);
  debugLog("Max hours per hour: " + maxHoursPerHour);
  
  // Calculate available hours for the month
  debugLog("=== STARTING CALCULATEAVAILABLEHOURS ===");
  var availabilityData = calculateAvailableHours(monthNum, year);
  debugLog("=== RETURNED FROM CALCULATEAVAILABLEHOURS ===");
  debugLog("availabilityData keys: " + Object.keys(availabilityData).join(", "));
  
  var availableGrid = availabilityData.availableGrid;
  var timeRows = availabilityData.timeRows;
  var daysInMonth = availabilityData.daysInMonth;
  var dayOpeningHours = availabilityData.dayOpeningHours;
  var earliestHour = availabilityData.earliestHour;
  var latestHour = availabilityData.latestHour;
  
  debugLog("=== DATA RECEIVED IN CREATEHOURSDAYSTABLE ===");
  debugLog("- dayOpeningHours: " + JSON.stringify(dayOpeningHours));
  debugLog("- dayOpeningHours keys: " + Object.keys(dayOpeningHours).join(", "));
  debugLog("- earliestHour: " + earliestHour);
  debugLog("- latestHour: " + latestHour);
  debugLog("- timeRows: " + timeRows.join(", "));
  
  console.log("Available hours grid created");
  debugLog("Time rows: " + timeRows.length + " from hour " + earliestHour + " to " + (latestHour - 1));
  debugLog("Time rows array: " + timeRows.join(", "));
  
  // Build day of week cache
  var dayOfWeekCache = buildDayOfWeekCache(monthNum, year);
  debugLog("Day of week cache created");
  
  // Get source tabs
  var sourceTabName = monthDigits + 'e_2';
  var sourceSheet = spreadsheet.getSheetByName(sourceTabName);
  
  if (!sourceSheet) {
    console.log("Source tab not found: " + sourceTabName);
    ui.alert('Error', 'Source tab ' + sourceTabName + ' not found.', ui.ButtonSet.OK);
    return;
  }
  debugLog("Found source tab: " + sourceTabName);
  
  // Get e tab for validation
  var eTabName = monthDigits + 'e';
  var eSheet = spreadsheet.getSheetByName(eTabName);
  var eTotal = 0;
  if (eSheet) {
    eTotal = sumColumnH(eSheet, eTabName);
    debugLog("Total from " + eTabName + ": " + eTotal);
  }
  
  // Get e_2 tab total
  var e2Total = sumColumnH(sourceSheet, sourceTabName);
  debugLog("Total from " + sourceTabName + ": " + e2Total);
  
  // Create or get destination tab
  var destTabName = monthName + '-' + yearShort + ' DH';
  var destSheet = spreadsheet.getSheetByName(destTabName);
  
  if (destSheet) {
    debugLog("Destination tab exists, deleting and recreating");
    spreadsheet.deleteSheet(destSheet);
  }
  
  console.log("Creating new destination tab: " + destTabName);
  destSheet = spreadsheet.insertSheet(destTabName);
  
  // Format the sheet
  debugLog("Formatting sheet");
  
  // Set column widths B to AF to 70 pixels
  debugLog("Setting column widths B to AF");
  for (var col = 2; col <= 32; col++) { // B=2, AF=32
    destSheet.setColumnWidth(col, 70);
  }
  
  // Hide gridlines
  try {
    destSheet.setHiddenGridlines(true);
    debugLog("Gridlines hidden");
  } catch (e) {
    debugLog("Could not hide gridlines: " + e.toString());
  }
  
  // Set base formatting for entire sheet
  var fullRange = destSheet.getRange(1, 1, 100, 50);
  fullRange.setFontFamily('Verdana');
  fullRange.setFontSize(10);
  fullRange.setBackground('#cbc9a2');
  
  // Prepare batch formatting specifications for header and validation sections
  var headerFormatSpecs = [
    {
      range: 'A1',
      formula: '=CONCATENATE("' + monthName + '-' + yearShort + ' Detailed Hours and Usage"," ",ClubInfo!$B$4)',
      fontWeight: 'bold',
      fontSize: 11
    },
    {
      range: 'A3:E3',
      values: [['Total E Hours', eTotal, 'Total Split Hours', e2Total, 
                Math.abs(eTotal - e2Total) > 0.01 ? 'Please Check' : 'OK']]
    }
  ];
  
  // Apply validation formatting
  if (Math.abs(eTotal - e2Total) > 0.01) {
    headerFormatSpecs.push({
      range: 'E3',
      fontWeight: 'bold',
      fontColor: '#F32C1E'
    });
  } else {
    headerFormatSpecs.push({
      range: 'E3',
      fontWeight: 'bold',
      fontColor: '#328332'
    });
  }
  
  // Max hours section
  headerFormatSpecs.push({
    range: 'A5:B5',
    values: [['Max Hours per Hour', maxHoursPerHour]]
  });
  
  // Apply header formatting
  batchFormat(destSheet, headerFormatSpecs);
  
  // Opening hours table - prepare data
  var openingHoursHeaders = [['Day', '', '', '', '', '', '', '']];
  var openingHoursOpen = [['Open', '', '', '', '', '', '', '']];
  var openingHoursClose = [['Close', '', '', '', '', '', '', '']];
  
  var dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  for (var d = 0; d < 7; d++) {
    var dayName = dayNames[d];
    openingHoursHeaders[0][d + 1] = dayName.substring(0, 3);
    
    // Find matching key for opening hours
    var foundKey = null;
    for (var key in dayOpeningHours) {
      if (key.toLowerCase().trim() === dayName.toLowerCase().trim()) {
        foundKey = key;
        break;
      }
    }
    
    if (foundKey) {
      openingHoursOpen[0][d + 1] = formatTimeAMPM(dayOpeningHours[foundKey].open);
      openingHoursClose[0][d + 1] = formatTimeAMPM(dayOpeningHours[foundKey].close);
    } else {
      openingHoursOpen[0][d + 1] = "Not Found";
      openingHoursClose[0][d + 1] = "Not Found";
    }
  }
  
  // Set opening hours data
  destSheet.getRange('D5:K5').setValues(openingHoursHeaders);
  destSheet.getRange('D6:K6').setValues(openingHoursOpen);
  destSheet.getRange('D7:K7').setValues(openingHoursClose);
  
  // Format opening hours table
  var openingHoursRange = destSheet.getRange('D5:K7');
  openingHoursRange.setBackground('#E3E2CD');
  openingHoursRange.setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  openingHoursRange.setHorizontalAlignment('left');
  
  // Find last row after header sections
  var tableStartRow = 10;
  
  // Build main table header row (days of month)
  debugLog("Building main table header row");
  var headerRow = ['Hours/Days'];  // Header for the time/days table
  
  for (var day = 1; day <= daysInMonth; day++) {
    // Format as DD-MMM
    var dateStr = ('0' + day).slice(-2) + '-' + monthName;
    headerRow.push(dateStr);
  }
  headerRow.push('Total Hours'); // Add total column
  headerRow.push(''); // Empty column (lcol1+1)
  headerRow.push('Total Available Hours'); // Add available hours column
  headerRow.push('% Utilization'); // Add utilization column
  
  // Set header row
  destSheet.getRange(tableStartRow, 1, 1, headerRow.length).setValues([headerRow]);
  destSheet.getRange(tableStartRow, 1, 1, headerRow.length).setFontWeight('bold');
  destSheet.getRange(tableStartRow, 1, 1, headerRow.length).setBorder(false, false, true, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // Build time labels
  debugLog("Building time labels");
  var timeLabels = [];
  
  // Safety check
  if (timeRows.length === 0) {
    console.log("ERROR: No time rows! This should not happen.");
    // Force some default rows
    timeRows = [8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23];
  }
  
  for (var i = 0; i < timeRows.length; i++) {
    timeLabels.push("'" + formatTimeAMPM(timeRows[i]));
  }
  
  // Initialize data grid
  var dataGrid = [];
  for (var i = 0; i < timeRows.length; i++) {
    var row = new Array(daysInMonth).fill(0);
    dataGrid.push(row);
  }
  
  // Read and process source data - SIMPLIFIED LOGIC
  debugLog("Reading source data from " + sourceTabName);
  var sourceDataRange = sourceSheet.getDataRange();
  var lastRow = sourceDataRange.getLastRow();
  var lastCol = sourceDataRange.getLastColumn();
  
  if (lastRow > 0 && lastCol > 0) {
    var sourceData = sourceDataRange.getValues();
    var processedCount = 0;
    
    console.log("Processing " + (sourceData.length - 1) + " data rows");
    
    // Process bookings - SIMPLIFIED VERSION
    for (var i = 1; i < sourceData.length; i++) {
      var rowData = sourceData[i];
      
      var dateValue = rowData[3];  // Column D
      var startTime = rowData[5];  // Column F
      var hours = rowData[7];      // Column H
      
      // Debug first few rows
      if (i <= 5 && DEBUG_MODE) {
        debugLog("Row " + i + " - Date: " + dateValue + " (type: " + typeof dateValue + 
                 "), Start: " + startTime + " (type: " + typeof startTime + 
                 "), Hours: " + hours);
      }
      
      if (dateValue && startTime && hours) {
        // Parse the date
        var bookingDate = new Date(dateValue);
        var bookingDay = bookingDate.getDate();
        var bookingMonth = bookingDate.getMonth() + 1;
        var bookingYear = bookingDate.getFullYear();
        
        // Check if booking is in the target month/year
        if (bookingMonth === monthNum && bookingYear === year) {
          // Parse the start time to get the hour
          var hour;
          
          // Handle different time formats
          if (startTime instanceof Date) {
            // If it's already a Date object (time only)
            hour = startTime.getHours();
          } else if (typeof startTime === 'number') {
            // If it's a decimal (0.5 = 12:00 PM)
            hour = Math.floor(startTime * 24);
          } else {
            // If it's a string, try parsing it
            var timeStr = startTime.toString();
            hour = parseHourFromTime(timeStr);
          }
          
          debugLog("Parsed hour " + hour + " from start time: " + startTime);
          
          // Find grid position
          var hourIndex = timeRows.indexOf(hour);
          var dayIndex = bookingDay - 1;
          
          if (hourIndex >= 0 && dayIndex >= 0 && dayIndex < daysInMonth) {
            // Simply add the hours to the grid
            dataGrid[hourIndex][dayIndex] += hours;
            processedCount++;
            
            if (processedCount <= 10 && DEBUG_MODE) {
              debugLog("Added " + hours + " hours to " + hour + ":00 on day " + bookingDay);
            }
          } else {
            debugLog("Warning: Could not place booking - hour: " + hour + ", day: " + bookingDay + 
                     ", hourIndex: " + hourIndex + ", dayIndex: " + dayIndex);
          }
        }
      }
    }
    
    console.log("Processed " + processedCount + " bookings for " + monthName + " " + year);
  }
  
  // Prepare all data rows for batch writing
  debugLog("Preparing data for batch writing");
  var allDataRows = [];
  var rowTotals = [];
  
  for (var r = 0; r < timeRows.length; r++) {
    var rowData = [timeLabels[r]];
    var rowTotal = 0;
    
    for (var c = 0; c < daysInMonth; c++) {
      var value = dataGrid[r][c];
      rowData.push(value === 0 ? '' : value);
      rowTotal += value;
    }
    
    rowData.push(rowTotal === 0 ? '' : rowTotal);
    rowTotals.push(rowTotal);
    allDataRows.push(rowData);
  }
  
  // Write all data rows at once
  if (allDataRows.length > 0) {
    destSheet.getRange(tableStartRow + 1, 1, allDataRows.length, allDataRows[0].length)
      .setValues(allDataRows);
  }
  
  // Add Total Dates row
  var totalDatesRow = ['Total Dates'];
  var grandTotal = 0;
  for (var c = 0; c < daysInMonth; c++) {
    var colTotal = 0;
    for (var r = 0; r < dataGrid.length; r++) {
      colTotal += dataGrid[r][c];
    }
    totalDatesRow.push(colTotal === 0 ? '' : colTotal);
    grandTotal += colTotal;
  }
  totalDatesRow.push(grandTotal);
  
  var totalRowIndex = tableStartRow + timeRows.length + 1;
  destSheet.getRange(totalRowIndex, 1, 1, totalDatesRow.length).setValues([totalDatesRow]);
  destSheet.getRange(totalRowIndex, 1, 1, totalDatesRow.length).setFontWeight('bold');
  destSheet.getRange(totalRowIndex, 1, 1, totalDatesRow.length).setBorder(true, false, false, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // Apply formatting
  debugLog("Applying formatting to table");
  
  // Batch apply alternating row colors
  var rowBackgrounds = [];
  for (var r = 0; r < timeRows.length; r++) {
    if (r % 2 === 0) {
      rowBackgrounds.push({
        range: destSheet.getRange(tableStartRow + 1 + r, 1, 1, headerRow.length),
        background: '#E3E2CD'
      });
    }
  }
  
  if (rowBackgrounds.length > 0) {
    for (var i = 0; i < rowBackgrounds.length; i++) {
      rowBackgrounds[i].range.setBackground(rowBackgrounds[i].background);
    }
  }
  
  // Format hours column (bold, left-aligned, right border)
  debugLog("Formatting hours column");
  var hoursColumnRange = destSheet.getRange(tableStartRow + 1, 1, timeRows.length, 1);
  hoursColumnRange.setFontWeight('bold');
  hoursColumnRange.setHorizontalAlignment('left');
  hoursColumnRange.setNumberFormat('@'); // Text format to preserve AM/PM
  hoursColumnRange.setBorder(false, false, false, true, false, false, 'black', SpreadsheetApp.BorderStyle.DOTTED);
  
  // Format total hours column (bold, left border)
  debugLog("Formatting total hours column");
  var totalHoursCol = daysInMonth + 2; // Column position for Total Hours
  var totalHoursRange = destSheet.getRange(tableStartRow + 1, totalHoursCol, timeRows.length, 1);
  totalHoursRange.setFontWeight('bold');
  totalHoursRange.setBorder(false, true, false, false, false, false, 'black', SpreadsheetApp.BorderStyle.DOTTED);
  
  // Apply conditional formatting for closed hours and max capacity
  debugLog("Applying conditional formatting");
  
  // Collect cells that need special formatting
  var closedCells = [];
  var maxCapacityCells = [];
  
  for (var r = 0; r < timeRows.length; r++) {
    var hour = timeRows[r];
    var currentRowNum = tableStartRow + 1 + r;
    
    for (var d = 1; d <= daysInMonth; d++) {
      var cellRef = destSheet.getRange(currentRowNum, d + 1);
      
      // Get available hours for this slot
      var availableHours = availableGrid[r][d - 1];
      
      // Check if club is closed (available hours = 0)
      if (availableHours === 0) {
        closedCells.push(cellRef);
      }
      
      // Check if at max capacity
      var cellValue = dataGrid[r][d - 1];
      if (cellValue && availableHours > 0 && cellValue === availableHours) {
        maxCapacityCells.push(cellRef);
      }
    }
  }
  
  // Apply closed cells formatting
  if (closedCells.length > 0) {
    for (var i = 0; i < closedCells.length; i++) {
      closedCells[i].setBackground('#DDDDD7');
    }
  }
  
  // Apply max capacity formatting
  if (maxCapacityCells.length > 0) {
    for (var i = 0; i < maxCapacityCells.length; i++) {
      maxCapacityCells[i].setBackground('#328332');
      maxCapacityCells[i].setFontWeight('bold');
      maxCapacityCells[i].setFontColor('#FFFFFF');
    }
  }
  
  // Find last column (lcol1)
  var lcol1 = daysInMonth + 2; // Last data column (Total Hours)
  debugLog("Last column (lcol1) is: " + lcol1);
  
  // Set width of column lcol1+1
  destSheet.setColumnWidth(lcol1 + 1, 10);
  debugLog("Set width of column " + (lcol1 + 1) + " to 10 pixels");
  
  // Prepare capacity and utilization data
  var capacityData = [];
  var utilizationData = [];
  
  for (var r = 0; r < timeRows.length; r++) {
    // Calculate total available hours for this time slot
    var totalAvailable = 0;
    for (var d = 0; d < daysInMonth; d++) {
      totalAvailable += availableGrid[r][d];
    }
    
    capacityData.push([totalAvailable]);
    
    // Calculate utilization percentage
    var totalHours = rowTotals[r];
    if (totalAvailable > 0) {
      utilizationData.push([totalHours / totalAvailable]);
    } else {
      utilizationData.push(['']);
    }
  }
  
  // Write capacity data
  if (capacityData.length > 0) {
    destSheet.getRange(tableStartRow + 1, lcol1 + 2, capacityData.length, 1)
      .setValues(capacityData);
  }
  
  // Make Total Available Hours column bold
  var availableHoursRange = destSheet.getRange(tableStartRow + 1, lcol1 + 2, timeRows.length, 1);
  availableHoursRange.setFontWeight('bold');
  availableHoursRange.setBorder(false, true, false, false, false, false, 'black', SpreadsheetApp.BorderStyle.DOTTED);
  
  // Write utilization data
  if (utilizationData.length > 0) {
    destSheet.getRange(tableStartRow + 1, lcol1 + 3, utilizationData.length, 1)
      .setValues(utilizationData)
      .setNumberFormat('0.0%');
  }
  
  // Make % Utilization column bold
  var utilizationRange = destSheet.getRange(tableStartRow + 1, lcol1 + 3, timeRows.length, 1);
  utilizationRange.setFontWeight('bold');
  
  // Find last row
  var lastDataRow = totalRowIndex;
  
  // Calculate total available hours for utilization percentage
  var totalAvailableHours = 0;
  for (var i = 0; i < capacityData.length; i++) {
    totalAvailableHours += capacityData[i][0];
  }
  
  // Add summary section using batch operations
  var summaryStartRow = lastDataRow + 3;
  
  var summaryData = [
    ['Total Days', grandTotal],
    ['Total Hours', grandTotal],
    ['Table Sum Status', 'OK'],
    ['Booked Hours', Math.abs(grandTotal - e2Total) < 0.01 ? 'OK' : 'Check Sum']
  ];
  
  // Add overall utilization if there are available hours
  if (totalAvailableHours > 0) {
    var overallUtilization = grandTotal / totalAvailableHours;
    summaryData.push(['Overall Utilization', overallUtilization]);
  }
  
  destSheet.getRange(summaryStartRow, 1, summaryData.length, 2).setValues(summaryData);
  
  // Format summary status cells
  destSheet.getRange(summaryStartRow + 2, 2).setFontWeight('bold').setFontColor('#328332');
  
  if (Math.abs(grandTotal - e2Total) < 0.01) {
    destSheet.getRange(summaryStartRow + 3, 2).setFontWeight('bold').setFontColor('#328332');
  } else {
    destSheet.getRange(summaryStartRow + 3, 2).setFontWeight('bold').setFontColor('#F32C1E');
  }
  
  // Format overall utilization percentage if added
  if (totalAvailableHours > 0) {
    destSheet.getRange(summaryStartRow + 4, 2).setNumberFormat('0.0%').setFontWeight('bold');
  }
  
  console.log("Completed createHoursDaysTable function");
  debugLog("Grand total: " + grandTotal + ", Source total: " + e2Total);
  
  // Auto-resize the new columns to fit headers
  destSheet.autoResizeColumn(lcol1 + 2); // Total Available Hours
  destSheet.autoResizeColumn(lcol1 + 3); // % Utilization
  
  ui.alert('Success', 'Hours-Days table created successfully in tab: ' + destTabName, ui.ButtonSet.OK);
}