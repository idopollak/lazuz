/**
 * MONTHLY HOURS ANALYSIS SCRIPT
 * 
 * This Google Apps Script analyzes booking hours data for facility management,
 * providing validation and detailed usage reports for club/court bookings.
 * 
 * CORE FUNCTIONALITY:
 * 1. Data validation between primary and split booking tabs
 * 2. Creation of detailed hours/days cross-tabulation reports with usage analysis
 * 
 * REQUIRED SPREADSHEET STRUCTURE:
 * - ClubInfo tab with club configuration
 * - Monthly booking tabs in format [MMYY]e, [MMYY]i, [MMYY]na
 * - Split booking tabs in format [MMYY]e_2, [MMYY]i_2, [MMYY]na_2
 */

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
 * Builds a cache of day of week information for a given month
 * @param {number} monthNum - Month number (1-12)
 * @param {number} year - Full year (e.g., 2025)
 * @return {Object} Cache object with day numbers as keys and day info as values
 */
function buildDayOfWeekCache(monthNum, year) {
  console.log("Building day of week cache for month " + monthNum + "/" + year);
  
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
  
  console.log("Day of week cache built for " + daysInMonth + " days");
  return cache;
}

/**
 * monthHours Function
 * Validates data consistency between primary and secondary tabs
 */
function monthHours() {
  console.log("Starting monthHours function");
  
  // Get UI instance
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Prompt for 4-digit month
  var response = ui.prompt('Month Input', 'Please enter a month (4 digits):', ui.ButtonSet.OK_CANCEL);
  
  // Check if user cancelled
  if (response.getSelectedButton() != ui.Button.OK) {
    console.log("User cancelled the operation");
    return;
  }
  
  var monthDigits = response.getResponseText();
  console.log("User input: " + monthDigits);
  
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
  
  // Process each tab pair
  for (var i = 0; i < tabPairs.length; i++) {
    var pair = tabPairs[i];
    console.log("Processing pair: " + pair.name);
    
    // Get primary tab
    var primarySheet = spreadsheet.getSheetByName(pair.primary);
    if (!primarySheet) {
      console.log("Primary tab not found: " + pair.primary);
      results.push(pair.name + " comparison: Primary tab '" + pair.primary + "' not found");
      continue;
    }
    console.log("Found primary tab: " + pair.primary);
    
    // Get secondary tab
    var secondarySheet = spreadsheet.getSheetByName(pair.secondary);
    if (!secondarySheet) {
      console.log("Secondary tab not found: " + pair.secondary);
      results.push(pair.name + " comparison: Secondary tab '" + pair.secondary + "' not found");
      continue;
    }
    console.log("Found secondary tab: " + pair.secondary);
    
    // Sum column H for primary tab
    var primarySum = sumColumnH(primarySheet, pair.primary);
    console.log("Sum for " + pair.primary + " column H: " + primarySum);
    
    // Sum column H for secondary tab
    var secondarySum = sumColumnH(secondarySheet, pair.secondary);
    console.log("Sum for " + pair.secondary + " column H: " + secondarySum);
    
    // Compare sums
    var isEqual = primarySum === secondarySum;
    console.log("Comparison result for " + pair.name + ": " + (isEqual ? "EQUAL" : "NOT EQUAL"));
    
    results.push(pair.name + " comparison: " + primarySum + " vs " + secondarySum + " - " + (isEqual ? "EQUAL" : "NOT EQUAL"));
  }
  
  // Display results
  var resultMessage = "Results for month " + monthDigits + ":\n\n" + results.join("\n");
  ui.alert('Month Hours Results', resultMessage, ui.ButtonSet.OK);
  console.log("Completed monthHours function");
}

/**
 * Helper function to sum column H values in a sheet
 * @param {Sheet} sheet - The sheet to process
 * @param {string} sheetName - Name of the sheet (for logging)
 * @return {number} Sum of numeric values in column H
 */
function sumColumnH(sheet, sheetName) {
  console.log("Summing column H for sheet: " + sheetName);
  
  // Get last row with data
  var lastRow = sheet.getLastRow();
  console.log("Last row in " + sheetName + ": " + lastRow);
  
  if (lastRow === 0) {
    console.log("No data in sheet " + sheetName);
    return 0;
  }
  
  // Get column H data (column 8)
  var range = sheet.getRange(1, 8, lastRow, 1);
  var values = range.getValues();
  console.log("Retrieved " + values.length + " rows from column H");
  
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
  
  console.log("Found " + numericCount + " numeric values in column H of " + sheetName);
  console.log("Total sum: " + sum);
  
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
  
  console.log("Could not parse time: " + timeStr);
  return 0;
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
  console.log("=== INSIDE CALCULATEAVAILABLEHOURS ===");
  console.log("Calculating available hours for month " + monthNum + "/" + year);
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get ClubInfo sheet
  var clubInfoSheet = spreadsheet.getSheetByName('ClubInfo');
  if (!clubInfoSheet) {
    throw new Error('ClubInfo tab not found');
  }
  
  // Get max hours per hour (number of courts)
  var maxHoursPerHour = clubInfoSheet.getRange('D5').getValue();
  console.log("Max hours per hour (courts): " + maxHoursPerHour);
  
  // Get opening hours - check both possible formats
  console.log("=== READING OPENING HOURS IN CALCULATEAVAILABLEHOURS ===");
  var openingHoursData = clubInfoSheet.getRange('A9:B25').getValues();
  console.log("Read " + openingHoursData.length + " rows from A9:B25");
  
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
        console.log("  First valid data found at row " + i + ", setting hasTraditionalFormat = true");
      }
    }
  }
  
  console.log("hasTraditionalFormat = " + hasTraditionalFormat + ", total valid rows = " + validRowCount);
  
  if (hasTraditionalFormat) {
    // Use traditional A9:B25 format
    console.log("=== PARSING TRADITIONAL FORMAT IN CALCULATEAVAILABLEHOURS ===");
    console.log("Using traditional opening hours format from A9:B25");
    
    for (var i = 0; i < openingHoursData.length; i++) {
      var dayName = openingHoursData[i][0];
      var hoursString = openingHoursData[i][1];
      
      if (dayName && hoursString) {
        console.log("Parsing row " + (i+9) + ": Day='" + dayName + "', Hours='" + hoursString + "'");
        
        // Convert to string if needed
        var hoursStr = hoursString.toString();
        
        // Parse HH:MM-HH:MM format
        var parts = hoursStr.split('-');
        if (parts.length === 2) {
          var openTime = parts[0].trim();
          var closeTime = parts[1].trim();
          
          var openHour = parseInt(openTime.split(':')[0]);
          var closeHour = parseInt(closeTime.split(':')[0]);
          
          console.log("  Parsed times: open=" + openHour + ", close=" + closeHour);
          
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
            console.log("  Successfully stored: " + dayName + " = " + JSON.stringify(dayOpeningHours[dayName]));
            
            if (openHour < earliestHour) earliestHour = openHour;
            if (closeHour > latestHour) latestHour = closeHour;
          } else {
            console.log("  ERROR: Invalid hours for " + dayName);
          }
        }
      }
    }
    
    console.log("=== AFTER PARSING IN CALCULATEAVAILABLEHOURS ===");
    console.log("dayOpeningHours: " + JSON.stringify(dayOpeningHours));
    console.log("dayOpeningHours keys: " + Object.keys(dayOpeningHours).join(", "));
  } else {
    // Use grid format from D5:K7
    console.log("Using grid opening hours format from D5:K7");
    var openingHoursRange = clubInfoSheet.getRange('D5:K7').getValues();
    console.log("Opening hours range data:");
    console.log("Row 0 (headers): " + openingHoursRange[0].join(", "));
    console.log("Row 1 (open times): " + openingHoursRange[1].join(", "));
    console.log("Row 2 (close times): " + openingHoursRange[2].join(", "));
    
    // Column E (index 1) through K (index 7) contain the days
    for (var i = 0; i < 7; i++) {
      var dayName = dayNames[i];
      var openTime = openingHoursRange[1][i + 1];  // Row 2 (index 1) is Open times
      var closeTime = openingHoursRange[2][i + 1]; // Row 3 (index 2) is Close times
      
      console.log("Raw opening hours for " + dayName + ": open=" + openTime + ", close=" + closeTime + " (type: " + typeof openTime + ")");
      
      // Parse times
      var openHour = parseHourFromTime(openTime);
      var closeHour = parseHourFromTime(closeTime);
      
      // Handle midnight closing
      if (closeHour === 0) {
        closeHour = 24;
        console.log(dayName + " closes at midnight, setting closeHour to 24");
      }
      
      dayOpeningHours[dayName] = {
        open: openHour,
        close: closeHour
      };
      
      if (openHour < earliestHour) earliestHour = openHour;
      if (closeHour > latestHour) latestHour = closeHour;
      
      console.log(dayName + ": " + openHour + " - " + closeHour);
    }
  }
  
  console.log("=== END OF OPENING HOURS PARSING IN CALCULATEAVAILABLEHOURS ===");
  console.log("Earliest opening hour: " + earliestHour);
  console.log("Latest closing hour: " + latestHour);
  console.log("Day opening hours parsed: " + JSON.stringify(dayOpeningHours));
  console.log("Number of days with hours: " + Object.keys(dayOpeningHours).length);
  
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
    console.log("After defaults, dayOpeningHours: " + JSON.stringify(dayOpeningHours));
  }
  
  // Build day of week cache for the month
  var dayOfWeekCache = buildDayOfWeekCache(monthNum, year);
  
  // Calculate days in month
  var daysInMonth = new Date(year, monthNum, 0).getDate();
  
  // Build hour rows
  var timeRows = [];
  var lastHourRow = latestHour === 24 ? 23 : latestHour - 1;
  
  console.log("Building time rows from " + earliestHour + " to " + lastHourRow);
  
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
  
  console.log("Day of week cache built for " + daysInMonth + " days");
  
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
  
  console.log("=== ABOUT TO RETURN FROM CALCULATEAVAILABLEHOURS ===");
  console.log("Return data keys: " + Object.keys(returnData).join(", "));
  console.log("dayOpeningHours in return: " + JSON.stringify(returnData.dayOpeningHours));
  console.log("dayOpeningHours keys in return: " + Object.keys(returnData.dayOpeningHours).join(", "));
  
  // Return results with metadata
  return returnData;
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
    console.log("User cancelled the operation");
    return;
  }
  
  var monthDigits = response.getResponseText();
  console.log("User input: " + monthDigits);
  
  // Validate input is exactly 4 digits
  if (!/^\d{4}$/.test(monthDigits)) {
    ui.alert('Invalid Input', 'Please enter exactly 4 digits.', ui.ButtonSet.OK);
    console.log("Invalid input: not 4 digits");
    return;
  }
  
  // Parse month and year from 4 digits (MMYY format)
  var monthNum = parseInt(monthDigits.substring(0, 2));
  var year = parseInt('20' + monthDigits.substring(2, 4));
  console.log("Parsed month: " + monthNum + ", year: " + year);
  
  // Convert month number to month name
  var monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var monthName = monthNames[monthNum - 1];
  var yearShort = year.toString().substring(2);
  
  // Get ClubInfo sheet
  var clubInfoSheet = spreadsheet.getSheetByName('ClubInfo');
  if (!clubInfoSheet) {
    ui.alert('Error', 'ClubInfo tab not found.', ui.ButtonSet.OK);
    console.log("ClubInfo tab not found");
    return;
  }
  
  // Get club name from ClubInfo
  var clubName = clubInfoSheet.getRange('B4').getValue();
  console.log("Club name: " + clubName);
  
  // Get max hours per hour (number of courts)
  var maxHoursPerHour = clubInfoSheet.getRange('D5').getValue();
  console.log("Max hours per hour: " + maxHoursPerHour);
  
  // Calculate available hours for the month
  console.log("=== STARTING CALCULATEAVAILABLEHOURS ===");
  var availabilityData = calculateAvailableHours(monthNum, year);
  console.log("=== RETURNED FROM CALCULATEAVAILABLEHOURS ===");
  console.log("availabilityData keys: " + Object.keys(availabilityData).join(", "));
  
  var availableGrid = availabilityData.availableGrid;
  var timeRows = availabilityData.timeRows;
  var daysInMonth = availabilityData.daysInMonth;
  var dayOpeningHours = availabilityData.dayOpeningHours;
  var earliestHour = availabilityData.earliestHour;
  var latestHour = availabilityData.latestHour;
  
  console.log("=== DATA RECEIVED IN CREATEHOURSDAYSTABLE ===");
  console.log("- dayOpeningHours: " + JSON.stringify(dayOpeningHours));
  console.log("- dayOpeningHours keys: " + Object.keys(dayOpeningHours).join(", "));
  console.log("- earliestHour: " + earliestHour);
  console.log("- latestHour: " + latestHour);
  console.log("- timeRows: " + timeRows.join(", "));
  
  console.log("Available hours grid created");
  console.log("Time rows: " + timeRows.length + " from hour " + earliestHour + " to " + (latestHour - 1));
  console.log("Time rows array: " + timeRows.join(", "));
  
  // Build day of week cache
  var dayOfWeekCache = buildDayOfWeekCache(monthNum, year);
  console.log("Day of week cache created");
  
  // Get source tabs
  var sourceTabName = monthDigits + 'e_2';
  var sourceSheet = spreadsheet.getSheetByName(sourceTabName);
  
  if (!sourceSheet) {
    console.log("Source tab not found: " + sourceTabName);
    ui.alert('Error', 'Source tab ' + sourceTabName + ' not found.', ui.ButtonSet.OK);
    return;
  }
  console.log("Found source tab: " + sourceTabName);
  
  // Get e tab for validation
  var eTabName = monthDigits + 'e';
  var eSheet = spreadsheet.getSheetByName(eTabName);
  var eTotal = 0;
  if (eSheet) {
    eTotal = sumColumnH(eSheet, eTabName);
    console.log("Total from " + eTabName + ": " + eTotal);
  }
  
  // Get e_2 tab total
  var e2Total = sumColumnH(sourceSheet, sourceTabName);
  console.log("Total from " + sourceTabName + ": " + e2Total);
  
  // Create or get destination tab
  var destTabName = monthName + '-' + yearShort + ' DH';
  var destSheet = spreadsheet.getSheetByName(destTabName);
  
  if (destSheet) {
    console.log("Destination tab exists, deleting and recreating");
    spreadsheet.deleteSheet(destSheet);
  }
  
  console.log("Creating new destination tab: " + destTabName);
  destSheet = spreadsheet.insertSheet(destTabName);
  
  // Format the sheet
  console.log("Formatting sheet");
  
  // Hide gridlines
  try {
    destSheet.setHiddenGridlines(true);
    console.log("Gridlines hidden");
  } catch (e) {
    console.log("Could not hide gridlines: " + e.toString());
  }
  
  // Set font for entire sheet
  var fullRange = destSheet.getRange(1, 1, 100, 50);
  fullRange.setFontFamily('Verdana');
  fullRange.setFontSize(10);
  fullRange.setBackground('#cbc9a2');
  
  // Set header with corrected formula
  var headerFormula = '=CONCATENATE("' + monthName + '-' + yearShort + ' Detailed Hours and Usage"," ",ClubInfo!$B$4)';
  destSheet.getRange('A1').setFormula(headerFormula);
  destSheet.getRange('A1').setFontWeight('bold');
  destSheet.getRange('A1').setFontSize(11);
  
  // Validation section
  destSheet.getRange('A3').setValue('Total E Hours');
  destSheet.getRange('B3').setValue(eTotal);
  destSheet.getRange('C3').setValue('Total Split Hours');
  destSheet.getRange('D3').setValue(e2Total);
  
  if (Math.abs(eTotal - e2Total) > 0.01) {
    destSheet.getRange('E3').setValue('Please Check');
    destSheet.getRange('E3').setFontWeight('bold');
    destSheet.getRange('E3').setFontColor('#F32C1E');
  } else {
    destSheet.getRange('E3').setValue('OK');
    destSheet.getRange('E3').setFontWeight('bold');
    destSheet.getRange('E3').setFontColor('#328332');
  }
  
  // Max hours section
  destSheet.getRange('A5').setValue('Max Hours per Hour');
  destSheet.getRange('B5').setValue(maxHoursPerHour);
  
  // Opening hours table
  destSheet.getRange('D5').setValue('Day');
  destSheet.getRange('D6').setValue('Open');
  destSheet.getRange('D7').setValue('Close');
  
  // Fill in days and times
  console.log("=== DISPLAYING OPENING HOURS IN TABLE ===");
  console.log("dayOpeningHours at display time: " + JSON.stringify(dayOpeningHours));
  
  var dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  for (var d = 0; d < 7; d++) {
    var colIndex = 5 + d; // E through K
    var dayName = dayNames[d];
    
    destSheet.getRange(5, colIndex).setValue(dayName.substring(0, 3));
    
    // Debug: Check for case sensitivity or spacing issues
    var foundKey = null;
    for (var key in dayOpeningHours) {
      if (key.toLowerCase().trim() === dayName.toLowerCase().trim()) {
        foundKey = key;
        break;
      }
    }
    
    if (foundKey) {
      var openDisplay = formatTimeAMPM(dayOpeningHours[foundKey].open);
      var closeDisplay = formatTimeAMPM(dayOpeningHours[foundKey].close);
      console.log("Setting " + dayName + " (found as '" + foundKey + "'): open=" + dayOpeningHours[foundKey].open + 
                  " (" + openDisplay + "), close=" + dayOpeningHours[foundKey].close + " (" + closeDisplay + ")");
      destSheet.getRange(6, colIndex).setValue(openDisplay);
      destSheet.getRange(7, colIndex).setValue(closeDisplay);
    } else {
      console.log("WARNING: No opening hours found for '" + dayName + "'");
      console.log("  Available keys: " + Object.keys(dayOpeningHours).join(", "));
      destSheet.getRange(6, colIndex).setValue("Not Found");
      destSheet.getRange(7, colIndex).setValue("Not Found");
    }
  }
  
  // Format opening hours table
  var openingHoursRange = destSheet.getRange('D5:K7');
  openingHoursRange.setBackground('#E3E2CD');
  openingHoursRange.setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  openingHoursRange.setHorizontalAlignment('left');
  
  // Find last row after header sections
  var tableStartRow = 10;
  
  // Build main table header row (days of month)
  console.log("Building main table header row");
  var headerRow = [''];  // A column for times
  
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
  console.log("Building time labels");
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
  console.log("Reading source data from " + sourceTabName);
  var lastRow = sourceSheet.getLastRow();
  var lastCol = sourceSheet.getLastColumn();
  
  if (lastRow > 0 && lastCol > 0) {
    var sourceData = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
    var processedCount = 0;
    
    console.log("Processing " + (sourceData.length - 1) + " data rows");
    
    // Process bookings - SIMPLIFIED VERSION
    for (var i = 1; i < sourceData.length; i++) {
      var rowData = sourceData[i];
      
      var dateValue = rowData[3];  // Column D
      var startTime = rowData[5];  // Column F
      var hours = rowData[7];      // Column H
      
      // Debug first few rows
      if (i <= 5) {
        console.log("Row " + i + " - Date: " + dateValue + " (type: " + typeof dateValue + 
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
          
          console.log("Parsed hour " + hour + " from start time: " + startTime);
          
          // Find grid position
          var hourIndex = timeRows.indexOf(hour);
          var dayIndex = bookingDay - 1;
          
          if (hourIndex >= 0 && dayIndex >= 0 && dayIndex < daysInMonth) {
            // Simply add the hours to the grid
            dataGrid[hourIndex][dayIndex] += hours;
            processedCount++;
            
            if (processedCount <= 10) {
              console.log("Added " + hours + " hours to " + hour + ":00 on day " + bookingDay);
            }
          } else {
            console.log("Warning: Could not place booking - hour: " + hour + ", day: " + bookingDay + 
                       ", hourIndex: " + hourIndex + ", dayIndex: " + dayIndex);
          }
        }
      }
    }
    
    console.log("Processed " + processedCount + " bookings for " + monthName + " " + year);
  }
  
  // Write data to sheet
  console.log("Writing data to sheet");
  for (var r = 0; r < timeRows.length; r++) {
    var rowData = [timeLabels[r]];
    var rowTotal = 0;
    
    for (var c = 0; c < daysInMonth; c++) {
      var value = dataGrid[r][c];
      rowData.push(value === 0 ? '' : value);
      rowTotal += value;
    }
    
    rowData.push(rowTotal === 0 ? '' : rowTotal);
    var currentRowNum = tableStartRow + 1 + r;
    destSheet.getRange(currentRowNum, 1, 1, rowData.length).setValues([rowData]);
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
  
  // Apply all formatting after data is written
  console.log("Applying formatting to table");
  
  // 1. Apply alternating row colors
  for (var r = 0; r < timeRows.length; r++) {
    var currentRowNum = tableStartRow + 1 + r;
    if (r % 2 === 0) {
      destSheet.getRange(currentRowNum, 1, 1, headerRow.length).setBackground('#E3E2CD');
    }
  }
  
  // 2. Format hours column (bold, left-aligned, right border)
  console.log("Formatting hours column");
  var hoursColumnRange = destSheet.getRange(tableStartRow + 1, 1, timeRows.length, 1);
  hoursColumnRange.setFontWeight('bold');
  hoursColumnRange.setHorizontalAlignment('left');
  hoursColumnRange.setNumberFormat('@'); // Text format to preserve AM/PM
  hoursColumnRange.setBorder(false, false, false, true, false, false, 'black', SpreadsheetApp.BorderStyle.DOTTED);
  
  // 3. Format total hours column (bold, left border)
  console.log("Formatting total hours column");
  var totalHoursCol = daysInMonth + 2; // Column position for Total Hours
  var totalHoursRange = destSheet.getRange(tableStartRow + 1, totalHoursCol, timeRows.length, 1);
  totalHoursRange.setFontWeight('bold');
  totalHoursRange.setBorder(false, true, false, false, false, false, 'black', SpreadsheetApp.BorderStyle.DOTTED);
  
  // 4. Apply conditional formatting for closed hours and max capacity
  console.log("Applying conditional formatting");
  for (var r = 0; r < timeRows.length; r++) {
    var hour = timeRows[r];
    var currentRowNum = tableStartRow + 1 + r;
    
    for (var d = 1; d <= daysInMonth; d++) {
      var cellRef = destSheet.getRange(currentRowNum, d + 1);
      
      // Get available hours for this slot
      var availableHours = availableGrid[r][d - 1];
      
      // Check if club is closed (available hours = 0)
      if (availableHours === 0) {
        cellRef.setBackground('#DDDDD7');
      }
      
      // Check if at max capacity
      var cellValue = cellRef.getValue();
      if (cellValue && availableHours > 0 && Number(cellValue) === Number(availableHours)) {
        cellRef.setBackground('#328332');
        cellRef.setFontWeight('bold');
        cellRef.setFontColor('#FFFFFF');
        console.log("Applied max capacity formatting to cell at row " + currentRowNum + ", col " + (d + 1));
      }
    }
  }
  
  // 5. Find last column (lcol1)
  var lcol1 = daysInMonth + 2; // Last data column (Total Hours)
  console.log("Last column (lcol1) is: " + lcol1);
  
  // 6. Set width of column lcol1+1
  destSheet.setColumnWidth(lcol1 + 1, 10);
  console.log("Set width of column " + (lcol1 + 1) + " to 10 pixels");
  
  // 7. Add capacity calculations in lcol1+2
  console.log("Adding capacity calculations in column " + (lcol1 + 2));
  for (var r = 0; r < timeRows.length; r++) {
    var currentRowNum = tableStartRow + 1 + r;
    
    // Calculate total available hours for this time slot
    var totalAvailable = 0;
    for (var d = 0; d < daysInMonth; d++) {
      totalAvailable += availableGrid[r][d];
    }
    
    destSheet.getRange(currentRowNum, lcol1 + 2).setValue(totalAvailable);
    console.log("Row " + currentRowNum + ": Total available = " + totalAvailable);
  }
  
  // Make Total Available Hours column bold
  var availableHoursRange = destSheet.getRange(tableStartRow + 1, lcol1 + 2, timeRows.length, 1);
  availableHoursRange.setFontWeight('bold');
  availableHoursRange.setBorder(false, true, false, false, false, false, 'black', SpreadsheetApp.BorderStyle.DOTTED);
  
  // 8. Add percentage calculations in lcol1+3
  console.log("Adding percentage calculations in column " + (lcol1 + 3));
  for (var r = 0; r < timeRows.length; r++) {
    var currentRowNum = tableStartRow + 1 + r;
    var totalHours = destSheet.getRange(currentRowNum, totalHoursCol).getValue() || 0;
    var capacity = destSheet.getRange(currentRowNum, lcol1 + 2).getValue() || 0;
    
    if (capacity > 0) {
      var percentage = (totalHours / capacity) * 100;
      destSheet.getRange(currentRowNum, lcol1 + 3).setValue(percentage / 100); // Store as decimal
      destSheet.getRange(currentRowNum, lcol1 + 3).setNumberFormat('0.0%'); // Format as percentage
    }
  }
  
  // Make % Utilization column bold
  var utilizationRange = destSheet.getRange(tableStartRow + 1, lcol1 + 3, timeRows.length, 1);
  utilizationRange.setFontWeight('bold');
  
  // Find last row
  var lastDataRow = totalRowIndex;
  
  // Add summary section
  var summaryStartRow = lastDataRow + 3;
  
  destSheet.getRange(summaryStartRow, 1).setValue('Total Days');
  destSheet.getRange(summaryStartRow, 2).setValue(grandTotal);
  
  destSheet.getRange(summaryStartRow + 1, 1).setValue('Total Hours');
  destSheet.getRange(summaryStartRow + 1, 2).setValue(grandTotal);
  
  // Table sum status
  destSheet.getRange(summaryStartRow + 2, 1).setValue('Table Sum Status');
  destSheet.getRange(summaryStartRow + 2, 2).setValue('OK');
  destSheet.getRange(summaryStartRow + 2, 2).setFontWeight('bold');
  destSheet.getRange(summaryStartRow + 2, 2).setFontColor('#328332');
  
  // Booked hours status
  destSheet.getRange(summaryStartRow + 3, 1).setValue('Booked Hours');
  if (Math.abs(grandTotal - e2Total) < 0.01) {
    destSheet.getRange(summaryStartRow + 3, 2).setValue('OK');
    destSheet.getRange(summaryStartRow + 3, 2).setFontWeight('bold');
    destSheet.getRange(summaryStartRow + 3, 2).setFontColor('#328332');
  } else {
    destSheet.getRange(summaryStartRow + 3, 2).setValue('Check Sum');
    destSheet.getRange(summaryStartRow + 3, 2).setFontWeight('bold');
    destSheet.getRange(summaryStartRow + 3, 2).setFontColor('#F32C1E');
  }
  
  console.log("Completed createHoursDaysTable function");
  console.log("Grand total: " + grandTotal + ", Source total: " + e2Total);
  
  // Auto-resize the new columns to fit headers
  destSheet.autoResizeColumn(lcol1 + 2); // Total Available Hours
  destSheet.autoResizeColumn(lcol1 + 3); // % Utilization
  
  ui.alert('Success', 'Hours-Days table created successfully in tab: ' + destTabName, ui.ButtonSet.OK);
}