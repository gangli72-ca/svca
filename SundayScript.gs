/**
 * Adds a custom menu when the spreadsheet is opened.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var editors = ss.getEditors().map(function(e) {
    return e.getEmail().toLowerCase();
  });
  var user = Session.getActiveUser().getEmail().toLowerCase();
  var isEditor = editors.indexOf(user) !== -1;
  
  var menu = ui.createMenu("Service Scheduler");
  
  if (isEditor) {
    // Full admin menu
    menu
      .addItem("Refresh Blackout Dates", "refreshBlackoutDates")
      .addItem("Lock Blackout Dates", "lockBlackoutDates")
      .addItem("Unlock Blackout Dates", "unlockBlackoutDates")
      .addSeparator()
      .addItem("Auto Populate Schedule", "autoPopulateSchedule")
      .addItem("Check Conflicts",     "highlightConflicts")
      .addItem("Clear Highlights",       "clearScheduleHighlights")
      .addItem("Copy to Schedule History","copyScheduleToHistory")
      .addSeparator()
      .addItem("Send Notification Emails", "sendEmails")
      .addToUi();
  } else {
    // Non-editors see an *empty* (or minimal) menu
    // You can choose one:
    
    // 1) A completely hidden menu (NO menu items):
    // (Do not add any items)
    
    // 2) Or a minimal menu with just one help item:
    // .addItem("About Scheduler", "showHelpMessage");
  }
  
  menu.addToUi();
}

/**
 * Calculates the date range for the next quarter based on a configurable start month.
 *
 * It reads the “Quarter Start Month” from Config!A2 (defaults to January if empty or invalid), 
 * determines which quarter today falls into relative to that start month, then computes the 
 * first day (startDate) and last day (endDate) of the *following* quarter.
 *
 * @return {{startDate: Date, endDate: Date}} 
 *   - startDate: JavaScript Date for the first day of next quarter  
 *   - endDate:   JavaScript Date for the last day of next quarter
 */
function getNextQuarterRange() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var today = new Date();
  var currentMonth = today.getMonth() + 1; // 1–12
  var year = today.getFullYear();
  
  // Read configured Q1 start month from Config!A2
  var raw = ss.getSheetByName("Config").getRange("A2").getValue();
  var q1Start = parseInt(raw, 10);
  if (isNaN(q1Start) || q1Start < 1 || q1Start > 12) {
    q1Start = 1;  // default to January
  }
  
  // Determine current quarter index (0–3) relative to q1Start
  var offset = (currentMonth - q1Start + 12) % 12;
  var currentQ = Math.floor(offset / 3);
  var nextQ    = (currentQ + 1) % 4;
  
  // Compute start month/year of next quarter
  var startMonthRaw = (q1Start - 1) + nextQ * 3;
  var startYear     = year + Math.floor(startMonthRaw / 12);
  var startMonth    = (startMonthRaw % 12) + 1;
  var startDate     = new Date(startYear, startMonth - 1, 1);
  
  // Compute end date of next quarter (last day of the third month)
  var endMonthRaw = startMonthRaw + 2;
  var endYear     = year + Math.floor(endMonthRaw / 12);
  var endMonth    = (endMonthRaw % 12) + 1;
  var endDate     = new Date(endYear, endMonth, 0);
  
  return { startDate: startDate, endDate: endDate };
}


/**
 * Returns an array of all Sundays (as Date objects) in the next quarter.
 * @return {Date[]} Array of Sunday Date objects.
 */
function getSundaysForNextQuarter() {
  var quarterRange = getNextQuarterRange();
  var startDate = quarterRange.startDate;
  var endDate = quarterRange.endDate;
  
  // Find the first Sunday on or after the start date.
  var firstSunday = new Date(startDate);
  while (firstSunday.getDay() !== 0) { // 0 means Sunday
    firstSunday.setDate(firstSunday.getDate() + 1);
  }
  
  var sundays = [];
  for (var d = new Date(firstSunday); d <= endDate; d.setDate(d.getDate() + 7)) {
    sundays.push(new Date(d));
  }
  return sundays;
}

/**
 * Refreshes the Blackout Dates sheet.
 * - Reads volunteer names from the "Roles" sheet.
 * - Populates the header row with Sunday dates for the next quarter.
 * - Fills the first column with volunteer names and inserts checkboxes for each Sunday.
 * - Applies background colors to header cells.
 */
function refreshBlackoutDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rolesSheet = ss.getSheetByName("Roles");
  var blackoutSheet = ss.getSheetByName("Blackout Dates");
  
  // Get volunteer names from the Roles sheet (assumes names are in column A starting at row 2).
  var lastRow = rolesSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No volunteer names found in the Roles sheet.");
    return;
  }
  var namesRange = rolesSheet.getRange(2, 1, lastRow - 1, 1);
  var namesData = namesRange.getValues();
  
  // Get all Sundays for the next quarter.
  var sundays = getSundaysForNextQuarter();
  
  // Thoroughly clear the Blackout Dates sheet, including old checkboxes/data validations
  var maxRows = blackoutSheet.getMaxRows();
  var maxCols = blackoutSheet.getMaxColumns();
  var fullRange = blackoutSheet.getRange(1, 1, maxRows, maxCols);
  fullRange.clearContent();         // remove all values
  fullRange.clearFormat();          // remove colors/borders/fonts
  fullRange.clearDataValidations(); // remove old checkbox rules
  
  // Set header row: first header is "Name"; subsequent headers are Sunday dates.
  var header = ["Name"];
  var dateFormat = "MM/dd/yyyy";
  sundays.forEach(function(date) {
    header.push(Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), dateFormat));
  });
  var headerRange = blackoutSheet.getRange(1, 1, 1, header.length);
  headerRange.setValues([header]);
  headerRange.setBackground("#CCCCCC"); // column header background color
  
  // Write volunteer names into column A starting at row 2.
  var nameRange = blackoutSheet.getRange(2, 1, namesData.length, 1);
  nameRange.setValues(namesData);
  nameRange.setBackground("#DDDDDD"); // row header background color
  
  // Fill remaining cells with checkboxes.
  var numRows = namesData.length;
  var numCols = header.length - 1;
  var dataRange = blackoutSheet.getRange(2, 2, numRows, numCols);
  dataRange.clearContent();
  
  // Set data validation for checkboxes.
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  dataRange.setDataValidation(rule);
  
  SpreadsheetApp.getUi().alert("Blackout Dates sheet refreshed successfully.");
}

/**
 * Locks the Blackout Dates sheet so that checkboxes (and all cells) become view-only.
 * Removes any previous protections on the sheet before applying a new one.
 */
function lockBlackoutDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var blackoutSheet = ss.getSheetByName("Blackout Dates");
  
  // Remove existing protections.
  var protections = blackoutSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  protections.forEach(function(protection) {
    protection.remove();
  });
  
  // Protect the entire sheet.
  var protection = blackoutSheet.protect().setDescription("Blackout Dates Locked");
  protection.setWarningOnly(false);
  
  // Allow only the effective user (admin) to edit.
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  
  // Remove any other editors.
  var editors = protection.getEditors();
  editors.forEach(function(editor) {
    if (editor.getEmail() !== me.getEmail()) {
      protection.removeEditor(editor);
    }
  });
  
  // Ensure domain users cannot edit.
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  
  SpreadsheetApp.getUi().alert("Blackout Dates sheet has been locked.");
}

function unlockBlackoutDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var blackoutSheet = ss.getSheetByName("Blackout Dates");
  
  // Get all sheet protections on the Blackout Dates sheet.
  var protections = blackoutSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  
  // Remove each protection.
  protections.forEach(function(protection) {
    protection.remove();
  });
  
  SpreadsheetApp.getUi().alert("Blackout Dates sheet unlocked successfully.");
}


/**
 * Automatically populates the Schedule sheet with service dates and volunteer assignments.
 * - Populates column A with all Sundays in the next quarter.
 * - Uses volunteer role information from the Roles sheet and blackout info to assign volunteers.
 * - Applies round-robin assignment for each role in a persistent fashion, ensuring that if a volunteer is skipped due to a blackout, the rotation continues from that point.
 * - Ensures no volunteer is assigned more than one role on the same day.
 * - Sets dropdowns in each position column based on qualified volunteers.
 * - Applies header background colors.
 */
function autoPopulateSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = ss.getSheetByName("Schedule");
  var rolesSheet = ss.getSheetByName("Roles");
  var blackoutSheet = ss.getSheetByName("Blackout Dates");
  var dateFormat = "MM/dd/yyyy";
  
  // Get all Sundays for the next quarter.
  var sundays = getSundaysForNextQuarter();
  
  // Thoroughly clear the Schedule sheet, including any old data validations
  // that might be lingering in columns to the right of the active area.
  var maxRows = scheduleSheet.getMaxRows();
  var maxCols = scheduleSheet.getMaxColumns();
  var fullRange = scheduleSheet.getRange(1, 1, maxRows, maxCols);
  
  fullRange.clearContent();        // remove all values
  fullRange.clearFormat();         // remove background colors, fonts, etc.
  fullRange.clearDataValidations(); // remove ALL dropdowns / validation rules 
  
  // Get role headers from the Roles sheet.
  // Assumes: Col A = Name, Col B..(second last) = roles, LAST column = Email.
  var lastCol = rolesSheet.getLastColumn();
  var rolesHeader = rolesSheet.getRange(1, 2, 1, lastCol - 2).getValues()[0];
  
  // Write header row in Schedule sheet: Column A: "Date", columns B onward: role names.
  var scheduleHeader = ["Date"].concat(rolesHeader);
  var headerRange = scheduleSheet.getRange(1, 1, 1, scheduleHeader.length);
  headerRange.setValues([scheduleHeader]);
  headerRange.setBackground("#CCCCCC");
  
  // Write Sunday dates into column A (starting at row 2).
  var dateValues = sundays.map(function(date) {
    return [date];
  });
  scheduleSheet.getRange(2, 1, sundays.length, 1).setValues(dateValues);
  scheduleSheet.getRange(2, 1, sundays.length, 1).setNumberFormat(dateFormat);
  scheduleSheet.getRange(2, 1, sundays.length, 1).setBackground("#DDDDDD");
  
  // Build mapping of roles to qualified volunteers from the Roles sheet.
  // Assumes data starts at row 2: column A is volunteer name; columns B onward are checkboxes.
  var rolesDataRange = rolesSheet.getRange(2, 1, rolesSheet.getLastRow() - 1, rolesSheet.getLastColumn());
  var rolesData = rolesDataRange.getValues();
  var roleVolunteers = {};
  rolesHeader.forEach(function(role) {
    roleVolunteers[role] = [];
  });
  rolesData.forEach(function(row) {
    var name = row[0];
    rolesHeader.forEach(function(role, i) {
      if (row[i + 1] === true) {
        roleVolunteers[role].push(name);
      }
    });
  });
  
  // For each role column in Schedule sheet, set a dropdown that lists all qualified volunteer names.
  rolesHeader.forEach(function(role, i) {
    var dvRule = SpreadsheetApp.newDataValidation().requireValueInList(roleVolunteers[role], true).build();
    // Column in Schedule sheet is i+2 (since column 1 is Date).
    scheduleSheet.getRange(2, i + 2, sundays.length, 1).setDataValidation(dvRule);
  });
  
  // Load blackout data.
  var blackoutData = blackoutSheet.getDataRange().getValues();
  var blackoutHeader = blackoutData[0];
  var blackoutDateMap = {};
  for (var j = 1; j < blackoutHeader.length; j++) {
    var d = blackoutHeader[j];
    if (d instanceof Date) {
      var formatted = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), dateFormat);
      blackoutDateMap[formatted] = j;
    } else {
      blackoutDateMap[d] = j;
    }
  }
  var volunteerRowMap = {};
  for (var i = 1; i < blackoutData.length; i++) {
    var volName = blackoutData[i][0];
    volunteerRowMap[volName] = i;
  }
  
  // Set up a persistent round-robin pointer for each role.
  var lastAssignedIndex = {};
  rolesHeader.forEach(function(role) {
    lastAssignedIndex[role] = -1;
  });
  
  var floatingRoles = getFloatingRoles();

  // Map each volunteer to their spouse (if any), from Couples sheet.
  var couplesMap = getCouplesMap();

  // Track who served on the previous Sunday (any role).
  var servedLastSunday = {};

  // For each Sunday (each row in Schedule starting at row 2) and each role,
  // assign a volunteer using round-robin that respects:
  //   - blackout dates
  //   - one non-floating role per person per Sunday
  //   - couples cannot serve on the same Sunday
  //   - no one serves two consecutive Sundays
  for (var r = 0; r < sundays.length; r++) {
    var currentSunday = sundays[r];
    var currentSundayFormatted = Utilities.formatDate(currentSunday, ss.getSpreadsheetTimeZone(), dateFormat);
    
    // Track volunteers already assigned on THIS date (any role)
    var assignedForDate = [];
    
    rolesHeader.forEach(function(role, i) {
      var volunteers = roleVolunteers[role];
      var assigned = "";
      var isFloating = floatingRoles.indexOf(role) !== -1;
      
      if (volunteers.length > 0) {
        // Start from the volunteer following the last assigned one for this role.
        var startIndex = (lastAssignedIndex[role] + 1) % volunteers.length;
        var candidate = null;
        
        for (var k = 0; k < volunteers.length; k++) {
          var index = (startIndex + k) % volunteers.length;
          var volName = volunteers[index];
          
          // 1) Skip if volunteer is already assigned a non-floating role on this date.
          if (!isFloating && assignedForDate.indexOf(volName) !== -1) {
            continue;
          }
          
          // 2) Skip if this volunteer served last Sunday (no back-to-back Sundays).
          if (servedLastSunday[volName]) {
            continue;
          }
          
          // 3) Skip if volunteer's spouse is already serving on this date.
          var spouse = couplesMap[volName];
          if (spouse && assignedForDate.indexOf(spouse) !== -1) {
            continue;
          }
          
          // 4) Check if volunteer has a blackout on this date.
          var isBlackout = false;
          if (volunteerRowMap.hasOwnProperty(volName) && blackoutDateMap.hasOwnProperty(currentSundayFormatted)) {
            var bdValue = blackoutData[volunteerRowMap[volName]][blackoutDateMap[currentSundayFormatted]];
            if (bdValue === true) {
              isBlackout = true;
            }
          }
          
          // 5) If passes all checks, pick this volunteer.
          if (!isBlackout) {
            candidate = volName;
            lastAssignedIndex[role] = index; // update the pointer for this role
            break;
          }
        }
        
        if (candidate) {
          assigned = candidate;
          assignedForDate.push(candidate);
        }
      }
      
      // Write assignment (or blank if no valid candidate)
      scheduleSheet.getRange(r + 2, i + 2).setValue(assigned);
    });
    
    // After finishing this Sunday, update servedLastSunday for the next iteration.
    servedLastSunday = {};
    assignedForDate.forEach(function(name) {
      servedLastSunday[name] = true;
    });
  }
  
  SpreadsheetApp.getUi().alert("Schedule auto-populated successfully.");
}

/**
 * Inserts a log entry into the "Logs" sheet.
 * The new log entry is inserted as the second row (right after the header).
 * Column A is the current timestamp in "yyyy/MM/dd HH:mm" format (Pacific Timezone),
 * Column B is the provided description,
 * and Column C is set to the logged in user's name extracted from their email.
 *
 * @param {string} description - The description text for the log entry.
 */
function logAction(description) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Logs");
  
  if (!logSheet) {
    // Create the Logs sheet if it doesn't exist and set the header row.
    logSheet = ss.insertSheet("Logs");
    logSheet.getRange(1, 1, 1, 3).setValues([["Date", "Description", "Person"]]);
  }
  
  // Insert a new row right before the first date row (row 2).
  logSheet.insertRowBefore(2);
  
  // Get the current time and format it in the Pacific Timezone.
  var now = new Date();
  var pacificTimeZone = "America/Los_Angeles";
  var formattedTimestamp = Utilities.formatDate(now, pacificTimeZone, "yyyy/MM/dd HH:mm");
  
  // Get the logged in user's email and extract the name before the "@".
  var email = Session.getActiveUser().getEmail();
  var name = email.split('@')[0];
  
  // Populate the new row: Column A = timestamp, Column B = description, Column C = name.
  logSheet.getRange(2, 1).setValue(formattedTimestamp);
  logSheet.getRange(2, 2).setValue(description);
  logSheet.getRange(2, 3).setValue(name);
}

/**
 * An installable onEdit trigger for logging changes in the Schedule sheet.
 * When a manual edit occurs, logs a message like:
 * "Position [role] is changed from [oldValue] to [newValue] for the date [formattedDate]".
 *
 * Only manual changes trigger this event. Programmatic changes (like autoPopulateSchedule) are ignored.
 */
function handleScheduleEdit(e) {
  // Ensure the event object is present.
  if (!e) return;
  
  var sheet = e.range.getSheet();
  
  // Only proceed if the edited sheet is "Schedule".
  if (sheet.getName() !== "Schedule") return;
  
  // Ignore edits in the header row or the first column (date column).
  if (e.range.getRow() < 2 || e.range.getColumn() < 2) return;
  
  // Get the role name from the header (row 1) at the edited column.
  var role = sheet.getRange(1, e.range.getColumn()).getValue();
  
  // Retrieve the date from column A in the same row.
  var dateCell = sheet.getRange(e.range.getRow(), 1).getValue();
  if (!(dateCell instanceof Date)) return; // if no valid date, skip.
  
  // Format the date in Pacific Time (yyyy/MM/dd).
  var formattedDate = Utilities.formatDate(new Date(dateCell), "America/Los_Angeles", "yyyy/MM/dd");
  
  // Retrieve the old and new values. (e.oldValue is only available with an installable trigger.)
  var oldValue = e.oldValue || "";
  var newValue = e.value || "";
  
  // If there is no change, exit.
  if(oldValue === newValue) return;
  
  // Build the log message.
  var description = "**" + role + "** is changed from (" + oldValue + ") to (" + newValue + ") for " + formattedDate;
  
  // Log the action using the logAction() function.
  logAction(description);
}

/**
 * Installable onEdit trigger to enforce:
 * A user may only edit the blackout row that corresponds to their own email,
 * UNLESS they are an admin (email listed in Config!C2), in which case they
 * may edit any row.
 */
function handleBlackoutEdit(e) {
  if (!e) return;
  
  var range = e.range;
  var sheet = range.getSheet();
  
  // Only enforce on the "Blackout Dates" sheet
  if (sheet.getName() !== "Blackout Dates") return;
  
  var row = range.getRow();
  var col = range.getColumn();
  
  // Ignore header row and name column
  if (row < 2 || col < 2) return;
  
  // Get the current user email
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    // If we can't see the user email (e.g. some account types), safest is to block edits.
    if (typeof e.oldValue !== "undefined") {
      range.setValue(e.oldValue);
    } else {
      range.clearContent();
    }
    SpreadsheetApp.getActive().toast(
      "Edit not allowed: unable to verify your account email.",
      "Blackout Dates",
      5
    );
    return;
  }
  userEmail = userEmail.toLowerCase().trim();
  
  // --- Admin bypass: admins can edit any row ---
  var adminEmails = getAdminEmails();  // from Config!C2
  if (adminEmails.indexOf(userEmail) !== -1) {
    // Admin – allow the edit with no further checks
    return;
  }
  // ------------------------------------------------
  
  // The volunteer name for this row (col A)
  var volunteerName = sheet.getRange(row, 1).getDisplayValue().trim();
  if (!volunteerName) return;  // no name => nothing to enforce
  
  // Map name -> email from Roles sheet
  var emailMap      = getVolunteerEmailMap();
  var expectedEmail = emailMap[volunteerName];
  
  if (!expectedEmail) {
    // No email configured for this name; block for regular users
    if (typeof e.oldValue !== "undefined") {
      range.setValue(e.oldValue);
    } else {
      range.clearContent();
    }
    SpreadsheetApp.getActive().toast(
      "Edit not allowed: no email configured for \"" + volunteerName + "\" in Roles.",
      "Blackout Dates",
      5
    );
    return;
  }
  
  expectedEmail = expectedEmail.toLowerCase().trim();
  
  // If the logged-in email doesn't match the row's email, revert
  if (expectedEmail !== userEmail) {
    if (typeof e.oldValue !== "undefined") {
      range.setValue(e.oldValue);
    } else {
      range.clearContent();
    }
    SpreadsheetApp.getActive().toast(
      "You can only edit your own blackout dates row.",
      "Blackout Dates",
      5
    );
    return;
  }
  
  // If we reach here, the user email matches the row's email → edit allowed
}

function adhocTest() {
    MailApp.sendEmail({
      to:       "gang.li@svca.cc",
      subject:  "Hello World!",
      body:     "Hello World!"
    });
  //logAction("Hello World!");
}

/**
 * Copies the current Schedule into Schedule History,
 * overwriting any existing rows for the same quarter.
 */
function copyScheduleToHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var scheduleSheet = ss.getSheetByName("Schedule");
  var historySheet = ss.getSheetByName("Schedule History");
  var scheduleData = scheduleSheet.getDataRange().getValues();
  
  // 1) Create history sheet if needed, and set header row
  if (!historySheet) {
    historySheet = ss.insertSheet("Schedule History");
    historySheet
      .getRange(1, 1, 1, scheduleData[0].length)
      .setValues([ scheduleData[0] ]);
  }
  
  // 2) Determine the next-quarter range
  var qr = getNextQuarterRange();
  var startDate = qr.startDate;
  var endDate   = qr.endDate;
  
  // 3) Remove any existing history rows for that quarter
  var historyData = historySheet.getDataRange().getValues();
  for (var i = historyData.length-1; i >= 0; i--) {
    var rowDate = historyData[i][0];
    if (rowDate instanceof Date &&
        rowDate >= startDate &&
        rowDate <= endDate) {
      historySheet.deleteRow(i + 1);
    }
  }
  // Remove the left over header row
  historyData = historySheet.getDataRange().getValues();
  if(historyData.length > 0) {
    if (historyData[historyData.length-1][0] === 'Date')
      historySheet.deleteRow(historyData.length);
  }
  
  // 4) Append all Schedule rows (skip header at index 0)
  for (var r = 0; r < scheduleData.length; r++) {
    historySheet.appendRow(scheduleData[r]);
  }
  
  // 5) Notify
  SpreadsheetApp.getUi().alert(
    'Schedule History updated for ' +
    Utilities.formatDate(startDate, tz, 'yyyy/MM/dd') +
    ' – ' +
    Utilities.formatDate(endDate,   tz, 'yyyy/MM/dd')
  );
}

/**
 * Reads the “floating roles” from Config!B2 (comma-separated),
 * trims each entry, and returns a clean array.
 *
 * @return {string[]} Array of floating role names.
 */
function getFloatingRoles() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = ss.getSheetByName('Config');
  var raw    = config.getRange('B2').getDisplayValue();  // e.g. "Role 1,Role 2,Role 3"
  
  if (!raw) return [];
  
  return raw
    .split(',')                         // [ "Role 1", "Role 2", "Role 3" ]
    .map(function(item) {               // trim whitespace
      return item.trim();
    })
    .filter(function(item) {            // drop any empty strings
      return item.length > 0;
    });
}

/**
 * Builds a map from volunteer name -> email using the "Roles" sheet.
 * Assumes:
 *   - Column A: Name
 *   - Last column: Email
 *
 * @return {Object<string,string>} e.g. { "Alice": "alice@example.com", ... }
 */
function getVolunteerEmailMap() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var rolesSheet = ss.getSheetByName("Roles");
  if (!rolesSheet) return {};
  
  var lastRow = rolesSheet.getLastRow();
  var lastCol = rolesSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return {};
  
  // Grab all data including the last column (email)
  var range = rolesSheet.getRange(2, 1, lastRow - 1, lastCol);
  var data  = range.getValues();
  
  var map = {};
  data.forEach(function(row) {
    var name  = (row[0] || "").toString().trim();            // Column A
    var email = (row[lastCol - 1] || "").toString().trim();  // Last column = email
    if (name && email) {
      map[name] = email.toLowerCase();
    }
  });
  
  return map;
}

/**
 * Reads admin emails from Config!C2 (comma-separated) and returns
 * a normalized lowercase array.
 *
 * Example: "admin1@x.com, admin2@x.com"
 *   => ["admin1@x.com", "admin2@x.com"]
 *
 * @return {string[]} admin email list in lowercase.
 */
function getAdminEmails() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = ss.getSheetByName('Config');
  if (!config) return [];
  
  var raw = config.getRange('C2').getDisplayValue();
  if (!raw) return [];
  
  return raw
    .split(',')
    .map(function(item) {
      return item.trim().toLowerCase();
    })
    .filter(function(item) {
      return item.length > 0;
    });
}

/**
 * Highlights three types of conflicts on the Schedule sheet:
 *
 *  1) Same person assigned more than once on the same Sunday (same row, different roles)
 *     -> Light red (#FFCCCC)
 *
 *  2) Same person assigned on two consecutive Sundays (adjacent rows, any roles)
 *     -> Light yellow (#FFF2CC)
 *
 *  3) Husband and wife serving on the same Sunday (from Couples sheet)
 *     -> Light blue (#CCE5FF)
 *
 * Colors are layered with simple priority:
 *   - Same-day duplicate (red) is applied first
 *   - Consecutive-week conflict (yellow) can override red
 *   - Couple conflict (blue) can override both (highest priority)
 */
function highlightConflicts() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Schedule");
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return;
  
  // Data range: from row 2 (first Sunday) and column 2 (first role)
  var numRows = lastRow - 1;
  var numCols = lastCol - 1;
  var range   = sheet.getRange(2, 2, numRows, numCols);
  var values  = range.getValues();
  
  // Clear previous backgrounds
  range.setBackground(null);
  
  // Conflict flags: same-day duplicates, consecutive weeks, couples same day
  var sameDayDup    = [];
  var consecWeekDup = [];
  var coupleConflict = [];
  
  for (var r = 0; r < numRows; r++) {
    sameDayDup[r]     = [];
    consecWeekDup[r]  = [];
    coupleConflict[r] = [];
    for (var c = 0; c < numCols; c++) {
      sameDayDup[r][c]     = false;
      consecWeekDup[r][c]  = false;
      coupleConflict[r][c] = false;
    }
  }
  
  // --- 1) Same-day duplicates (existing behavior, but now via arrays) ---
  for (var r = 0; r < numRows; r++) {
    var counts = {};
    // Count occurrences per name in this row
    for (var c = 0; c < numCols; c++) {
      var name = values[r][c];
      if (name) {
        counts[name] = (counts[name] || 0) + 1;
      }
    }
    // Mark cells where the name appears more than once
    for (var c = 0; c < numCols; c++) {
      var name = values[r][c];
      if (name && counts[name] > 1) {
        sameDayDup[r][c] = true;
      }
    }
  }
  
  // --- 2) Duplicates on two consecutive Sundays (adjacent rows) ---
  // For each pair of consecutive rows r and r+1, if a name appears in both,
  // mark all occurrences of that name in both rows.
  for (var r = 0; r < numRows - 1; r++) {
    var rowNow  = values[r];
    var rowNext = values[r + 1];
    
    var namesNow  = {};
    var namesNext = {};
    
    // Collect names in current row
    for (var c = 0; c < numCols; c++) {
      var name = rowNow[c];
      if (name) namesNow[name] = true;
    }
    
    // Collect names in next row
    for (var c = 0; c < numCols; c++) {
      var name = rowNext[c];
      if (name) namesNext[name] = true;
    }
    
    // Intersection: names serving on consecutive Sundays
    for (var name in namesNow) {
      if (namesNext[name]) {
        // Mark all occurrences in row r
        for (var c = 0; c < numCols; c++) {
          if (values[r][c] === name) {
            consecWeekDup[r][c] = true;
          }
        }
        // Mark all occurrences in row r+1
        for (var c = 0; c < numCols; c++) {
          if (values[r + 1][c] === name) {
            consecWeekDup[r + 1][c] = true;
          }
        }
      }
    }
  }
  
  // --- 3) Husband & wife on the same Sunday ---
  // Use Couples sheet via getCouplesMap()
  var couplesMap = getCouplesMap();  // { "HusbandName": "WifeName", "WifeName": "HusbandName", ... }
  
  for (var r = 0; r < numRows; r++) {
    var rowValues = values[r];
    var rowNames = {};
    
    // Collect who is serving this Sunday
    for (var c = 0; c < numCols; c++) {
      var name = rowValues[c];
      if (name) {
        rowNames[name] = true;
      }
    }
    
    // For each cell, if this name has a spouse also in this row, mark as couple conflict
    for (var c = 0; c < numCols; c++) {
      var name = rowValues[c];
      if (!name) continue;
      
      var spouse = couplesMap[name];
      if (spouse && rowNames[spouse]) {
        coupleConflict[r][c] = true;
      }
    }
  }
    
  // --- Apply background colors based on conflicts ---
  // We'll build a 2D array of colors to set in one go.
  var colors = [];
  for (var r = 0; r < numRows; r++) {
    colors[r] = [];
    for (var c = 0; c < numCols; c++) {
      var color = null;
      
      if (sameDayDup[r][c]) {
        color = "#FFCCCC"; // light red: same-day multiple roles
      }
      if (consecWeekDup[r][c]) {
        color = "#FFF2CC"; // light yellow: consecutive Sundays
      }
      if (coupleConflict[r][c]) {
        color = "#CCE5FF"; // light blue: couple serving same day (highest priority)
      }
      
      colors[r][c] = color;
    }
  }
  
  range.setBackgrounds(colors);
}

/**
 * Clears all conflict highlight colors on the Schedule sheet
 * without touching the header row or the Date column.
 *
 * This is meant to be used after running highlightDuplicates()
 * and reviewing the conflicts.
 */
function clearScheduleHighlights() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Schedule");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Schedule sheet not found.");
    return;
  }
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  // No data rows or no role columns → nothing to clear
  if (lastRow < 2 || lastCol < 2) {
    SpreadsheetApp.getUi().alert("No schedule data to clear highlights from.");
    return;
  }
  
  // Only clear backgrounds in the data area:
  // rows 2..lastRow, columns 2..lastCol (roles only)
  var numRows = lastRow - 1;
  var numCols = lastCol - 1;
  var range   = sheet.getRange(2, 2, numRows, numCols);
  
  range.setBackground(null);
  
  SpreadsheetApp.getUi().alert("All conflict highlights have been cleared.");
}

/**
 * Builds a map of name -> spouse name from the "Couples" sheet.
 *
 * Couples sheet format:
 *   Row 1: Husband | Wife
 *   Rows 2+: pairs of names
 *
 * Example result:
 *   { "John": "Mary", "Mary": "John", ... }
 *
 * @return {Object<string,string>}
 */
function getCouplesMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Couples");
  if (!sheet) return {};
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  
  // Read columns A (Husband) and B (Wife) starting from row 2.
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var map = {};
  
  data.forEach(function(row) {
    var husband = (row[0] || "").toString().trim();
    var wife    = (row[1] || "").toString().trim();
    
    if (husband && wife) {
      map[husband] = wife;
      map[wife]    = husband;
    }
  });
  
  return map;
}

/**
 * Sends a sample blackout-dates notification email to the volunteer(s)
 * on the currently selected row(s) in the "Roles" sheet.
 *
 * Assumes:
 *   Roles!A = Name
 *   Roles!last column = Email
 *
 * The email includes a hyperlink that points directly to the
 * "Blackout Dates" sheet for this spreadsheet.
 *
 * Supports:
 *   - Single cell selection on a row
 *   - Selection of multiple rows
 *   - Multiple ranges (using Shift/Ctrl/Cmd selection)
 */
function sendEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rolesSheet = ss.getSheetByName("Roles");
  var blackoutSheet = ss.getSheetByName("Blackout Dates");
  
  if (!rolesSheet || !blackoutSheet) {
    SpreadsheetApp.getUi().alert('Missing "Roles" or "Blackout Dates" sheet.');
    return;
  }
  
  var activeSheet = ss.getActiveSheet();
  if (activeSheet.getName() !== "Roles") {
    SpreadsheetApp.getUi().alert('Please select one or more rows on the "Roles" sheet first.');
    return;
  }
  
  // Gather selected ranges (can be multiple ranges)
  var rangeList = ss.getActiveRangeList();
  if (!rangeList) {
    SpreadsheetApp.getUi().alert('Please select one or more volunteer rows on the "Roles" sheet.');
    return;
  }
  
  var lastCol = rolesSheet.getLastColumn();
  var ranges  = rangeList.getRanges();
  var recipients = [];  // {name, email, row}
  
  ranges.forEach(function(range) {
    var startRow = range.getRow();
    var endRow   = range.getLastRow();
    
    for (var r = startRow; r <= endRow; r++) {
      // Skip header row
      if (r < 2) continue;
      
      var name  = rolesSheet.getRange(r, 1).getDisplayValue().trim();        // Col A
      var email = rolesSheet.getRange(r, lastCol).getDisplayValue().trim();  // Last col = Email
      
      if (!name || !email) continue;
      
      recipients.push({
        name: name,
        email: email,
        row: r
      });
    }
  });
  
  if (recipients.length === 0) {
    SpreadsheetApp.getUi().alert('No valid name+email pairs found in the selected rows.');
    return;
  }
  
  // Build a URL that opens directly to the "Blackout Dates" sheet
  var baseUrl = ss.getUrl().split('#')[0];   // strip any existing gid
  var blackoutGid = blackoutSheet.getSheetId();
  var blackoutUrl = baseUrl + '#gid=' + blackoutGid;
  
  var sentCount = 0;
  var summaryLines = [];
  
  recipients.forEach(function(rec) {
    var name = rec.name;
    var email = rec.email;
    
    var subject = "SVCA Sunday School Blackout Dates";
    
    var plainBody =
      'Dear Co-Workers in Christ,\n\n' +

      'Praise the Lord! When you receive this email, it means that we are serving together in nurturing the next generation for the Lord and are committed to the SVCA Children’s Sunday School ministry.\n\n' +

      'To make the overall scheduling process smoother, the Children’s Ministry plans to prepare the Sunday serving rotations quarterly in advance (Dec–Feb, Mar–May, Jun–Aug, Sep–Nov). Each area coordinator will then make adjustments according to actual needs. To support this process, we have created a Blackout Dates form. Please mark the dates when you cannot serve, and the system will automatically generate a rotation schedule based on the rules, minimizing the chance of human errors in manual scheduling.\n\n' +

      'Please go to https://docs.google.com/spreadsheets/d/1UmGhZH8p5cqZSktto-i2qV5PuGH607UF8UTv_VGe9C8/edit#gid=1893596443 (you must log in with your SVCA email) and check the dates you **cannot serve** on the row corresponding to your name. Please be careful not to make changes on other co-workers’ rows.\n\n' +

      'This is the first time we are opening this scheduling process. We kindly ask everyone to complete it before November 23. We also warmly welcome any suggestions you may have—please write them in the Improvement Ideas sheet. We will do our best to continually improve this system.\n\n' +

      'May the Lord help us improve the quality of our service together and be good stewards of the time He gives us.\n\n' +

      'In Christ,\n\n' +
      'Sister Deborah\n' +
      'Children’s Sunday School Co-Worker';
    
    var htmlBody =
      '親愛的同工' + name + ':' + '<br><br>' +

      '感謝主，當您收到這張排班表，表示我們一起為主培育下一代，委身於SVCA兒童主日學事工。<br><br>' +

      '為了整體排班上更順暢，兒童部擬將主日服事輪值表以季度方式（12-2, 3-5, 6-8, 9-11月)預先總體安排，屆時由各項負責同工按照實際情況調動，特別設計了 Blackout Dates 表格。同工自已把<b>“無法上崗”</b>的日期圈選出來，其他時間開放讓系統自動按照規則輪值，盡量避免人工安排時的疏漏。<br><br>' +

      '請點擊 <a href="https://docs.google.com/spreadsheets/d/1UmGhZH8p5cqZSktto-i2qV5PuGH607UF8UTv_VGe9C8/edit#gid=1893596443">Blackout Dates 表格鏈接</a>（需要用您的svca email），在您名字對應的那一行勾選您<strong>無法上崗</strong>的日期。注意請不要在其他同工的行上勾選。<br><br>' +

      '這是第一次開放排班，敬請大家在 11/23 周日之前完成，所有改進意見也非常歡迎填寫在 Improvement Ideas 表格上，我們盡力將這個系統不斷完善。<br><br>' +

      '求主幫助我們一起提升服事的品質，做時間的好管家。<br><br>' +


      '雅慧姐妹<br>' +
      '兒童主日學同工<br><br><br>' +

      'Dear Co-Workers ' + name + ' in Christ,<br><br>' +

      'Praise the Lord! When you receive this email, it means that we are serving together in nurturing the next generation for the Lord and are committed to the SVCA Children’s Sunday School ministry.<br><br>' +

      'To make the overall scheduling process smoother, the Children’s Ministry plans to prepare the Sunday serving rotations quarterly in advance (Dec–Feb, Mar–May, Jun–Aug, Sep–Nov). Each area coordinator will then make adjustments according to actual needs. To support this process, we have created a Blackout Dates form. Please mark the dates when you <strong>cannot serve</strong>, and the system will automatically generate a rotation schedule based on the rules, minimizing the chance of human errors in manual scheduling.<br><br>' +

      'Please click the <a href="https://docs.google.com/spreadsheets/d/1UmGhZH8p5cqZSktto-i2qV5PuGH607UF8UTv_VGe9C8/edit#gid=1893596443">Blackout Dates link</a> (you must log in with your SVCA email) and check the dates you <strong>cannot serve</strong> on the row corresponding to your name. Please be careful not to make changes on other co-workers’ rows.<br><br>' +

      'This is the first time we are opening this scheduling process. We kindly ask everyone to complete it before November 23. We also warmly welcome any suggestions you may have—please write them in the Improvement Ideas sheet. We will do our best to continually improve this system.<br><br>' +

      'May the Lord help us improve the quality of our service together and be good stewards of the time He gives us.<br><br>' +

      'In Christ,<br><br>' +
      'Sister Deborah<br><br>' +
      'Children’s Sunday School Co-Worker<br>';
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    });
    
    // Log each email send
    var desc = "Sent blackout dates email to " + name + " (" + email + ")";
    if (typeof logAction === "function") {
      logAction(desc);
    }
        
    sentCount++;
    summaryLines.push(name + " (" + email + ")");
  });
  
  SpreadsheetApp.getUi().alert(
    "Sent " + sentCount + " test email(s) to:\n\n" + summaryLines.join("\n")
  );
}

