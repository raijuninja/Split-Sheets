/**
 * SPLIT SHEETS - Dynamic Bill Splitter for Google Sheets
 * https://github.com/YOUR_USERNAME/split-sheets
 * 
 * A flexible bill splitting script that works with any number of people.
 * 
 * SETUP:
 *   1. Create a Google Sheet with headers:
 *      A: Description | B: Who Paid | C: Amount | D: How to split | E+: [Person Names]
 *   2. Add person names in columns E, F, G, etc.
 *   3. The "Breakdown" column is auto-created after the last person.
 * 
 * SPLIT TYPES:
 *   - "Equally"   → Checkboxes - splits evenly among checked people
 *   - "Variably"  → Percentages - set one %, the other auto-calculates
 *   - "Fixed"     → Dollar amounts - set one $, the other auto-calculates
 * 
 * @license MIT
 */

var DESCRIPTION_COL = 1;
var WHO_PAID_COL = 2;
var AMOUNT_COL = 3;
var SPLIT_TYPE_COL = 4;
var FIRST_SPLITTER_COL = 5; // Column E is the first splitter
var BREAKDOWN_COL = 7; // Column G for breakdown (after splitter columns)

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  
  // If the "How to split" column (D) was edited, update the splitter columns
  if (col === SPLIT_TYPE_COL && row > 1) {
    updateSplitterCells(sheet, row, e.value);
  }
  
  // If a splitter column was edited, auto-calculate remaining percentage/amount
  if (col >= FIRST_SPLITTER_COL && row > 1) {
    autoCalculateRemaining(sheet, row, col, e.value);
  }
  
  calculate(sheet);
}

// Update splitter cells based on split type (checkboxes vs percentages)
function updateSplitterCells(sheet, row, splitType) {
  // Get header row to find splitter columns
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find all splitter columns (columns E onwards, but stop before Breakdown column)
  var splitterCols = [];
  for (var col = FIRST_SPLITTER_COL - 1; col < headerRow.length; col++) {
    var headerValue = headerRow[col] ? headerRow[col].toString().trim().toLowerCase() : '';
    
    // Stop if we hit the Breakdown column or any column that looks like a summary
    if (headerValue === 'breakdown' || headerValue === 'summary' || headerValue === 'total' || headerValue === '') {
      break;
    }
    
    if (headerRow[col] && headerRow[col].toString().trim() !== '') {
      splitterCols.push(col + 1); // Convert to 1-based
    }
  }
  
  if (splitterCols.length === 0) return;
  
  var splitTypeLower = splitType ? splitType.toString().toLowerCase().trim() : 'equally';
  
  for (var i = 0; i < splitterCols.length; i++) {
    var cell = sheet.getRange(row, splitterCols[i]);
    
    if (splitTypeLower === 'variably' || splitTypeLower === 'variable') {
      // Remove checkbox validation and set to percentage
      cell.removeCheckboxes();
      cell.setDataValidation(null);
      cell.setNumberFormat('0%'); // Percentage format
      
      // Set default percentage (equal split as starting point)
      var defaultPercent = splitterCols.length > 0 ? 1 / splitterCols.length : 0.5;
      cell.setValue(defaultPercent); // Store as decimal, display as %
    } else if (splitTypeLower === 'fixed') {
      // Remove checkbox validation and set to dollar amount
      cell.removeCheckboxes();
      cell.setDataValidation(null);
      cell.setNumberFormat('$#,##0.00'); // Currency format
      
      // Leave empty - user will enter fixed amount for one person
      cell.setValue('');
    } else {
      // Set as checkbox for "Equally"
      cell.removeCheckboxes();
      cell.setNumberFormat('@'); // Plain text for checkboxes
      cell.insertCheckboxes();
      cell.setValue(true); // Default to checked
    }
  }
}

// Auto-calculate remaining percentage or fixed amount when a splitter cell is edited
function autoCalculateRemaining(sheet, row, editedCol, newValue) {
  // Get split type for this row
  var splitType = sheet.getRange(row, SPLIT_TYPE_COL).getValue();
  var splitTypeLower = splitType ? splitType.toString().toLowerCase().trim() : 'equally';
  
  // Only auto-calculate for Variably or Fixed
  if (splitTypeLower !== 'variably' && splitTypeLower !== 'variable' && splitTypeLower !== 'fixed') {
    return;
  }
  
  // Get header row to find splitter columns
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find all splitter columns
  var splitterCols = [];
  for (var col = FIRST_SPLITTER_COL - 1; col < headerRow.length; col++) {
    var headerValue = headerRow[col] ? headerRow[col].toString().trim().toLowerCase() : '';
    if (headerValue === 'breakdown' || headerValue === 'summary' || headerValue === 'total' || headerValue === '') {
      break;
    }
    if (headerRow[col] && headerRow[col].toString().trim() !== '') {
      splitterCols.push(col + 1); // 1-based
    }
  }
  
  if (splitterCols.length < 2) return; // Need at least 2 splitters
  
  // Check if edited column is a splitter column
  var editedIndex = splitterCols.indexOf(editedCol);
  if (editedIndex === -1) return;
  
  if (splitTypeLower === 'variably' || splitTypeLower === 'variable') {
    // Auto-calculate remaining percentage
    var editedPercent = parsePercentage(newValue);
    
    // Calculate total of all OTHER cells (not the one just edited)
    var totalOtherPercent = 0;
    var emptyCells = [];
    for (var i = 0; i < splitterCols.length; i++) {
      if (splitterCols[i] !== editedCol) {
        var cellValue = sheet.getRange(row, splitterCols[i]).getValue();
        var pct = parsePercentage(cellValue);
        if (pct === 0 || cellValue === '' || cellValue === null) {
          emptyCells.push(splitterCols[i]);
        } else {
          totalOtherPercent += pct;
        }
      }
    }
    
    // If there's exactly one empty/zero cell, fill it with the remainder
    var remaining = 100 - editedPercent - totalOtherPercent;
    if (emptyCells.length === 1 && remaining >= 0) {
      var remainCell = sheet.getRange(row, emptyCells[0]);
      remainCell.setNumberFormat('0%');
      remainCell.setValue(remaining / 100); // Store as decimal
    } else if (splitterCols.length === 2) {
      // For 2 splitters, always auto-calc the other one
      var otherCol = splitterCols[0] === editedCol ? splitterCols[1] : splitterCols[0];
      var otherPercent = 100 - editedPercent;
      if (otherPercent >= 0) {
        var otherCell = sheet.getRange(row, otherCol);
        otherCell.setNumberFormat('0%');
        otherCell.setValue(otherPercent / 100); // Store as decimal
      }
    }
  } else if (splitTypeLower === 'fixed') {
    // Auto-calculate remaining fixed amount
    var amount = sheet.getRange(row, AMOUNT_COL).getValue();
    var totalAmount = parseAmount(amount);
    if (totalAmount <= 0) return;
    
    var editedAmount = parseAmount(newValue);
    
    // Calculate total of all OTHER cells
    var totalOtherAmount = 0;
    var emptyCells = [];
    for (var i = 0; i < splitterCols.length; i++) {
      if (splitterCols[i] !== editedCol) {
        var cellValue = sheet.getRange(row, splitterCols[i]).getValue();
        var amt = parseAmount(cellValue);
        if (amt === 0 || cellValue === '' || cellValue === null) {
          emptyCells.push(splitterCols[i]);
        } else {
          totalOtherAmount += amt;
        }
      }
    }
    
    // Fill the remaining amount in empty cells
    var remaining = totalAmount - editedAmount - totalOtherAmount;
    if (emptyCells.length === 1 && remaining >= 0) {
      var remainCell = sheet.getRange(row, emptyCells[0]);
      remainCell.setNumberFormat('$#,##0.00');
      remainCell.setValue(remaining);
    } else if (splitterCols.length === 2) {
      // For 2 splitters, always auto-calc the other one
      var otherCol = splitterCols[0] === editedCol ? splitterCols[1] : splitterCols[0];
      var otherAmount = totalAmount - editedAmount;
      if (otherAmount >= 0) {
        var otherCell = sheet.getRange(row, otherCol);
        otherCell.setNumberFormat('$#,##0.00');
        otherCell.setValue(otherAmount);
      }
    }
  }
}

function calculate(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSheet();
  
  // Get header row to find splitter names (starting from column E)
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find all splitter columns (columns E onwards with names in header)
  var splitters = [];
  for (var col = FIRST_SPLITTER_COL - 1; col < headerRow.length; col++) {
    if (headerRow[col] && headerRow[col].toString().trim() !== '' && col < BREAKDOWN_COL - 1) {
      splitters.push({
        name: headerRow[col].toString().trim(),
        colIndex: col, // 0-based index
        owes: 0,
        paid: 0
      });
    }
  }
  
  if (splitters.length === 0) {
    SpreadsheetApp.getUi().alert('No splitter names found in header row (columns E+)');
    return;
  }
  
  // Determine breakdown column (first column after last splitter)
  var breakdownCol = splitters[splitters.length - 1].colIndex + 2; // 1-based, after last splitter
  
  // Make sure header exists for breakdown column
  if (!sheet.getRange(1, breakdownCol).getValue()) {
    sheet.getRange(1, breakdownCol).setValue('Breakdown');
  }
  
  // Find all data rows (starting from row 2)
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) lastRow = 2;
  var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  var monthlyTotal = 0;
  var rowBreakdowns = []; // Store breakdown for each row
  var lastBillRow = 1; // Track the last row with a valid bill
  
  // Process ALL rows with data
  for (var i = 0; i < dataRange.length; i++) {
    var rowData = dataRange[i];
    var description = rowData[DESCRIPTION_COL - 1];
    var whoPaid = rowData[WHO_PAID_COL - 1];
    var amount = parseAmount(rowData[AMOUNT_COL - 1]);
    var splitType = rowData[SPLIT_TYPE_COL - 1];
    var rowNumber = i + 2; // Actual row number in sheet
    
    // Stop at empty rows (no description AND no amount)
    // A valid bill row must have a description and an amount
    var hasDescription = description && description.toString().trim() !== '';
    var hasAmount = amount && !isNaN(amount) && amount > 0;
    
    // Skip rows that look like previous summary rows (have "Due:" or "Summary")
    var descStr = description ? description.toString() : '';
    var payerStr = whoPaid ? whoPaid.toString() : '';
    if (descStr.includes('Due:') || payerStr.includes('Summary')) {
      continue; // Skip old summary rows, don't count them as data
    }
    
    // Skip rows without valid amounts (must have description AND amount to be a bill)
    if (!hasAmount) {
      continue;
    }
    
    // This is a valid bill row
    lastBillRow = rowNumber;
    
    monthlyTotal += amount;
    
    // Normalize split type (case-insensitive)
    var splitTypeLower = splitType ? splitType.toString().toLowerCase().trim() : 'equally';
    
    var breakdownParts = []; // Who pays what for this row
    
    // ============================================
    // EQUALLY: Split among checked people
    // ============================================
    if (splitTypeLower === 'equally') {
      // Find which splitters are checked for this row
      var checkedSplitters = [];
      for (var s = 0; s < splitters.length; s++) {
        var cellValue = rowData[splitters[s].colIndex];
        if (cellValue === true || cellValue === 'TRUE' || cellValue === 1) {
          checkedSplitters.push(splitters[s]);
        }
      }
      
      // If no one is checked, skip this row
      if (checkedSplitters.length === 0) {
        rowBreakdowns.push({ row: rowNumber, text: '' });
        continue;
      }
      
      // Calculate each person's equal share
      var sharePerPerson = amount / checkedSplitters.length;
      
      // Credit the payer and debit the splitters
      for (var s = 0; s < checkedSplitters.length; s++) {
        var splitter = checkedSplitters[s];
        
        // If this person paid, credit them the full amount
        if (splitter.name === whoPaid) {
          splitter.paid += amount;
        }
        
        // Everyone checked owes their share
        splitter.owes += sharePerPerson;
        
        // Add to breakdown (only show people who didn't pay)
        if (splitter.name !== whoPaid) {
          breakdownParts.push(splitter.name + ' Pays: ' + formatCurrency(sharePerPerson));
        }
      }
    }
    // ============================================
    // VARIABLY: Split by percentages
    // ============================================
    else if (splitTypeLower === 'variably' || splitTypeLower === 'variable') {
      // Read percentages from each splitter column
      var percentages = [];
      var totalPercent = 0;
      var hasCheckboxes = false;
      var checkedCount = 0;
      
      for (var s = 0; s < splitters.length; s++) {
        var cellValue = rowData[splitters[s].colIndex];
        
        // Check if this is a checkbox (TRUE/FALSE) instead of a percentage
        if (cellValue === true || cellValue === 'TRUE') {
          hasCheckboxes = true;
          checkedCount++;
        }
        
        var percent = parsePercentage(cellValue);
        percentages.push({
          splitter: splitters[s],
          percent: percent,
          isChecked: (cellValue === true || cellValue === 'TRUE' || cellValue === 1)
        });
        totalPercent += percent;
      }
      
      // If checkboxes are used instead of percentages, fall back to equal split
      if (hasCheckboxes && totalPercent === 0 && checkedCount > 0) {
        var equalPercent = 100 / checkedCount;
        for (var p = 0; p < percentages.length; p++) {
          if (percentages[p].isChecked) {
            percentages[p].percent = equalPercent;
            totalPercent += equalPercent;
          }
        }
      }
      
      // Skip if no percentages entered and no checkboxes
      if (totalPercent === 0) {
        rowBreakdowns.push({ row: rowNumber, text: '' });
        continue;
      }
      
      // Calculate each person's share based on percentage
      for (var p = 0; p < percentages.length; p++) {
        var pct = percentages[p];
        if (pct.percent > 0) {
          var share = amount * (pct.percent / 100);
          
          // If this person paid, credit them the full amount
          if (pct.splitter.name === whoPaid) {
            pct.splitter.paid += amount;
          }
          
          // They owe their percentage share
          pct.splitter.owes += share;
          
          // Add to breakdown (only show people who didn't pay)
          if (pct.splitter.name !== whoPaid) {
            breakdownParts.push(pct.splitter.name + ' Pays: ' + formatCurrency(share));
          }
        }
      }
    }
    // ============================================
    // FIXED: Split by fixed dollar amounts
    // ============================================
    else if (splitTypeLower === 'fixed') {
      // Read fixed amounts from each splitter column
      var fixedAmounts = [];
      var totalFixed = 0;
      
      for (var s = 0; s < splitters.length; s++) {
        var cellValue = rowData[splitters[s].colIndex];
        var fixedAmt = parseAmount(cellValue);
        fixedAmounts.push({
          splitter: splitters[s],
          amount: fixedAmt
        });
        totalFixed += fixedAmt;
      }
      
      // Skip if no amounts entered
      if (totalFixed === 0) {
        rowBreakdowns.push({ row: rowNumber, text: '' });
        continue;
      }
      
      // Calculate each person's share based on fixed amount
      for (var f = 0; f < fixedAmounts.length; f++) {
        var fix = fixedAmounts[f];
        if (fix.amount > 0) {
          // If this person paid, credit them the full amount
          if (fix.splitter.name === whoPaid) {
            fix.splitter.paid += amount;
          }
          
          // They owe their fixed amount
          fix.splitter.owes += fix.amount;
          
          // Add to breakdown (only show people who didn't pay)
          if (fix.splitter.name !== whoPaid) {
            breakdownParts.push(fix.splitter.name + ' Pays: ' + formatCurrency(fix.amount));
          }
        }
      }
    }
    
    // Store breakdown for this row
    rowBreakdowns.push({ row: rowNumber, text: breakdownParts.join(', ') });
  }
  
  // Write breakdowns to the breakdown column (with light blue background)
  var lightBlue = '#E3F2FD'; // Soothing cool blue
  for (var b = 0; b < rowBreakdowns.length; b++) {
    var breakdownCell = sheet.getRange(rowBreakdowns[b].row, breakdownCol);
    breakdownCell.setValue(rowBreakdowns[b].text);
    breakdownCell.setBackground(lightBlue);
  }
  
  // Calculate net balances (positive = owed money, negative = owes money)
  var balances = [];
  for (var s = 0; s < splitters.length; s++) {
    balances.push({
      name: splitters[s].name,
      balance: splitters[s].paid - splitters[s].owes // positive = others owe them
    });
  }
  
  // Summary row: always right after the last bill row
  var summaryRow = lastBillRow + 1;
  
  // Clear the summary row
  sheet.getRange(summaryRow, 1, 1, sheet.getLastColumn()).clearContent();
  
  // Write Monthly Sum label and total
  sheet.getRange(summaryRow, DESCRIPTION_COL).setValue('Due: 1st');
  
  // Clear data validation and set Summary in column B
  var summaryPayerCell = sheet.getRange(summaryRow, WHO_PAID_COL);
  summaryPayerCell.setDataValidation(null);
  summaryPayerCell.setValue('Summary');
  
  sheet.getRange(summaryRow, AMOUNT_COL).setValue(formatCurrency(monthlyTotal));
  
  // Clear splitter columns on summary row (remove any checkboxes)
  for (var s = 0; s < splitters.length; s++) {
    var cell = sheet.getRange(summaryRow, splitters[s].colIndex + 1);
    cell.removeCheckboxes();
    cell.setValue('');
  }
  
  // Build summary text for breakdown column
  var summaryParts = [];
  for (var s = 0; s < balances.length; s++) {
    var balance = balances[s].balance;
    if (balance < 0) {
      // They owe money
      summaryParts.push(balances[s].name + ' owes ' + formatCurrency(Math.abs(balance)));
    }
  }
  
  // Write summary to breakdown column
  sheet.getRange(summaryRow, breakdownCol).setValue(summaryParts.join(' | '));
  
  // Format summary row: bold and light blue background
  var summaryRange = sheet.getRange(summaryRow, 1, 1, sheet.getLastColumn());
  summaryRange.setFontWeight('bold');
  summaryRange.setBackground(lightBlue);
}

// Parse percentage from various formats: 50%, 0.5, 50, "50%"
function parsePercentage(value) {
  if (value === null || value === undefined || value === '' || value === false) return 0;
  if (value === true) return 0; // Checkbox checked but we need a number for variable
  
  var str = value.toString().trim();
  
  // Handle percentage sign: "50%" -> 50
  if (str.includes('%')) {
    return parseFloat(str.replace('%', '')) || 0;
  }
  
  var num = parseFloat(str);
  if (isNaN(num)) return 0;
  
  // If it's a decimal like 0.5, convert to 50%
  // Assumes values <= 1 are decimals (except 0 and 1 which could be either)
  if (num > 0 && num <= 1) {
    return num * 100;
  }
  
  // Otherwise treat as percentage already (e.g., 50 = 50%)
  return num;
}

// Parse amount from various formats ($1,234.56 or 1234.56)
function parseAmount(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  var str = value.toString().replace(/[$,]/g, '');
  return parseFloat(str) || 0;
}

// Format number as currency
function formatCurrency(amount) {
  return '$' + amount.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/**
 * Manual recalculate - can be run from Extensions > Macros
 * or added to a custom menu
 */
function recalculate() {
  calculate(SpreadsheetApp.getActiveSheet());
}

/**
 * Creates a custom menu when the spreadsheet opens
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 Split Sheets')
    .addItem('Recalculate', 'recalculate')
    .addToUi();
}
