{\rtf1\ansi\ansicpg1252\cocoartf1561\cocoasubrtf600
{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 // Martin Schulman\
// IEEE Northern VA Section\
// March 2019\
\
// Add menu items to spreadsheet\
function onOpen() \{\
  var ui = SpreadsheetApp.getUi();\
  ui.createMenu('Merge')\
  .addItem('Cover Letters', 'coverLetters')\
  .addItem('Certificates', 'certificates')\
  .addItem('Address Labels', 'addressLabels')\
  .addToUi();\
\};\
\
\
// Return the numeric index for the given column label\
function getIndex(rangeValues, lastColumn, colHeader) \{\
  var i;\
  for (i=0; i<lastColumn; i++) \{\
    if (rangeValues[0][i] === colHeader) \{\
      return(i);\
    \};\
  \};\
  return(-1);\
\};\
\
\
// Duplicate the template\
function createDuplicateDocument(sourceName, name) \{\
\
  // Find first file with template's name\
  var fileList = DriveApp.getFilesByName(sourceName);\
  if ( ! fileList.hasNext()) \{\
    SpreadsheetApp.getUi().alert("Unable to find template: " + sourceName);\
    return(-1);\
  \};\
  var source = fileList.next();\
\
  // Copy it to memory\
  var newFile = source.makeCopy(name);\
\
  // Get the directory of the template\
  var fileParents = source.getParents();\
  if ( ! fileParents.hasNext() ) \{\
    SpreadsheetApp.getUi().alert("Could not get " + sourceName + " parent directory.");\
    return(-1);\
  \};\
  \
  // Get parent directory  (.getName() returns printable string)\
  var targetFolder = fileParents.next();\
\
  // Save to directory\
  targetFolder.addFile(newFile);\
\
  // Return the id\
  return newFile.getId();\
\};\
\
\
// Generate cover letters\
function coverLetters() \{\
  var ss = SpreadsheetApp.getActiveSpreadsheet();\
  var sheet = ss.getActiveSheet();\
  var rangeData = sheet.getDataRange();\
  var lastColumn = rangeData.getLastColumn();\
  var lastRow = rangeData.getLastRow();\
  var searchRange = sheet.getRange(1,1, lastRow, lastColumn);\
\
  // Get array of values in the search Range\
  var rangeValues = searchRange.getValues();\
\
  // Set template to the index of the Template column\
  templateIndex=getIndex(rangeValues, lastColumn, "CoverTemplate");\
  \
  if ( templateIndex == -1 ) \{\
    SpreadsheetApp.getUi().alert("Missing CoverTemplate column; exiting.");\
    return(0);\
  \};\
  \
  // Loop over rows\
  for ( i = 1; i < lastRow; i++)\{\
\
    // Array begins at zero, but label files with spreadsheet row\
    var row = i + 1;\
\
    // Zero-padded sequence number\
    var seq="";\
    if (row < 10 ) \{\
      seq = "00";\
    \} else if (row < 100) \{\
      seq = "0";\
    \};\
  \
    // Get last name - SHOULD always be present!\
    var lastNameIndex = getIndex(rangeValues, lastColumn, "Last");\
\
    // Omit if not defined\
    if ( lastNameIndex == -1 ) \{\
      var lastName = "";\
    \} else \{\
      var lastName = "_" + rangeValues[i][lastNameIndex];\
    \};\
    \
    // Form document name using sequence number and last name\
    var newName = seq + row + lastName + "_" + rangeValues[i][templateIndex];\
\
    // Duplicate document and get the body\
    var docId = DocumentApp.openById(\
      createDuplicateDocument(rangeValues[i][templateIndex], newName)\
    );\
    \
    // Check for error\
    if ( docId == -1 ) \{\
      return(0);\
    \};\
\
    // Get the document body\
    var body = docId.getBody();\
    \
    // Loop over column headings and replace bracketed items\
    for ( j = 0; j < lastColumn; j++)\{\
      body.replaceText("<<"+rangeValues[0][j]+">>", rangeValues[i][j]);\
    \};\
  \}; \
  i--;\
  SpreadsheetApp.getUi().alert("Generated " + i + " cover letters");\
\};\
\
\
// Generate certificates\
function certificates() \{\
  var ss = SpreadsheetApp.getActiveSpreadsheet();\
  var sheet = ss.getActiveSheet();\
  var rangeData = sheet.getDataRange();\
  var lastColumn = rangeData.getLastColumn();\
  var lastRow = rangeData.getLastRow();\
  var searchRange = sheet.getRange(1,1, lastRow, lastColumn);\
\
  // Get array of values in the search Range\
  var rangeValues = searchRange.getValues();\
\
  // Prompt for certificate template file\
  var ui=SpreadsheetApp.getUi();\
  var result = ui.prompt(\
    "Select Certificate Template",\
    "Enter name of certificate template",\
    ui.ButtonSet.OK_CANCEL\
  );\
\
  // Quit if Canceled\
  if ( result.getSelectedButton() == ui.Button.CANCEL ) \{\
    return(0);\
  \};\
\
  // Certificate template file\
  certTemplate = result.getResponseText();\
  \
  // Loop over rows\
  for ( i = 1; i < lastRow; i++)\{\
\
    // Array begins at zero, but label files with spreadsheet row\
    var row = i + 1;\
\
    // Zero-padded sequence number\
    var seq = "";\
    if (row < 10 ) \{\
      seq = "00";\
    \} else if (row < 100) \{\
      seq = "0";\
    \};\
      \
    // Get last name\
    var last = getIndex(rangeValues, lastColumn, "Last");\
\
    // Duplicate the appropriate template\
    var newName = seq + row + "_" + rangeValues[i][last] + "_" + certTemplate;\
    \
    // Duplicate document and get the Id\
    var copyId = createDuplicateDocument(certTemplate, newName);\
    if ( copyId == -1 ) \{\
      return(0);\
    \};\
\
    // Build requests; use JSON.stringify(requests) for printable string\
    var requests = [];\
\
    for ( j = 0; j < lastColumn; j++ ) \{\
      var ct = \{\};\
      ct.text = '<<' + rangeValues[0][j] + '>>';\
      ct.matchCase = true;\
      var rat = \{\};\
      rat.containsText = ct;\
      rat.replaceText = rangeValues[i][j].toString();\
      var req = \{\};\
      req.replaceAllText = rat;\
      requests[j]=req;\
    \};\
      \
    // Apply changes to certificate\
    var result = Slides.Presentations.batchUpdate(\{\
      requests: requests\
    \}, copyId);\
  \}; \
  i--;\
  SpreadsheetApp.getUi().alert("Generated " + i + " certificates");\
\};\
  \
\
  // Generate address labels\
function addressLabels() \{\
  var ss = SpreadsheetApp.getActiveSpreadsheet();\
  var sheet = ss.getActiveSheet();\
  var rangeData = sheet.getDataRange();\
  var lastColumn = rangeData.getLastColumn();\
  var lastRow = rangeData.getLastRow();\
  var searchRange = sheet.getRange(1,1, lastRow, lastColumn);\
\
  var labelsPerPage = 6;\
  \
  // Prompt for address label template file\
  var ui=SpreadsheetApp.getUi();\
  var result = ui.prompt(\
    "Select Address Label Template",\
    "Enter name of address label template with " + labelsPerPage + " labels per page.",\
    ui.ButtonSet.OK_CANCEL\
  );\
\
  // Quit if Canceled\
  if ( result.getSelectedButton() == ui.Button.CANCEL ) \{\
    return(0);\
  \};\
\
  // Address label template file\
  labelTemplate = result.getResponseText();\
  \
  // Get array of values in the search Range\
  var rangeValues = searchRange.getValues();\
\
  // Set template to the index of the Template column\
  templateIndex=getIndex(rangeValues, lastColumn, "CoverTemplate");\
\
  \
  // Page counter\
  var pageCount = 0;\
  \
  // Declare label counter and set beyond range\
  var labelId = labelsPerPage + 1;\
\
  // Get indices for first and last names and addresses\
  var lastNameIndex = getIndex(rangeValues, lastColumn, "Last");\
  var firstNameIndex = getIndex(rangeValues, lastColumn, "First");\
  var address1Index = getIndex(rangeValues, lastColumn, "Address1");\
  var address2Index = getIndex(rangeValues, lastColumn, "Address2");\
  \
  // Loop over rows\
  for ( i = 1; i < lastRow; i++)\{\
\
    // Open new document if all labels used\
    if ( labelId > 6 ) \{\
\
      // Increment page counter\
      pageCount++;\
      \
      // Reset label counter\
      labelId = 1;\
      \
      // Duplicate label template\
      var docId = DocumentApp.openById(\
      createDuplicateDocument(labelTemplate, "AddressLabel_" + pageCount)\
      );\
    \};\
    \
    // Check for error\
    if ( docId == -1 ) \{\
      return(0);\
    \};\
\
    // Get the document body\
    var body = docId.getBody();\
\
    // Replace Name field    \
    body.replaceText("Label"+labelId+"Name", rangeValues[i][firstNameIndex] + " " +\
      rangeValues[i][lastNameIndex]);\
    \
    // Replace Address fields\
    body.replaceText("Label"+labelId+"Address1", rangeValues[i][address1Index]);\
    body.replaceText("Label"+labelId+"Address2", rangeValues[i][address2Index]);\
    \
    // Increment label counter\
    labelId++;\
  \}; \
  i--;\
  SpreadsheetApp.getUi().alert("Generated " + i + " address labels on " + pageCount + " pages.");\
\};\
}