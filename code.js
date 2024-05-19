function copySheetDataWithFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheetNames = ["Kitchen", "Living Room", "Balcony", "Cleaning"];
  var targetSheetName = "Summary";

  // Check if the target sheet exists; if not, create it
  var targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
  } else {
    targetSheet.clear();  // Clear the target sheet
  }

  // Add keys with NAMES in row 1
  var nameKeys = [
    { name: "Names", color: "#FFFFFF" },  // Default background color for header
    { name: "Lili", color: "#CFFFF1" },
    { name: "Nina", color: "#E3CFFF" },
    { name: "Vienna", color: "#CFF5FF" },
    { name: "Cece", color: "#FFCFFB" }
  ];

  for (var i = 0; i < nameKeys.length; i++) {
    var key = nameKeys[i];
    var column = i + 1; // Columns are 1-indexed in Google Sheets
    
    // Set key name and background color
    targetSheet.getRange(1, column).setValue(key.name).setBackground(key.color);
  }

  sourceSheetNames.forEach(function(sheetName, index) {
    var sourceSheet = ss.getSheetByName(sheetName);
    if (!sourceSheet) {
      Logger.log("Sheet not found: " + sheetName);
      return;
    }

    var dataRange = sourceSheet.getDataRange();
    Logger.log("Data range for " + sheetName + ": " + dataRange.getA1Notation());
    
    var values = dataRange.getValues();
    var backgrounds = dataRange.getBackgrounds();
    
    
    for (var row = 0; row < backgrounds.length; row++) {
      for (var col = 0; col < backgrounds[row].length; col++) {
       
      var nameColors = {
      "#ffcffb": "Cece",
      "#cffff1": "Lili",
      "#cff5ff": "Vienna",
      "#e3cfff": "Nina"
      };

      var value = values[row][col];
    var background = backgrounds[row][col];
      // Check if the background color matches any of the colors in nameColor
      for (var color in nameColors) {
        //Logger.log("this color." + color + '.');
        //Logger.log("this background." + background + ".");
        if (background === color) {
          Logger.log("Matched " + background);
          var nameSheet = ss.getSheetByName(nameColors[color]);
          var nameValues = nameSheet.getRange("A:A").getValues().flat();
          if (nameSheet) {
             // Normalize the new value (trim leading/trailing spaces, convert to lowercase)
          newValue = value.trim().toLowerCase();

          // Check if the new value matches any existing value in the name sheet
          var isDuplicate = nameValues.map(function(existingValue) {
            return existingValue.trim().toLowerCase();
          }).includes(newValue);

          if (!isDuplicate) {
            // Append the new value to the name sheet if it's not a duplicate
            nameSheet.appendRow([newValue]);
            Logger.log("Added '" + newValue + "' to sheet '" + nameSheet.getName() + "'");
          } else {
            Logger.log("'" + newValue + "' already exists in sheet '" + nameSheet.getName() + "'");
          }
          }
          break; // Stop further iteration for this cell if color matched
        }
      }
      }

    }

    
    // Calculate the starting column in the target sheet for this source sheet
    var startColumn = index * dataRange.getNumColumns() + 1;
    var startRow = 3;  // Start at row 2 (below the keys)

    Logger.log("Starting column for " + sheetName + ": " + startColumn);
    Logger.log("Starting row for " + sheetName + ": " + startRow);

    // Copy values to the target sheet
    targetSheet.getRange(startRow, startColumn, dataRange.getNumRows(), dataRange.getNumColumns()).setValues(values);
    
    // Copy backgrounds (colors) to the target sheet
    targetSheet.getRange(startRow, startColumn, dataRange.getNumRows(), dataRange.getNumColumns()).setBackgrounds(backgrounds);
  });

}

