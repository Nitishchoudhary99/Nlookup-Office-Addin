/**
 * Custom NLOOKUP function for Excel.
 * @param {string|number} lookupValue - The value to search for.
 * @param {Array} lookupArray - The column range to search in.
 * @param {Array} returnArray - The column range to return values from.
 * @returns {string|number} The matched value or "Not Found" if no match.
 */
function NLOOKUP(lookupValue, lookupArray, returnArray) {
  for (let i = 0; i < lookupArray.length; i++) {
      if (lookupArray[i][0] === lookupValue) {
          return returnArray[i][0];
      }
  }
  return "Not Found";
}

// Register the function in Excel
Excel.Script.CustomFunctions.associate("NLOOKUP", NLOOKUP);
