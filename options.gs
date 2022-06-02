function getObjects(specSheet, column, lastRow, numCols = 1) {
  return specSheet.getRange(6, column, lastRow, numCols).getValues()
          .map(function(arr, i)
          { 
            return arr.concat(i+6);
          })
          .filter(function(e)
          {
            return e[0] != "";
          });
}
