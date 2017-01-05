(function() {
  'use strict';
  
  var data;
  var spinnerComponent;
  
  Office.initialize = function(reason) {
    $(document).ready(function() {
    
    // Initialize spinner
    var element = document.querySelector('.ms-Spinner');
    spinnerComponent = new fabric['Spinner'](element);

    // Hide the spinner initially
    $('.ms-Spinner').hide();
  
    $('#load-data').click(loadData);    
    $('#load-first-data').click(loadData);
    $('#load-last-data').click(loadData);
  });
};

function loadData() {
  $('.ms-Spinner').show();
  spinnerComponent.start();
  
  var size = $(this).data('load');
  console.log(size);

  var sheetName = 'Categories_' + size;
  var tableName = sheetName;
  
  $.getJSON('https://gist.githubusercontent.com/renil/fef6142dfe8e707061a399cca7fa1d32/raw/a17597dfe1b8a1c81e92e0d89a853d1aa00b31d2/data.json', function (result) {    
    Excel.run(function (ctx) {
      var sheet = ctx.workbook.worksheets.add(sheetName);
      sheet.getRange().format.fill.color = 'white';
      
      var startRowIndex = 3;
      var startColumnIndex = 1
      
      var columnHeadersRowIndex = startRowIndex + 1;
      var tableStartRowIndex = columnHeadersRowIndex + 1;
      
      var data = getTableData(result.details, columnHeadersRowIndex, tableStartRowIndex, startColumnIndex, size);
      
      var endColumnIndex = data.headerValues.length;
      var startColumnName = indexToName(startColumnIndex);
      var endColumnName = indexToName(endColumnIndex);
      var endRowIndex = tableStartRowIndex + data.values.length;
      
      // Create table
      var tableRange = "'" + sheetName + "'!" + startColumnName + tableStartRowIndex + ":" + endColumnName + endRowIndex;
      var dataTable = ctx.workbook.tables.add(tableRange, true);
      dataTable.name = tableName;
      dataTable.showTotals = true;
      
      // Set the table data
      dataTable.getHeaderRowRange().values = [data.headerValues];
      dataTable.getDataBodyRange().formulas = data.values;
      dataTable.getTotalRowRange().formulas = [data.totalRow];
      
      // Format the table
      dataTable.style = 'TableStyleMedium23';

      return ctx.sync().then(function() {
        spinnerComponent.stop();
        $('.ms-Spinner').hide();
      }).catch(function (error) {
        spinnerComponent.stop();
        $('.ms-Spinner').hide();
        if (error instanceof OfficeExtension.Error) {
            console.log(JSON.stringify(error));
        }
      });;
    });
  });
}
  
function getTableData(data, columnHeaderRowIndex, startRowIndex, startColumnIndex, size) {    
  var isFirstRow = true;
  var categoryHeaders = ['', ''];
  var columnHeaders = ['Jurisdiction Id', 'Jurisdiction'];
  var headerValues = ['Jurisdiction Id', 'Jurisdiction'];
  var totalRow = ['', 'Total Everywhere'];
  var startColumnName = indexToName(startColumnIndex);
  var endColumnName = '';
  var totalBeginningColumnName = '';
  var totalEndingColumnName = '';
  var values = [];
  var currentRow = startRowIndex + 1; // Add 1 for the header row  
  _.each(data, function (jurisdiction) {
    var temp = [];
    temp.push(jurisdiction.jurisdictionId, jurisdiction.jurisdiction);
    _.each(jurisdiction.apportionments, function (apportionment) {      
      if (isFirstRow) {
        categoryHeaders.push(apportionment.category, '', '', '');
        columnHeaders.push('Beginning', 'Ending Raw', 'Allocation', 'Ending');
        headerValues.push(apportionment.category + ' Beginning', apportionment.category + ' Ending Raw', apportionment.category + ' Allocation', apportionment.category + ' Ending');
        totalRow.push("=SUBTOTAL(109,[" + sanitizeExcelColumnNameForFormula(apportionment.category) + " Beginning])", "=SUBTOTAL(109,[" + sanitizeExcelColumnNameForFormula(apportionment.category) + " Ending Raw])", '', "=SUBTOTAL(109,[" + sanitizeExcelColumnNameForFormula(apportionment.category) + " Ending])");
      }
      temp.push(apportionment.beginningAmount ? apportionment.beginningAmount : '', apportionment.endingRawAmount ? apportionment.endingRawAmount : '', apportionment.allocationAmount ? apportionment.allocationAmount : '');
      
      // Ending Balance is calculated by summing Ending Raw amount and Allocation Amount
      temp.push("=" + indexToName(temp.length - 1) + currentRow + "+" + indexToName(temp.length) + currentRow);
    });
    if (isFirstRow) {
      categoryHeaders.push('', '');
      columnHeaders.push('Beginning', 'Ending');
      endColumnName = indexToName(headerValues.length);
      headerValues.push('Total Beginning', 'Total Ending');
      totalRow.push('=SUBTOTAL(109,[Total Beginning])', '=SUBTOTAL(109,[Total Ending])');
      totalBeginningColumnName = indexToName(headerValues.length - 1);
      totalEndingColumnName = endColumnName;
    }
    
    // Add the Beginning Ending column values
    var beginningTotal = ("=SUMIF($" + startColumnName + "$" + columnHeaderRowIndex + ":$" + endColumnName + "$" + columnHeaderRowIndex + ",");
    beginningTotal = beginningTotal + (totalBeginningColumnName + "$" + columnHeaderRowIndex + ", " + startColumnName + currentRow + ":" + endColumnName + currentRow + ")");
    var endingToal = ("=SUMIF($" + startColumnName + "$" + columnHeaderRowIndex + ":$" + endColumnName + "$" + columnHeaderRowIndex + ",");
    endingToal = endingToal + (totalEndingColumnName + "$" + columnHeaderRowIndex + ", " + startColumnName + currentRow + ":" + endColumnName + currentRow + ")");
    temp.push(beginningTotal, endingToal);
    values.push(temp);
    isFirstRow = false;
    currentRow++;
  });
  return {
    categoryHeaders: categoryHeaders,
    columnHeaders: columnHeaders,
    headerValues: headerValues,
    values: values,
    totalRow: totalRow
  };
}
  
function indexToName(index){
  var num = index;
  var name = '';
  while (num > 0) {
    var modulo = (num - 1) % 26;
    name = String.fromCharCode(modulo + 65) + name;
    num = parseInt(((num - modulo) / 26).toString(), 10);
  }
  return name;
}

function sanitizeExcelColumnNameForFormula(name) {
    var result = name
        .replace(/[']/g, '\'\'');
    return result;
}
})();