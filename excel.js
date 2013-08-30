var xlsx        = require('xlsx'),
    xlsxBuilder = require('msexcel-builder');

function nodeExcel (config) {

  this.path       = config.path       || './' ;
  this.fileName   = config.fileName   || 'sample.xlsx';
  this.sheetName  = config.sheetName  || 'sheet1';
  this.columns    = config.columns    || 10;
  this.rows       = config.rows       || 12;
  this.content    = config.content    || [];

  var that = this;
}

nodeExcel.prototype.createExcel = function() {
  var workbook , sheet ;

  // Create a new workbook file in current working-path
  workbook = excelbuilder.createWorkbook (that.path, that.fileName );

  // Create a new worksheet 
  sheet = workbook.createSheet(that.sheetName, that.columns, that.rows);

  // Save it
  workbook.save( function (ok) {
    if (!ok ) workbook.cancel();
    else console.log('congratulations, your workbook created');
    return workbook;
  });
};



exports.updataExcel = function(config) {

}

exports.readExcel = function (config) {

}