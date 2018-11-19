const assert = require('assert');
const Excel = require('../src/Excel');




const excel = new Excel('./Book1.xlsx');
const getWorksheet = excel.getWorksheet(1);
console.log(getWorksheet.body);