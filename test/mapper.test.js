const Mapper = require('../src/Mapper');
const AdmZip = require('adm-zip');
const fs = require('fs');
const path = require('path');
const { DOMParser } = require('xmldom');

let sheetString = fs.readFileSync(path.resolve(__dirname, 'sheet1.xml'), {
    encoding: 'utf8'
});

let sharedStrings = fs.readFileSync(path.resolve(__dirname, 'sharedStrings.xml'), {
    encoding: 'utf8'
});
let parser = new DOMParser();

let sheetXml = parser.parseFromString(sheetString);

let sharedStringsXml = parser.parseFromString(sharedStrings);

let mapper = new Mapper(sharedStringsXml, sheetXml);

let header = mapper.mapHeaderToObject()
let object = mapper.mapBodyToObject();


for (let i = 0; i < object.length; i++) {
    console.log(object[i].rowNumber);
    console.log(object[i].cells);
}

console.log('=====================');

console.log(header.cells);