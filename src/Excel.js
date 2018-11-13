const Mapper = require('./Mapper');
const AdmZip = require('adm-zip');
const { DOMParser } = require('xmldom');
class Excel {

    /**
     * Constructor 
     * @param {string}  zipFile 
     */
    constructor(zipFile) {
        /**
         * @var {DOMParser}
         */
        this.parser = new DOMParser();

        /**
         * @var {AdmZip}
         */
        this.zipper = new AdmZip(zipFile);
        /**
         * @var {Document}
         */
        this.sharedStrings = this.parser.parseFromString(this.zipper.readAsText('xl/sharedStrings.xml'), 'text/xml');
        /**
         * @var {Row[]}
         */
        this.rows = null;
    }


    /**
     * Gets rows and cells by excel notations eg. A1:C5
     * @param {string} range
     */
    getRange(range) {
        if (!testRange(range)) {
            throw new Error('Range is not in correct format');
        }
    }

    /**
     * Gets all rows from excel worksheet
     * @param {number} worksheet 
     * @returns {Row[]}
     * @return { Array<{rowNumber, cells}> }
     */
    getRows(worksheet = 1) {
        let xml = this.parseXML(worksheet);
        let mapper = new Mapper(this.sharedStrings, xml);

        this.rows = [mapper.mapHeader(), ...mapper.mapBody()];
        return this.rows;
    }

    /**
     * Return single row from excel worksheet
     * @param {number} rowNumber 
     * @param {number|string} worksheet 
     */
    getRow(rowNumber, worksheet = 1) {
        let row;
        if (typeof worksheet === 'string') {
            row = this.getWorksheetByName(worksheet).body;
        } else if (typeof worksheet === 'number') {
            if (this.rows === null) {
                this.getRow(worksheet);
            }
        }

        let row = this.rows.filter(r => r.getRowNumber() === rowNumber);
        return row.length === 0 ? null :
            this.rows[rowNumber];
    }

    /**
     * Do not use this method outside of a class
     * @private
     * @param {number} worksheet 
     * @returns {Document}
     */
    parseXML(worksheet) {
        return this.parser
            .parseFromString(this.admZip
                .readAsText(`xl/worksheets/sheet${worksheet}.xml`), 'text/xml');
    }

    /**
     * Do not use this method outside of a class
     * @private
     * @param {string} range 
     */
    testRange(range) {
        // TODO: Needs implemenation
        throw new Error('Function is no implemented');
    }

    /**
     * Gets all rows from excel forksheet
     * @param {number} sheetNumber worksheet number, begins at one
     * @return {{header, body}}
     * @throws {Error} 
     */
    getWorksheet(sheetNumber = 1) {
        let xml = this.parseXML(sheetNumber);
        let mapper = new Mapper(this.sharedStrings, xml);

        return {
            header: mapper.mapHeaderToObject(),
            body: mapper.mapBodyToObject()
        }
    }

    /**
     * Gets Excel Worksheet by name 
     * @param {string} name
     * @throws {Error} 
     * @returns { {header: {rowNumber, cells}, body: Array<{rowNumber, cells}> } }
     */
    getWorksheetByName(name) {
        // TODO: Find file that contains names of excel worksheet and return getWorksheet
    }
}

module.exports = Excel;