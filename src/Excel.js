const Mapper = require('./Mapper');
const AdmZip = require('adm-zip');
/**
 * 
 * @param {string} range
 * @returns {boolean} 
 */
function testRange(range) {
    // TODO: Needs implemenation
    throw new Error('Function is no implemented');
}
class Excel {
    /**
     * 
     * @param {DOMParser} parser 
     * @param {string}  zipFile 
     */
    constructor(parser, zipFile) {
        this.parser = parser;
        this.zipper = new AdmZip(zipFile);
        this.sharedStrings = this.parser.parseFromString(this.zipper.readAsText('xl/sharedStrings.xml'), 'text/xml');
        this.rows = null;
    }

    getRange(range) {
        if (!testRange(range)) {
            throw new Error('Range is not in correct format');
        }
    }

    /**
     * 
     * @param {number} worksheet 
     * @returns {Row[]}
     */
    getRows(worksheet = 1) {
        let xml = this.parseXML(worksheet);
        let mapper = new Mapper(this.sharedStrings, xml);

        this.rows = [mapper.mapHeader(), ...mapper.mapBody()];
        return this.rows;
    }

    getRow(rowNumber, worksheet = 1) {
        if (this.rows === null) {
            return this.getRow(worksheet)[rowNumber];
        }

        return this.rows[rowNumber];
    }

    /**
     * Do not use this method
     * @private
     * @param {number} worksheet 
     */
    parseXML(worksheet) {
        return this.parser
            .parseFromString(this.admZip
                .readAsText(`xl/worksheets/sheet${worksheet}.xml`), 'text/xml');
    }
    /**
     * 
     * @param {number} sheetNumber
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
}

module.exports = Excel;