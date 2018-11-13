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
    }

    getRange(range) {
        if (!testRange(range)) {
            throw new Error('Range is not in correct format');
        }
    }

    /**
     * 
     * @param {number} sheetNumber
     * @return {{header, body}}
     * @throws {Error} 
     */
    getWorksheet(sheetNumber) {
        let xml = this.parser
            .parseFromString(this.admZip
                .readAsText(`xl/worksheets/sheet${sheetNumber}.xml`), 'text/xml');
        let mapper = new Mapper(this.sharedStrings, xml);

        return {
            header: mapper.mapHeaderToObject(),
            body: mapper.mapBodyToObject()
        }
    }
}

module.exports = Excel;