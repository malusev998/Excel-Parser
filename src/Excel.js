const Mapper = require("./Mapper");
const AdmZip = require("adm-zip");
const xml = require("xmldom");
class Excel {

    /**
     * Constructor 
     * @param {string}  zipFile 
     * @param domParser Parser must support w3c level 2 compatibility, so everything can work as expected
     */
    constructor(zipFile, domParser = null) {
        /**
         * @var {DOMParser}
         */
        if (!domParser) {
            this.parser = new xml.DOMParser();
        } else {
            this.parser = domParser;
        }

        /**
         * @var {AdmZip}
         */
        this.zipper = new AdmZip(zipFile);
        /**
         * @var {Document}
         */
        const readText = this.zipper.readAsText("xl/sharedStrings.xml");
        this.sharedStrings = this.parser.parseFromString(readText, "text/xml");
        /**
         * @var {Row[]}
         */
        this.rows = null;
    }


    getSharedString() {
        return this.sharedStrings;
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
        if (typeof worksheet === "string") {
            row = this.getWorksheetByName(worksheet).body;
        } else if (typeof worksheet === "number") {
            if (this.rows === null) {
                this.getRow(worksheet);
            }
        }

        row = this.rows.filter(r => r.getRowNumber() === rowNumber);
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
            .parseFromString(this.zipper
                .readAsText(`xl/worksheets/sheet${worksheet}.xml`), "text/xml");
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

}

module.exports = Excel;