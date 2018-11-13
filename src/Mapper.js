const AdmZip = require('adm-zip');
const Cell = require('./Cell');
const Row = require('./Row');

class Mapper {
    /**
     * Constructor
     * @param {DOMParser} parser
     * @param {AdmZip} 
     * @param {number} sheet
     */
    constructor(sharedStrings, sheetXml) {
        this.parser = parser;
        this.admZip = admZip;
        this.sharedStrings = sharedStrings;
        this.sheet = sheetXml

        /**
         * @var {Row}
         */
        this.header = null;
        /**
         * @var {Row[]}
         */
        this.rows = [];
    }

    mapHeader() {
        return new Row(this.mapCells(1), 1, true);
    }

    mapHeaderToObject() {
        return this.mapHeader().toObject();
    }

    /**
     * Maps the cells in excel to object
     * @param {number} rowNumber 
     * @throws {Error}
     * @returns {Cell[]}
     */
    mapCells(rowNumber) {
        if (typeof rowNumber !== 'number') {
            throw new Error('rowNumber is not type of number');
        }

        try {
            let row = this.sheet.getElementsByTagName('row')[rowNumber];
            let cells = [];
            for (let cell of row.children) {
                let value = this.mapType(cell.getAttribute('t'));
                cells.push(new Cell(value, typeof value, cell.getAttribute('r')));
            }
            return cells;
        } catch (err) {
            throw err;
        }

    }

    /**
     * @returns {Row[]}
     */
    mapBody() {
        let header = this.mapHeaderToObject().cells.map(cell => cell.value);
        let body = [];
        for (let i = 1; i < this.sheet.getElementsByTagName('row').length; i++) {
            body.push(new Row(this.mapCells(i), i + 1, false, header));
        }
        return body;
    }


    mapBodyToObject() {
        return this.mapBody().map(r => r.toObject());
    }
    /**
     * Finds appropriate string inside sharedStrings.xml
     * @param {number} sharedStringPosition 
     * @returns {string}
     */
    findInSharedString(sharedStringPosition) {
        return this.sharedStrings
            .getElementsByTagName('si')[sharedStringPosition]
            .firstChild
            .nodeValue;
    }

    /**
     * Maps the element value based on the type in excel cell
     * @param {string} type 
     */
    mapType(type) {
        switch (type) {
            case null:
                return cell.firstElementChild === null ? null : parseFloat(cell.firstElementChild.nodeValue);
            case 's':
                return cell.firstElementChild.nodeValue === null ? null :
                    this.findInSharedString(parseInt(cell.firstElementChild.nodeValue));
        }
    }
}


module.exports = Mapper;
