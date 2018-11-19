
const Cell = require("./Cell");
const Row = require("./Row");

class Mapper {
    /**
     * Constructor
     * @param {Document} sharedStrings
     * @param {Document} sheetXml
     */
    constructor(sharedStrings, sheetXml) {
        this.sharedStrings = sharedStrings;
        this.sheet = sheetXml;
    }

    /**
     * Parses the first row inside excel file
     * @returns {Row}
     */
    mapHeader() {
        return new Row(this.mapCells(0), 1, true);
    }

    /**
     * Maps excel header row to javascript literal object
     * @returns {{rowNumber, cells}[]}
     */
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
        if (typeof rowNumber !== "number") {
            throw new Error("rowNumber is not type of number");
        }
        let row = this.sheet.getElementsByTagName("row")[rowNumber];
        let cells = [];
        let children = row.getElementsByTagName("c");
        for (let i = 0; i < children.length; i++) {
            let value = this.mapType(children[i].getAttribute("t"), children[i].childNodes[0].textContent);
            cells.push(new Cell(value, typeof value, children[i].getAttribute("r")));
        }
        return cells;

    }

    /**
     * Maps the cells to an object literal
     * @returns { { value, position, type }[] }
     */
    mapCellsToObject() {
        return this.mapCells.map(cell => cell.toObject());
    }

    /**
     * Parsers the rest of the excel file
     * Without the header row
     * @returns {Row[]}
     */
    mapBody() {
        const header = this.mapHeaderToObject().cells.map(cell => cell.value);
        const body = [];
        for (let i = 1; i < this.sheet.getElementsByTagName("row").length; i++) {
            body.push(new Row(this.mapCells(i), i + 1, false, header));
        }
        return body;
    }

    /**
     * Parsers the rest of the excel file
     * Without the header row
     * Header row is used for keys in object
     * @returns { { {rowNumber, cells} }[] }
     */
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
            .getElementsByTagName("si")[sharedStringPosition]
            .firstChild
            .textContent;
    }

    /**
     * Maps the element value based on the type in excel cell
     * @param {string} type 
     * @param {ChildNode}
     */
    mapType(type, cell) {
        // console.log(cell);
        switch (type) {

            case "":
                // TODO: Add More type checks with regex
                return cell === null ? null : parseFloat(cell);
            case "s":
                return cell === null ? null :
                    this.findInSharedString(parseInt(cell));
        }
    }
}


module.exports = Mapper;
