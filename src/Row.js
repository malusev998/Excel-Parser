class Row {
    /**
     * @param {Cell[]} cells cells in row
     * @param {number} rowNumber current row number
     * @param {string[]} keys header values
     */
    constructor(cells, rowNumber, isHeader = false, keys = []) {
        this.cells = cells;
        this.rowNumber = rowNumber;
        this.keys = keys;
        this.isHeader = isHeader;
    }


    /**
     * Gets the cell with the its excel position ex. A1, B6
     * @param {string} position 
     * @returns {Cell}
     */
    getCellByPosition(position) {
        return this.cells.filter(cell => cell.getPosition() === position)[0];
    }

    /**
     * @returns {Cell[]}
     */
    getAllCells() {
        return this.cells;
    }

    /**
     * @returns {number}
     */
    getRowNumber() {
        return this.rowNumber;
    }


    /**
     * Formats Excel row into javascript literal object
     * @returns {rowNumber, cells}
     */
    toObject() {
        let row = {
            rowNumber: this.rowNumber,
            cells: this.cells.map(
                (cell, i) => {
                    if (this.isHeader) {
                        return cell.toObject();
                    }
                    let r = {};
                    r[this.keys[i] === undefined ||
                        this.keys[i] === null ? i :
                        this.keys[i].toString()
                    ] = cell.toObject();
                    return r;
                }
            )
        };
        return row;
    }


}

module.exports = Row;

