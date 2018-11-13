class Row {
    /**
     * @param {Cell[]} cells
     * @param {number} rowNumber
     * @param {string[]} keys
     */
    constructor(cells, rowNumber, isHeader = false, keys = []) {
        this.cells = cells;
        this.rowNumber = rowNumber;
        this.keys = keys;
        this.isHeader = isHeader;
    }

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

