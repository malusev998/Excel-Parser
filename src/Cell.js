class Cell {
    /**
     * Constructor
     * @param {object|string|number|Date|null} value value contained inside a cell 
     * @param {string} type  type of data inside a cell
     * @param {string} position excel position eg A1
     */
    constructor(value, type, position) {
        /**
         * @var {string|number|Date|object|null}
         */
        this.value = value;
        /**
         * @var {string}
         */
        this.type = type;
        /**
         * @var {string}
         */
        this.position = position;
    }


    /**
     * Returns the type of data stored in cell
     * @returns {string}
     */
    getType() {
        return this.type;
    }

    /**
     * Returns the value of the cell
     * @returns {string|number|object|Date|null}
     */
    getValue() {
        return this.value;
    }

    /**
     * Returns position of the cell in excel
     * @return {string}
     */
    getPosition() {
        return this.position;
    }


    toString() {
        return `${this.position}: ${this.value}`;
    }

    /**
     * Maps cell object into object literal
     * @returns { { value, position, type } }
     */
    toObject() {
        return {
            value: this.value,
            position: this.position,
            type: this.type
        }
    }
}

module.exports = Cell;