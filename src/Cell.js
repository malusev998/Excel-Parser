class Cell {
    /**
     * 
     * @param {object|string|number} value 
     * @param {string} type 
     * @param {string} position 
     */
    constructor(value, type, position) {
        this.value = value;
        this.type = type;
        this.position = position;
    }

    getType() {
        return this.type;
    }

    getValue() {
        return this.value;
    }

    getPosition() {
        return this.position;
    }

    toString() {
        return `${this.position}: ${this.value}`;
    }

    toObject() {
        return {
            value: this.value,
            position: this.position,
            type: this.type
        }
    }
}

module.exports = Cell;