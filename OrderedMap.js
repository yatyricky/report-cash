class OrderedMap {
    constructor() {
        this._keys = []
        this._values = []
        this._dict = new Map()
    }

    add(key, value) {
        if (this._dict.has(key)) {
            throw new Error("key already presented in map")
        }
        const len = this._values.length
        this._keys.push(key)
        this._values.push(value)
        this._dict.set(key, len)
    }

    getValue(key, init) {
        const i = this._dict.get(key)
        if (i === undefined) {
            if (init === undefined) {
                return undefined
            }

            const newElem = init()
            this.add(key, newElem)
            return newElem
        }
        return this._values[i]
    }

    getKVPair(index) {
        return [this._keys[index], this._values[index]]
    }

    size() {
        return this._values.length
    }
}

module.exports = OrderedMap
