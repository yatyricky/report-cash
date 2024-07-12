class KVArray {
    constructor() {
        this._arr = []
        this._dict = new Map()
        this._key = []
    }

    add(elem, key) {
        const len = this._arr.length
        if (key === undefined) {
            key = `__key${len}`
        }
        this._arr.push(elem)
        this._dict.set(key, len)
        this._key.push(key)
    }

    getByKey(key, init) {
        const i = this._dict.get(key)
        if (i === undefined) {
            if (init !== undefined) {
                const newElem = init()
                this.add(newElem, key)
                return newElem
            } else {
                return undefined
            }
        }
        return this._arr[i]
    }

    getKey(index) {
        return this._key[index]
    }

    getByIndex(index) {
        return this._arr[index]
    }

    size() {
        return this._arr.length
    }
}

module.exports = KVArray
