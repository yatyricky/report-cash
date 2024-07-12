const fs = require("fs")
const xlsx = require("xlsx")
const { utils } = xlsx

const config = JSON.parse(fs.readFileSync("config.json", "utf8"))

const wb = xlsx.readFile(config.filePath)
const ws = wb.Sheets[config.sheetName]

/**
 * @returns {Date}
 */
function GetDate(ws, r, c) {
    const cell = ws[utils.encode_cell({ r, c })]
    if (cell === undefined || cell === null || cell.t !== "n") {
        return new Date()
    }
    return new Date(-2209161600000 + 864E5 * (cell.v || 0))
}

/**
 * @returns {String}
 */
function GetString(ws, r, c) {
    const cell = ws[utils.encode_cell({ r, c })]
    if (cell === undefined || cell === null || cell.t !== "s") {
        return ""
    }
    return cell.v || ""
}

/**
 * @returns {Number}
 */
function GetDouble(ws, r, c) {
    const cell = ws[utils.encode_cell({ r, c })]
    if (cell === undefined || cell === null || cell.t !== "n") {
        return 0
    }
    return cell.v
}

/**
 * @returns {String}
 */
function padWith0(inStr, len) {
    let str = "" + inStr
    let strLen = str.length
    while (strLen < len) {
        str = "0" + str
        strLen++
    }
    return str
}

function GetOrCreateObject(obj, key) {
    let o = obj[key]
    if (o === undefined) {
        const n = Object.keys(obj).length
        o = { sort: n, sum: 0 }
        obj[key] = o
    }
    return o
}

function ObjectToArray(obj) {
    const arr = []
    for (const key in obj) {
        if (Object.hasOwnProperty.call(obj, key)) {
            const value = obj[key];
            arr.push({ _k: key, _v: value })
        }
    }
    arr.sort((a, b) => {
        return b.sort - a.sort
    })
    return arr
}

const report = {
    income: {
        ["收入"]: { sort: 0, sum: 0 } // key=account, value={sort:1, current:12, total:100}
    },
    outcome: {}, // key=secondary, value={bulk}
}

const outWb = utils.book_new()
const outSheets = []

function flushReport(moonTag) {
    const newSheet = [
        [`孙世龙个人借支明细账${moonTag}`],
        ["科目", "类别明细", "序号", "内容", "本期发生", "累计金额", "备注"],
    ]

    // income
    const first = ObjectToArray(report.income)
    for (let i = 0; i < first.length; i++) {
        const data = first[i];
        const second = ObjectToArray(data)
        for (let j = 0; j < second.length; j++) {
            const entry = second[j];
            const row = []
            if (i === 0) {
                row.push("收入")
            } else {
                row.push(null)
            }
            if (i === 0) {
                row.push(data._k)
            } else {
                row.push(null)
            }
            row.push(j + 1)
            row.push()
        }

    }

    const ws = utils.aoa_to_sheet(newSheet)
    ws["!merges"] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }
        // { s: { r: 2, c: 0 }, e: { r: 5, c: 0 } }
    ]

    outSheets.push({ ws, name: moonTag })
}

const range = utils.decode_range(ws["!ref"])
let currentMonth
for (let r = 1; r <= range.e.r; r++) {
    const account = GetString(ws, r, 3).trim()
    if (account.length === 0) {
        continue
    }

    const date = GetDate(ws, r, 0)

    const moonTag = `${date.getFullYear()}-${padWith0(date.getMonth() + 1, 2)}`
    if (currentMonth !== undefined && moonTag !== currentMonth) {
        flushReport(moonTag)
    }
    currentMonth = moonTag

    const desc = GetString(ws, r, 1)
    const debit = GetDouble(ws, r, 4)
    const credit = GetDouble(ws, r, 5)

    const tokens = account.split("_")
    const acc1 = tokens[0]
    const acc2 = tokens.slice(1).join("_")

    let tab
    if (acc1 === "收入") {
        tab = report.income.income
    } else {
        tab = GetOrCreateObject(report.outcome, acc1)
    }

    const entry = GetOrCreateObject(tab, acc2)
    if (entry.current === undefined) {
        entry.current = 0
    }
    if (entry.total === undefined) {
        entry.total = 0
    }

    const delta = debit - credit
    entry.current += delta
    entry.total += delta
    tab.sum += delta
}

flushReport(currentMonth)

for (let i = outSheets.length - 1; i >= 0; i--) {
    const data = outSheets[i]
    utils.book_append_sheet(outWb, data.ws, data.name, true)
}


xlsx.writeFileXLSX(outWb, "test.xlsx")

console.log(ws.A1.s);
