const fs = require("fs")
const xlsx = require("xlsx")
const { utils } = xlsx
const XlsxPopulate = require('xlsx-populate');

const KVArray = require("./KVArray.js")

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

const KW_INCOME = "收入"
const KW_EXPENDITURE = "支出"

function newAccountData() {
    return {
        data: new KVArray(),
        current: 0,
        total: 0
    }
}

const report = newAccountData()
report.data.add(newAccountData(), KW_INCOME)
report.data.add(newAccountData(), KW_EXPENDITURE)

async function program() {
    // const outWb = await XlsxPopulate.fromBlankAsync()
    // const outWb = utils.book_new()

    function flushReport(moonTag) {
        const newSheet = [
            [`个人借支明细账${moonTag}`],
            ["科目", "类别明细", "序号", "内容", "本期发生", "累计金额", "备注"],
        ]

        for (let i = 0; i < report.data.size(); i++) {
            const acc1 = report.data.getByIndex(i);
            for (let j = 0; j < acc1.data.size(); j++) {
                const acc2 = acc1.data.getByIndex(j);
                const row = [null, null, 0, null, 0, 0, ""]
                for (let k = 0; k < acc2.data.size(); k++) {
                    const entry = acc2.data.getByIndex(k);
                    if (k === 0) {
                        row[1] = acc1.data.getKey(j)
                        if (j === 0) {
                            row[0] = report.data.getKey(i)
                        }
                    }
                    row[2] = k + 1
                    row[3] = acc2.data.getKey(k)
                    row[4] = entry.current
                    row[5] = entry.total
                }
                newSheet.push(row)
            }
        }
        for (const row of newSheet) {
            console.log(row.join(","));
        }
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
        let acc1
        const acc2 = tokens[0]
        const acc3 = tokens.slice(1).join("_")

        if (tokens[0] === KW_INCOME) {
            acc1 = KW_INCOME
        } else {
            acc1 = KW_EXPENDITURE
        }

        const tab1 = report.data.getByKey(acc1, newAccountData)
        const tab = tab1.data.getByKey(acc2, newAccountData)
        const entry = tab.data.getByKey(acc3, () => {
            return { current: 0, total: 0 }
        })

        const delta = debit - credit
        entry.current += delta
        entry.total += delta
        tab.current += delta
        tab.total += delta
        tab1.current += delta
        tab1.total += delta
        report.current += delta
        report.total += delta
    }

    flushReport(currentMonth)
}

program().catch(e => {
    console.log("FATAL ERROR OCCURRED");
    console.log(e);
})
