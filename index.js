const path = require("path")
const xlsx = require("xlsx")
const XlsxPopulate = require('xlsx-populate');
const OrderedMap = require("./OrderedMap.js")
const commander = require("commander");

const DefaultOutFp = "<InputFilePath>-<yyyymm>.xlsx"

commander.program.requiredOption("-i, --input <InputFilePath>", "Input file path")
commander.program.requiredOption("-s, --sheet <InputSheetName>", "The sheet to be parsed")
commander.program.option("-o, --output <OutputFilePath>", "Output file path", DefaultOutFp)
commander.program.option("-r --rowstart <RowStartAt>", "Specify the first row contains data", "2");
commander.program.option("--cdate <ColumnName>", "Specify the date column", "A");
commander.program.option("--cdescription <ColumnName>", "Specify the description column", "B");
commander.program.option("--caccount <ColumnName>", "Specify the account column", "D");
commander.program.option("--accountmarker <TheMarker>", "Account begins after this string", "*");
commander.program.option("--cdebit <ColumnName>", "Specify the debit column", "E");
commander.program.option("--ccredit <ColumnName>", "Specify the credit column", "F");

commander.program.addHelpText("afterAll", `
Example:
$ report_cash -i /path/to/file.xlsx -s Sheet2
$ report_cash -i /path/to/file.xlsx -s Sheet2 -o /path/to/report/out-01.xlsx -r5 --cdebit=C --ccredit=D --caccount=G
`)

commander.program.parse()
const options = commander.program.opts()

const { utils } = xlsx

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

const inFilePath = options.input
let outFilePath = options.output
if (outFilePath === DefaultOutFp) {
    const parsedInFp = path.parse(inFilePath)
    const now = new Date()
    outFilePath = path.join(parsedInFp.dir, `${parsedInFp.name}-${now.getFullYear()}${padWith0(now.getMonth() + 1, 2)}${padWith0(now.getDate(), 2)}.xlsx`)
}

const wb = xlsx.readFile(inFilePath)
const ws = wb.Sheets[options.sheet]

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

function round(n) {
    return Math.round(n * 100) / 100
}

const KW_INCOME = "收入"
const KW_EXPENDITURE = "支出"

function newAccountData() {
    return {
        data: new OrderedMap(),
        current: 0,
        total: 0
    }
}

async function program() {
    const report = newAccountData()
    report.data.add(KW_INCOME, newAccountData())
    report.data.add(KW_EXPENDITURE, newAccountData())

    const outWb = await XlsxPopulate.fromBlankAsync()

    let prevBalance
    function flushReport(moonTag) {
        const outWs = outWb.addSheet(moonTag, 0)
        outWs.cell("A1").value(`个人借支明细账${moonTag}`)
        outWs.column("B").width(12)
        outWs.column("C").width(6)
        outWs.column("D").width(30)
        outWs.column("E").width(15).style({ numberFormat: "0.00" })
        outWs.column("F").width(15).style({ numberFormat: "0.00" })
        outWs.range("A1:G1").merged(true).style({ horizontalAlignment: "center", fontSize: 16 })
        outWs.range("A2:G2").value([["科目", "类别明细", "序号", "内容", "本期发生", "累计金额", "备注"]]).style({ bold: true })
        let r = 3

        const totalRows = []
        for (let i = 0; i < report.data.size(); i++) {
            const [acc1Key, acc1Value] = report.data.getKVPair(i);
            const credit = acc1Key === KW_EXPENDITURE ? -1 : 1
            let rowSpan = 0
            const subTotalRows = []
            for (let j = 0; j < acc1Value.data.size(); j++) {
                const [acc2Key, acc2Value] = acc1Value.data.getKVPair(j);
                const acc2Len = acc2Value.data.size()
                for (let k = 0; k < acc2Len; k++) {
                    const row = [null, null, 0, null, 0, 0, ""]
                    const [entryKey, entryValue] = acc2Value.data.getKVPair(k);
                    if (k === 0) {
                        row[1] = acc2Key
                        if (j === 0) {
                            row[0] = acc1Key
                        }
                    }
                    row[2] = k + 1
                    row[3] = entryKey
                    row[4] = round(entryValue.current * credit)
                    row[5] = round(entryValue.total * credit)
                    outWs.range(`A${r}:G${r}`).value([row])
                    r++
                }
                outWs.range(`B${r - acc2Len}:B${r - 1}`).merged(true).style({ verticalAlignment: "center" })
                const rowSubtotal = [null, "小计", null, null, round(acc2Value.current * credit), round(acc2Value.total * credit)]
                outWs.range(`A${r}:G${r}`).value([rowSubtotal])
                outWs.cell(`E${r}`).formula(`SUM(E${r - acc2Len}:E${r - 1})`)
                outWs.cell(`F${r}`).formula(`SUM(F${r - acc2Len}:F${r - 1})`)
                subTotalRows.push(r)
                outWs.range(`B${r}:D${r}`).merged(true)
                outWs.range(`B${r}:G${r}`).style({ fill: "DDDDDD" })
                r++

                rowSpan += acc2Len + 1
            }
            const rowGrandTotal = ["合计", null, null, null, round(acc1Value.current * credit), round(acc1Value.total * credit)]
            outWs.range(`A${r}:G${r}`).value([rowGrandTotal]).style({ fill: "BBBBBB" })
            outWs.cell(`E${r}`).formula(subTotalRows.map(e => `E${e}`).join("+"))
            outWs.cell(`F${r}`).formula(subTotalRows.map(e => `F${e}`).join("+"))
            outWs.range(`A${r}:D${r}`).merged(true)
            totalRows.push(r)
            if (rowSpan > 0) {
                outWs.range(`A${r - rowSpan}:A${r - 1}`).merged(true).style({ verticalAlignment: "center" })
            }
            r++
        }
        const rowBalance = ["余额", null, null, null, round(report.current), round(report.total)]
        outWs.range(`A${r}:G${r}`).value([rowBalance]).style({ bold: true })
        outWs.range(`A${r}:D${r}`).merged(true)
        outWs.cell(`E${r}`).formula(`E${totalRows[0]}-E${totalRows[1]}`)
        outWs.cell(`F${r}`).formula(`F${totalRows[0]}-F${totalRows[1]}`)
        if (prevBalance !== undefined) {
            outWs.cell(`G${r}`).formula(`${prevBalance}+E${r}`).style({ numberFormat: "0.00" })
        }
        prevBalance = `'${moonTag}'!F${r}`

        outWs.range(`A2:G${r}`).style({ borderStyle: "thin" })

        // clear current
        report.current = 0
        for (let i = 0; i < report.data.size(); i++) {
            const [_1, acc1Value] = report.data.getKVPair(i);
            acc1Value.current = 0
            for (let j = 0; j < acc1Value.data.size(); j++) {
                const [_2, acc2Value] = acc1Value.data.getKVPair(j);
                acc2Value.current = 0
                for (let k = 0; k < acc2Value.data.size(); k++) {
                    const [_3, entryValue] = acc2Value.data.getKVPair(k);
                    entryValue.current = 0
                }
            }
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
        if (currentMonth === undefined) {
            currentMonth = moonTag
        }
        if (moonTag !== currentMonth) {
            flushReport(currentMonth)
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

        const tab1 = report.data.getValue(acc1, newAccountData)
        const tab = tab1.data.getValue(acc2, newAccountData)
        const entry = tab.data.getValue(acc3, () => {
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
    outWb.deleteSheet("Sheet1")
    outWb.activeSheet(currentMonth)

    await outWb.toFileAsync(outFilePath)
}

program().catch(e => {
    console.log("FATAL ERROR OCCURRED");
    console.log(e);
})
