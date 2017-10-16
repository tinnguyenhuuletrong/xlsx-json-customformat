const Excel = require("exceljs")
const fs = require("fs")

//---------------------------------------------------------------------------//
//---------------------------------------------------------------------------//
function exportArrayObject(sheet, config) {
	const numColumns = sheet.columnCount;
	const numRows = sheet.actualRowCount
	const headerRow = config.headerRow || 1
	const dataRowBegin = config.dataRowBegin || headerRow + 1
	const ignore = config.ignore || []

	const headers = sheet.getRow(headerRow).values

	const outputs = []
	for (var i = 0; i < numRows; i++) {
		const rowData = sheet.getRow(dataRowBegin + i)
		const tmp = {}
		rowData.eachCell((cell, colNum) => {
			if (ignore.indexOf(headers[colNum]) != -1)
				return
			tmp[headers[colNum]] = cell.value
		})
		outputs.push(tmp)
	}

	return outputs
}

function exportArrayValue(sheet, config) {
	const numColumns = sheet.columnCount;
	const numRows = sheet.actualRowCount
	const headerRow = config.headerRow || 1
	const dataRowBegin = config.dataRowBegin || headerRow + 1
	const ignore = config.ignore || []

	const valueCol = config.valueCol || "Value"
	const headers = sheet.getRow(headerRow).values
	const valueColIndex = headers.indexOf(valueCol)

	const outputs = []
	for (var i = 0; i < numRows; i++) {
		const rowData = sheet.getRow(dataRowBegin + i)
		if (rowData.values[valueColIndex] != null)
			outputs.push(rowData.values[valueColIndex])
	}

	return outputs
}

function exportObject(sheet, config) {
	const numColumns = sheet.columnCount;
	const numRows = sheet.actualRowCount
	const headerRow = config.headerRow || 1
	const dataRowBegin = config.dataRowBegin || headerRow + 1
	const ignore = config.ignore || []
	const keyCol = config.keyCol || "Key"

	// ignore key col by default
	ignore.push(keyCol)

	const headers = sheet.getRow(headerRow).values
	const keyColIndex = headers.indexOf(keyCol)

	const outputs = {}
	for (var i = 0; i < numRows; i++) {
		const rowData = sheet.getRow(dataRowBegin + i)
		const tmp = {}
		rowData.eachCell((cell, colNum) => {
			if (ignore.indexOf(headers[colNum]) != -1)
				return
			tmp[headers[colNum]] = cell.value
		})

		if (rowData.values[keyColIndex] != null)
			outputs[rowData.values[keyColIndex]] = tmp
	}

	return outputs
}

//---------------------------------------------------------------------------//
const EXPORT_TYPE_MAP = {
	"array_object": exportArrayObject,
	"array_value": exportArrayValue,
	"object": exportObject
}


let inputFile = null
let inputFormatFile = null
let outputFile = null


const args = process.argv
for (var i = 0; i < args.length; i++) {
	if(args[i] == "-in")
		inputFile = args[i+1]
	else if(args[i] == "-config")
		inputFormatFile = args[i+1]
	else if(args[i] == "-out")
		outputFile = args[i+1]
}

//Debug
// let inputFile = "./test/GameConfig.xlsx"
// let inputFormatFile = "./test/format.json"
// let outputFile = "./test/out.json"

if(!inputFile || !inputFormatFile || !outputFile) {
	console.log("Usage: -in <input xlsx> -out <output json> -config <config json>")
	process.exit(0)
}

const begin = Date.now()

let Format = null
try {
	const formatInput = fs.readFileSync(inputFormatFile, 'utf8')
	Format = JSON.parse(formatInput)

} catch (ex) {
	console.error(ex);
	process.exit(0)
}

const RESULT = {}

// read from a file
var workbook = new Excel.Workbook();
workbook.xlsx.readFile(inputFile)
	.then(function() {
		for (var key in Format) {
			console.log("[Info] export sheet", key)

			var sheet = workbook.getWorksheet(key)
			const config = Format[key]
			if (sheet == null) {
				console.warning("[Warning] Could not found sheet", key)
				continue
			}

			const exportMethod = EXPORT_TYPE_MAP[config.export]
			if (exportMethod == null) {
				console.warning("[Warning] Export method not found", config.export)
				continue
			}

			RESULT[key] = exportMethod(sheet, config)
		}

		// Write output
		fs.writeFile(outputFile, JSON.stringify(RESULT), err => {
			console.log("Complete - Running time: ", Date.now() - begin, "ms")
		})
	})
	.catch(err => {
		console.error("[Error] ", err)
	})

