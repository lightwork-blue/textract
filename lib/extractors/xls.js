var path = require('path'),
  XLSX = require('xlsx')
function extractText(filePath, options, cb) {
  var result, error

  try {
    const workbook = XLSX.readFile(filePath)
    result = ''

    // Convert each sheet to CSV and concatenate
    Object.keys(workbook.Sheets).forEach(function (sheetName) {
      const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName])
      result += csv
    })
  } catch (err) {
    error = new Error(
      'Could not extract ' + path.basename(filePath) + ', ' + err
    )
    cb(error, null)
    return
  }

  cb(null, result)
}

module.exports = {
  types: [
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
    'application/vnd.ms-excel.sheet.macroEnabled.12',
    'application/vnd.oasis.opendocument.spreadsheet',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
    'application/vnd.oasis.opendocument.spreadsheet-template',
  ],
  extract: extractText,
}
