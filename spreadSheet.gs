const SHEET_ID = "xxxxxxxxxxxxxxxxxxxxxx"
const FOLDER_ID = "xxxxxxxxxxxxxxxxxxxxxx"

// この関数を実行する
const convert = () => {
  const sheets = SpreadsheetApp.openById(SHEET_ID).getSheets()
  const folder = DriveApp.getFolderById(FOLDER_ID)
  for (const sheet of sheets) {
    saveAsJson(sheet, folder)
  }
}

const saveAsJson = (sheet, folder) => {
  const fileName = sheet.getSheetName() + ".json"
  console.log("save: " + fileName)
  const jsonStr = JSON.stringify(tableToJson(sheet))
  const file = getFileFromFolder(folder, fileName)
  // nullの場合はファイルがないので新規作成する
  if (file !== null) {
    file.setContent(jsonStr)
  } else {
    folder.createFile(fileName, jsonStr, MimeType.PLAIN_TEXT)
  }
  console.log("finish!")
}

const tableToJson = sheet => {
  const json_data = []
  const lastRow = sheet.getLastRow()
  const lastColumn = sheet.getLastColumn()
  const rows = sheet.getRange(1, 1, lastRow, lastColumn).getValues()
  console.log(rows)
  const keys = rows.shift()
  for (const row of rows) {
    const tmpObj = {}
    for (let i = 0; i < row.length; i++) {
      tmpObj[keys[i]] = row[i]
    }
    json_data.push(tmpObj)
  }
  return json_data
}

const getFileFromFolder = (folder, fileName) => {
  const files = folder.getFilesByName(fileName)
  while (files.hasNext()) {
    const file = files.next();
    if(fileName === file.getName()) {
      return file
    }
  }
  return null
}
