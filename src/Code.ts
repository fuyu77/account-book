// スプレッドシートに編集があった際に動作する。
function onEdit(event) {
  const DATE = 1
  const EXPENSE = 3
  const MONTH_TOTAL = 4
  const PREVIOUS_AMOUNT = 5
  const TOTAL = 6
  const ADJUSTED_AMOUNT = 7 
  const NEXT_AMOUNT = 8
  const DONE = 9
  const LAST_COLUMN = DONE
　 
  // 現在開いているシートを取得する。
  const sheet = SpreadsheetApp.getActiveSheet()
  
  // 値の存在する最終行を取得する。
  const lastRow = sheet.getLastRow()
  
  const amounts = sheet.getRange(2, MONTH_TOTAL, 1, LAST_COLUMN - MONTH_TOTAL + 1).getValues()
  
  if (amounts[0][DONE - MONTH_TOTAL]) {
    nextMonth()
    sheet.getRange(2, DONE).clear()
    return
  }
  let date = sheet.getRange(lastRow, DATE).getValue()
  const expenses = sheet.getRange(2, EXPENSE, lastRow - 1).getValues()
  
  if (date === "") {
      const d = new Date()
      date = Utilities.formatDate(d, 'Asia/Tokyo', 'M/d')
   }
 
  const sum = expenses.reduce((a, b) => {
    return [Number(a) + Number(b)]
  })
  
  amounts[0][0] = sum
  amounts[0][TOTAL - MONTH_TOTAL] = Number(sum) + amounts[0][PREVIOUS_AMOUNT - MONTH_TOTAL]
  amounts[0][ADJUSTED_AMOUNT - MONTH_TOTAL] = Math.floor(amounts[0][TOTAL - MONTH_TOTAL] / 10000) * 10000
  amounts[0][NEXT_AMOUNT - MONTH_TOTAL] = amounts[0][TOTAL - MONTH_TOTAL] - amounts[0][ADJUSTED_AMOUNT - MONTH_TOTAL]
  
  sheet.getRange(lastRow, DATE).setValue(date)
  sheet.getRange(2, MONTH_TOTAL, 1, LAST_COLUMN - MONTH_TOTAL + 1).setValues(amounts)
}

function nextMonth() {
  const DATEW = 1
  const EXPENSE = 3
  const MONTH_TOTAL = 4
  const PREVIOUS_AMOUNT = 5
  const TOTAL = 6
  const ADJUSTED_AMOUNT = 7 
  const NEXT_AMOUNT = 8
  const DONE = 9
  const LAST_COLUMN = DONE
  
  var date = new Date()
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var name = ss.getSheets()[0].getName().replace('年', '').replace('月', '')
  if (name === Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyM')) {
    var nextDate = new Date(date.getFullYear(), date.getMonth() + 1)
    ss.duplicateActiveSheet().setName(Utilities.formatDate(nextDate, 'Asia/Tokyo', 'yyyy年M月'))
  } else {
    ss.duplicateActiveSheet().setName(Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年M月'))
  }
  ss.moveActiveSheet(0)
  var sheet = SpreadsheetApp.getActiveSheet()
  var lastRow = sheet.getLastRow()
  var nextAmount = sheet.getRange(2, NEXT_AMOUNT).getValue()
  sheet.getRange(2, 1, lastRow - 1, LAST_COLUMN).clear()
  sheet.getRange(2, PREVIOUS_AMOUNT).setValue(nextAmount)
}
