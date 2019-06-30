import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet
import Sheet = GoogleAppsScript.Spreadsheet.Sheet
import Range = GoogleAppsScript.Spreadsheet.Range

export default class AccountBook {
  readonly DATE = 0
  readonly NAME = 1
  readonly PRICE = 2
  readonly MONTH_TOTAL = 3
  readonly PREVIOUS_AMOUNT = 4
  readonly TOTAL  = 5
  readonly ADJUSTED_AMOUNT = 6
  readonly NEXT_AMOUNT = 7
  readonly DONE = 8
  readonly LAST_COLUMN = this.DONE

  private spreadSheet: Spreadsheet
  private sheet: Sheet
  private range: Range
  private lastRow: number
  protected values: any[][]

  constructor() {
    this.spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
    this.sheet = SpreadsheetApp.getActiveSheet()
    this.lastRow = this.sheet.getLastRow()
    this.range = this.sheet.getRange(2, 1, this.lastRow - 1, this.LAST_COLUMN + 1)
    this.values = this.range.getValues()
  }

  isDone() {
    return !!this.values[0][this.DONE]
  }

  setDate() {
    if (this.values[this.lastRow - 2][this.DATE] === "") {
      this.values[this.lastRow - 2][this.DATE] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d')
    }
  }

  setMonthTotal() {
    this.values[0][this.MONTH_TOTAL] = 0
    this.values.forEach(record => {
      this.values[0][this.MONTH_TOTAL] += Number(record[this.PRICE]) 
    })
  }

  setTotal() {
    this.values[0][this.TOTAL] = this.values[0][this.MONTH_TOTAL] + this.values[0][this.PREVIOUS_AMOUNT]
  }

  setAdjustedAmount() {
    this.values[0][this.ADJUSTED_AMOUNT] = Math.floor(this.values[0][this.TOTAL] / 10000) * 10000
  }

  setNextAmount() {
    this.values[0][this.NEXT_AMOUNT] = this.values[0][this.TOTAL] - this.values[0][this.ADJUSTED_AMOUNT]
  }

  setValues() {
    this.range.setValues(this.values)
  }

  duplicateSheet() {
    const date = new Date()
    const sheetName = this.sheet.getName()
    if (sheetName === Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年M月')) {
      const nextMonth = new Date(date.getFullYear(), date.getMonth() + 1)
      this.spreadSheet.duplicateActiveSheet().setName(Utilities.formatDate(nextMonth, 'Asia/Tokyo', 'yyyy年M月'))
    } else {
      this.spreadSheet.duplicateActiveSheet().setName(Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年M月'))
    }
  }

  moveActiveSheet() {
    this.spreadSheet.moveActiveSheet(0)
  }

  getNextAmount() {
    return this.values[0][this.NEXT_AMOUNT]
  }
}
