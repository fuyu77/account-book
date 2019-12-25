import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet
import Sheet = GoogleAppsScript.Spreadsheet.Sheet
import Range = GoogleAppsScript.Spreadsheet.Range

export default class AccountBook {
  protected readonly DATE = 0
  protected readonly NAME = 1
  protected readonly PRICE = 2
  protected readonly MONTH_TOTAL = 3
  protected readonly PREVIOUS_AMOUNT = 4
  protected readonly TOTAL  = 5
  protected readonly ADJUSTED_AMOUNT = 6
  protected readonly NEXT_AMOUNT = 7
  protected readonly DONE = 8
  protected readonly LAST_COLUMN = this.DONE

  private spreadSheet: Spreadsheet
  private sheet: Sheet
  private range: Range
  private lastRow: number
  protected values: any[][]

  public constructor() {
    this.spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
    this.sheet = SpreadsheetApp.getActiveSheet()
    this.range = this.sheet.getRange(2, 1, this.sheet.getLastRow() - 1, this.LAST_COLUMN + 1)
    this.values = this.range.getValues()
    this.lastRow = this.values.length - 1
  }

  public isDone(): boolean {
    return !!this.values[0][this.DONE]
  }

  public setDate(): void {
    if (this.values[this.lastRow][this.DATE] === "") {
      this.values[this.lastRow][this.DATE] = Utilities.formatDate(new Date(), "Asia/Tokyo", "M/d")
    }
  }

  public setMonthTotal(): void {
    this.values[0][this.MONTH_TOTAL] = 0
    this.values.forEach((record): void => this.values[0][this.MONTH_TOTAL] += Number(record[this.PRICE]))
  }

  public setTotal(): void {
    this.values[0][this.TOTAL] = this.values[0][this.MONTH_TOTAL] + this.values[0][this.PREVIOUS_AMOUNT]
  }

  public setAdjustedAmount(): void {
    this.values[0][this.ADJUSTED_AMOUNT] = this.values[0][this.TOTAL] > 0 ? Math.floor(this.values[0][this.TOTAL] / 10000) * 10000 : 0
  }

  public setNextAmount(): void {
    this.values[0][this.NEXT_AMOUNT] = this.values[0][this.TOTAL] - this.values[0][this.ADJUSTED_AMOUNT]
  }

  public setValues(): void {
    this.range.setValues(this.values)
  }

  public duplicateSheet(): void {
    const date = new Date()
    const sheetName = this.sheet.getName()
    if (sheetName === Utilities.formatDate(date, "Asia/Tokyo", "yyyy年M月")) {
      const nextMonth = new Date(date.getFullYear(), date.getMonth() + 1)
      this.spreadSheet.duplicateActiveSheet().setName(Utilities.formatDate(nextMonth, "Asia/Tokyo", "yyyy年M月"))
    } else {
      this.spreadSheet.duplicateActiveSheet().setName(Utilities.formatDate(date, "Asia/Tokyo", "yyyy年M月"))
    }
  }

  public moveActiveSheet(): void {
    this.spreadSheet.moveActiveSheet(0)
  }

  public getNextAmount(): number {
    return this.values[0][this.NEXT_AMOUNT]
  }
}
