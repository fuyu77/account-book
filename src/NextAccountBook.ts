import AccountBook from "./AccountBook"

export default class NextAccountBook extends AccountBook {
  public clearValues(): void {
    this.values = this.values.map((record): string[] => record.map((_value): string => ""))
  }

  public setPreviousAmount(previousAmount: number): void {
    this.values[0][this.PREVIOUS_AMOUNT] = previousAmount
  }
}
