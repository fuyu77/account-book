import AccountBook from "./AccountBook"

export default class NextAccountBook extends AccountBook {
  clearValues() {
    this.values = this.values.map(record => record.map(value => ""))
  }

  setPreviousAmount(previousAmount: number) {
    this.values[0][this.PREVIOUS_AMOUNT] = previousAmount
  }
}
