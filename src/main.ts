import AccountBook from "./AccountBook"
import NextAccountBook from "./NextAccountBook"

const createNextSheet = (accountBook: AccountBook): void => {
  accountBook.duplicateSheet()
  accountBook.moveActiveSheet()
  const nextAccountBook = new NextAccountBook()
  nextAccountBook.clearValues()
  nextAccountBook.setPreviousAmount(accountBook.getNextAmount())
  nextAccountBook.setValues()
}

function onEdit(): void {
  const accountBook = new AccountBook()
  if (accountBook.isDone()) {
    createNextSheet(accountBook)
    return
  }
  accountBook.setDate()
  accountBook.setMonthTotal()
  accountBook.setTotal()
  accountBook.setAdjustedAmount()
  accountBook.setNextAmount()
  accountBook.setValues()
}
