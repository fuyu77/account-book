import AccountBook from "./AccountBook"
import NextAccountBook from "./NextAccountBook"

function onEdit(e: Event) {
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

const createNextSheet = (accountBook: AccountBook) => {
  accountBook.duplicateSheet()
  accountBook.moveActiveSheet()
  const nextAccountBook = new NextAccountBook()
  nextAccountBook.clearValues()
  nextAccountBook.setPreviousAmount(accountBook.getNextAmount())
  nextAccountBook.setValues()
}
