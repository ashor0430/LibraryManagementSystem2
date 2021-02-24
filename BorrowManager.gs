function BorrowBook(bookData, SS){

  let answers = GetBorrowData(bookData);
  if (answers == null){
    return;
  }

  InsertBorrowLogData(answers, SS);

  ResisterStatus(answers, SS);

  UpdateFormByBorrow(answers, SS);
}

function GetBorrowData(bookData){

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = bookData.bookNumber +"-貸出";
  error.where = "GetBorrowData(BorrowManager)";

  const TriggerSS = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = TriggerSS.getSheetByName(bookData.sheetName);

  let lastRow = sheet.getLastRow();
  let range = sheet.getRange(lastRow, 2, 1, sheet.getLastColumn());

  //回答の場所を探す
  let col = 1;
  while (range.getCell(1, col).isBlank()){
    if (col >= sheet.getLastColumn()){
      error.employeeName = "";
      error.employeeNumber = "";
      error.formAnswer1 = "";
      error.formAnswer2 = "";
      error.what = "フォームの回答がありません（トリガーシート" + bookData.sheetName + "，"
      　　　　　　　　 + lastRow + "行目のタイムスタンプ）";
      InsertError(error);
      return;
    }
    col++
  }

  let answers = {};
  answers.bookNumber = bookData.bookNumber;
  answers.employeeName = range.getCell(1, col).getValue();
  answers.employeeNumber = range.getCell(1, col + 1).getValue();
  answers.borrowDate = range.getCell(1, col + 2).getValue();
  answers.backDeadline = range.getCell(1, col + 3).getValue();

  if (answers.employeeName == null || answers.employeeName == "" ||
      answers.employeeNumber == null || answers.employeeNumber == "" ||
      answers.borrowDate == null || answers.borrowDate == "" ||
      answers.backDeadline == null || answers.backDeadline == ""){
    error.employeeName = answers.employeeName;
    error.employeeNumber = answers.employeeNumber;
    error.formAnswer1 = answers.borrowDate;
    error.formAnswer2 = answers.backDeadline;
    error.what = "フォームの回答の取得に失敗しました（トリガーシート" + bookData.sheetName + "，"
    　　　　　　　　 + lastRow + "行目のタイムスタンプ，フォームの回答" + col + "列目～）";
    InsertError(error);
    return;
  }

  return answers;
}

function InsertBorrowLogData(answers, SS){

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = answers.bookNumber +"-貸出";
  error.employeeName = answers.employeeName;
  error.employeeNumber = answers.employeeNumber;
  error.formAnswer1 = answers.borrowDate;
  error.formAnswer2 = answers.backDeadline;
  error.where = "InsertBorrowLogData(BorrowManager)";

  let sheet = SS.getSheetByName(answers.bookNumber);
  if (sheet == null || sheet == ""){
    error.what = "貸出履歴シート「" + answers.bookNumber + "」の取得に失敗しました";
    InsertError(error);
    return;
  }

  let range = sheet.getRange("B:E")
  let lastRow = sheet.getLastRow();
  range.getCell(lastRow +1, 1).setValue(answers.employeeName);
  range.getCell(lastRow +1, 2).setValue(answers.employeeNumber);
  range.getCell(lastRow +1, 3).setValue(answers.borrowDate);
  range.getCell(lastRow +1, 4).setValue(answers.backDeadline);
  
}

function ResisterStatus(answers, SS){

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = answers.bookNumber +"-貸出";
  error.employeeName = answers.employeeName;
  error.employeeNumber = answers.employeeNumber;
  error.formAnswer1 = answers.borrowDate;
  error.formAnswer2 = answers.backDeadline;
  error.where = "ResisterStatus(BorrowManager)";
  
  const STATUS_SHEET = SS.getSheetByName("貸出状況");
  if (STATUS_SHEET == null || STATUS_SHEET == ""){
    error.what = "スプレッドシート「図書貸出管理」内，「貸出状況」シートの名前が間違っています";
    InsertError(error);
    return;
  }

  let range = STATUS_SHEET.getRange("A:G");
  let lastRow = STATUS_SHEET.getLastRow();

  let flag = 0;
  for (let i = 2; i <= lastRow; i++){
    if (range.getCell(i, 1).getValue() == answers.bookNumber){
      if (flag > 0){
        error.what = "「貸出状況」シートから書籍番号" + answers.bookNumber + "が２か所以上見つかりました";
        InsertError(error);
        return;
      }
      range.getCell(i, 3).setValue(answers.employeeName);
      range.getCell(i, 4).setValue(answers.employeeNumber);
      range.getCell(i, 5).setValue(answers.borrowDate);
      range.getCell(i, 6).setValue(answers.backDeadline);
      flag++;
    }
  }
  if (flag == 0){
    error.what = "「貸出状況」シートから書籍番号が見つかりませんでした";
    InsertError(error);
    return;
  }
}

function UpdateFormByBorrow(answers, SS){

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = answers.bookNumber +"-貸出";
  error.employeeName = answers.employeeName;
  error.employeeNumber = answers.employeeNumber;
  error.formAnswer1 = answers.borrowDate;
  error.formAnswer2 = answers.backDeadline;
  error.where = "UpdateFormByBorrow(BorrowManager)";

  const STATUS_SHEET = SS.getSheetByName("貸出状況");
  if (STATUS_SHEET == null || STATUS_SHEET == ""){
    error.what = "スプレッドシート「図書貸出管理」内，「貸出状況」シートの名前が間違っています";
    InsertError(error);
    return;
  }

  let range = STATUS_SHEET.getRange("A:G");
  let lastRow = STATUS_SHEET.getLastRow();

  if (answers.bookNumber == null || answers.bookNumber == "" 
      || answers.backDeadline == null || answers.backDeadline ==""){
    error.what = "answersが取得できませんでした";
    InsertError(error);
    return;
  }
  answers.backDeadline = Utilities.formatDate(answers.backDeadline,"JST", "yyyy/MM/dd");

  let flag = 0;
  for (let i = 2; i <= lastRow; i++){
    if (range.getCell(i, 1).getValue() == answers.bookNumber){
      if (flag > 0){
        error.what = "「貸出状況」シートから書籍番号" + answers.bookNumber + "が２か所以上見つかりました";
        InsertError(error);
        return;
      }
      var formId = range.getCell(i, 7).getValue();
      flag++;
    }
  }
  if (flag == 0){
    error.what = "「貸出状況」シートから書籍番号が見つかりませんでした";
    InsertError(error);
    return;
  }
  if (formId == null || formId == ""){
    error.what = "「貸出状況」シートにフォームIDがありません";
    InsertError(error);
    return;
  }
  
  try {
    var form = FormApp.openById(formId);
  }
  catch(e){
    error.what = "「貸出状況」シートのフォームIDが間違っています";
    InsertError(error);
    return;
  }
 
  let items = form.getItems();  
  for (let i = 0; i < items.length; i++){
    form.deleteItem(items[i]);
  }
  form.setDescription("貸出中につき現在借りられません。しばらくお待ちください。 \n返却予定日：" + answers.backDeadline);
}
