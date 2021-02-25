function BackBook(bookData, SS){
  
  let answers = GetBackData(bookData);
    if (answers == null){
    return;
  }

  InsertBackLogData(answers, SS);

  ResetStatus(answers, SS);

  UpdateFormByBack(answers, SS);
}

function GetBackData(bookData){

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = bookData.bookNumber +"-返却";
  error.where = "GetBackData(BorrowManager)";

  const TriggerSS = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = TriggerSS.getSheetByName(bookData.sheetName);

  let lastRow = sheet.getLastRow();
  let range = sheet.getRange("A:D");

  let answers = {};
  answers.bookNumber = range.getCell(lastRow, 2).getValue();
  answers.employeeName = range.getCell(lastRow, 3).getValue();
  answers.employeeNumber = range.getCell(lastRow, 4).getValue();
  answers.backDate = range.getCell(lastRow, 1).getValue();

  if (answers.employeeName == null || answers.employeeName == "" ||
      answers.employeeNumber == null || answers.employeeNumber == "" ||
      answers.backDate == null || answers.backDate == ""){
    error.employeeName = answers.employeeName;
    error.employeeNumber = answers.employeeNumber;
    error.formAnswer1 = answers.backDate;
    error.formAnswer2 = "-";
    error.what = "フォームの回答の取得に失敗しました（トリガーシート" + bookData.sheetName + "，"
    　　　　　　　　 + lastRow + "行目のタイムスタンプ）";
    InsertError(error);
    return;
  }
  // Logger.log(answers);
  return answers;
}

function InsertBackLogData(answers, SS){

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = answers.bookNumber　+ "-返却";
  error.employeeName = answers.employeeName;
  error.employeeNumber = answers.employeeNumber;
  error.formAnswer1 = answers.backDate;
  error.formAnswer2 = "-";
  error.where = "InsertBackLogData(BackManager)";

  let sheet = SS.getSheetByName(answers.bookNumber);
  if (sheet == null || sheet == ""){
    error.what = "貸出履歴シート「" + answers.bookNumber + "」の取得に失敗しました";
    InsertError(error);
    return;
  }

  let range = sheet.getRange("B:F");
  let flag = 0;
  for (let row = 2; row <= sheet.getLastRow(); row++){
    if (range.getCell(row, 2).getValue() == answers.employeeNumber && range.getCell(row, 5).isBlank()){
      if (flag > 0){
        error.what = "こちらの社員番号による，返却のない貸出記録が２か所以上見つかりました";
        InsertError(error);
        return;
      }
      range.getCell(row, 5).setValue(answers.backDate);
      flag++;
    }
  }
  if (flag == 0){
    error.what = "こちらの社員番号による，返却のない貸出記録が見つかりませんでした";
    InsertError(error);
    return;
  }

}

function ResetStatus(answers, SS){

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = answers.bookNumber　+ "-返却";
  error.employeeName = answers.employeeName;
  error.employeeNumber = answers.employeeNumber;
  error.formAnswer1 = answers.backDate;
  error.formAnswer2 = "-";
  error.where = "ResetStatus(BackManager)";

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
      let cells = STATUS_SHEET.getRange(i, 3, 1, 4);
      cells.clear();
      flag++;
    }
  }
  if (flag == 0){
    error.what = "「貸出状況」シートから書籍番号が見つかりませんでした";
    InsertError(error);
    return;
  }
}

function UpdateFormByBack(answers, SS) {

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = answers.bookNumber　+ "-返却";
  error.employeeName = answers.employeeName;
  error.employeeNumber = answers.employeeNumber;
  error.formAnswer1 = answers.backDate;
  error.formAnswer2 = "-";
  error.where = "UpdateFormByBack(BackManager)";

  const STATUS_SHEET = SS.getSheetByName("貸出状況");
  if (STATUS_SHEET == null || STATUS_SHEET == ""){
    error.what = "スプレッドシート「図書貸出管理」内，「貸出状況」シートの名前が間違っています";
    InsertError(error);
    return;
  }

  let range = STATUS_SHEET.getRange("A:G");
  let lastRow = STATUS_SHEET.getLastRow();

  if (answers.bookNumber == null || answers.bookNumber == ""){
    error.what = "answersが取得できませんでした";
    InsertError(error);
    return;
  }

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
  form.setDescription("");
  form.addTextItem().setTitle("お名前").setRequired(true);
  const validation = FormApp.createTextValidation().requireNumber().build();
  form.addTextItem().setTitle("社員番号").setRequired(true).setValidation(validation);
  form.addDateItem().setTitle('貸出日').setRequired(true);
  form.addDateItem().setTitle('返却日').setRequired(true);
}
