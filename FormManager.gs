function ManageLibrary(){

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = "";
  error.employeeName = "";
  error.employeeNumber = "";
  error.formAnswer1 = "";
  error.formAnswer2 = "";
  error.where = "ManageLibrary(FormManager)";

  try {
    const SS = SpreadsheetApp.openById("1yNNGqzBplAVxBqNMa6D_0eskQzMIy89NPrE1uGlYbfs");
  }
  catch (e) {
    error.what = "スプレッドシート「図書貸出管理」のIDが間違っています";
    InsertError(error);
    return;
  }
  const SS = SpreadsheetApp.openById("1yNNGqzBplAVxBqNMa6D_0eskQzMIy89NPrE1uGlYbfs");

  const TriggerSS = SpreadsheetApp.getActiveSpreadsheet();
  const SHEETS = TriggerSS.getSheets();
  let timestamp = [];
  let sortedTimestamp = [];
  let bookData = {};

  //それぞれのシートの一番新しいタイムスタンプを取得
  for (let i = 0; i < SHEETS.length; i++){
    if (SHEETS[i].getLastRow() == 1){
      timestamp[i] = 0;
    } else {
      timestamp[i] = SHEETS[i].getRange(SHEETS[i].getLastRow(), 1).getCell(1,1).getValue();
      sortedTimestamp[i] = SHEETS[i].getRange(SHEETS[i].getLastRow(), 1).getCell(1,1).getValue();
 
      if (SHEETS[i].getRange(SHEETS[i].getLastRow(), 1).getCell(1,1).getValue() == ""){
        error.what = "シート「" + SHEETS[i].getName() +"」の最終行" + SHEETS[i].getLastRow() +"行目にタイムスタンプがありません";
        InsertError(error);
        return;
      }
  　}
  }
  sortedTimestamp.sort(function(a, b) {return b - a;});

  //一番新しいタイムスタンプの本を探す
  for (let i = 0; i < SHEETS.length; i++){
    if (sortedTimestamp[0].toString() == timestamp[i].toString()){
      var triggerSheet = SHEETS[i]; 
      bookData.sheetName = triggerSheet.getName();
      var sheetNameSplit = triggerSheet.getName().split("-");
      bookData.bookNumber = sheetNameSplit[0];
    }
  }

  if (bookData.sheetName.indexOf("貸出")　>= 0){
    BorrowBook(bookData, SS);
  } else if(bookData.sheetName.indexOf("返却")　>= 0){
    BackBook(bookData, SS);
  }
}



function CreateNewForm() {
  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.employeeName = "";
  error.employeeNumber = "";
  error.formAnswer1 = "";
  error.formAnswer2 = "";

  try {
    const SS = SpreadsheetApp.openById("1yNNGqzBplAVxBqNMa6D_0eskQzMIy89NPrE1uGlYbfs");
  }
  catch (e) {
    // Logger.log("error");
    error.book = "";
    error.where = "CreateNewForm(FormManager)";
    error.what = "スプレッドシート「図書貸出管理」のIDが間違っています";
    InsertError(error);
    return;
  }
  const SS = SpreadsheetApp.openById("1yNNGqzBplAVxBqNMa6D_0eskQzMIy89NPrE1uGlYbfs");
  
  const STATUS_SHEET = SS.getSheetByName("貸出状況");
  if (STATUS_SHEET == null){
    error.book = "";
    error.where = "CreateNewForm(FormManager)";
    error.what = "スプレッドシート「図書貸出管理」内，「貸出状況」シートの名前が間違っています";
    InsertError(error);
    return;
  }

  let range = STATUS_SHEET.getRange("A:G");
  let lastRow = STATUS_SHEET.getLastRow();

  //貸出状況シートから、一番下の書籍番号を取得
  let bookNumber = range.getCell(lastRow, 1).getValue();

  if (bookNumber == ""){
    let error = {};
    error.book = bookNumber;
    error.where = "CreateNewForm(FormManager)";
    error.what = "書籍番号がありません";
    InsertError(error);
    return;
  }

  let bookTitle = range.getCell(lastRow, 2).getValue();
  if (bookTitle == ""){
    let error = {};
    error.book = bookNumber;
    error.where = "CreateNewForm(FormManager)";
    error.what = "タイトルがありません";
    InsertError(error);
    return;
  }

  //貸出履歴シートの作成
  SS.insertSheet();
  SS.getActiveSheet().setName(bookNumber);
  SS.moveActiveSheet(SS.getNumSheets()); //新しい貸出履歴シートを最後尾に移動

  let logSheet = SS.getActiveSheet()
  logSheet.getRange(1, 1).getCell(1, 1).setValue("bookTitle");
  logSheet.getRange(2, 1).getCell(1, 1).setValue(bookTitle);
  logSheet.getRange(1, 2).getCell(1, 1).setValue("employeeName");
  logSheet.getRange(1, 3).getCell(1, 1).setValue("employeeNumber");
  logSheet.getRange(1, 4).getCell(1, 1).setValue("borrowDate");
  logSheet.getRange(1, 5).getCell(1, 1).setValue("backDeadline");
  logSheet.getRange(1, 6).getCell(1, 1).setValue("backDate");
　SS.setFrozenRows(1);


  //貸出フォームの作成
  let borrowFormTitle = bookNumber + "-『" + bookTitle + "』の貸出";

  let borrowForm = FormApp.create(borrowFormTitle);
  let borrowFormId = borrowForm.getId();
  let borrowFormFile = DriveApp.getFileById(borrowFormId);

  // borrowForm.setDescription();
  borrowForm.addTextItem().setTitle("お名前").setRequired(true);
  const validation = FormApp.createTextValidation().requireNumber().build();//社員番号を数字のみ入力可に
  borrowForm.addTextItem().setTitle("社員番号").setRequired(true).setValidation(validation);
  borrowForm.addDateItem().setTitle('貸出日').setRequired(true);
  borrowForm.addDateItem().setTitle('返却日').setRequired(true);

  //貸出フォームをフォームフォルダへ移動
  try {
    DriveApp.getFolderById("1Wcv9gLhsLTftWAbnTaqrE5NTKShw0E-V").addFile(borrowFormFile);
    DriveApp.getRootFolder().removeFile(borrowFormFile);
  }
  catch (e) {
    error.book = bookNumber　+"-貸出";
    error.where = "CreateNewForm(FormManager)";
    error.what = "フォームフォルダのIDが間違っています";
    InsertError(error);
    return;
  }

  //貸出フォームIDを「貸出状況」シートに追加
  range.getCell(lastRow, 7).setValue(borrowFormId);


  //返却フォームの作成
  let backFormTitle = bookNumber + "-『" + bookTitle + "』の返却";

  let backForm = FormApp.create(backFormTitle);
  let backFormId = backForm.getId();
  let backFormFile = DriveApp.getFileById(backFormId);
 
  // backForm.setDescription();
  backForm.addTextItem().setTitle("お名前").setRequired(true);
  backForm.addTextItem().setTitle("社員番号").setRequired(true).setValidation(validation);//社員番号を数字のみ入力可に
  backForm.addDateItem().setTitle('返却日').setRequired(true);

  //返却フォームをフォームフォルダへ移動
  try {
    DriveApp.getFolderById("1Wcv9gLhsLTftWAbnTaqrE5NTKShw0E-V").addFile(backFormFile);
    DriveApp.getRootFolder().removeFile(backFormFile);
  }
  catch (e) {
    error.book = bookNumber　+"-返却";
    error.where = "CreateNewForm(FormManager)";
    error.what = "フォームフォルダのIDが間違っています";
    InsertError(error);
    return;
  }
  DriveApp.getRootFolder().removeFile(backFormFile);

  //貸出フォームとシートを紐づけ
  const TRIGGER_SS = SpreadsheetApp.getActiveSpreadsheet();

  borrowForm.setDestination(FormApp.DestinationType.SPREADSHEET, TRIGGER_SS.getId());

  //紐づけされたシートの名前変更
  var triggerSheets = TRIGGER_SS.getSheets();
  for (let i = 0; i < triggerSheets.length; i++) {
    if (triggerSheets[i].getName() == bookNumber +"-貸出"){
      // Logger.log("in「5-貸出」は既に存在しています")
      error.book = bookNumber +"-貸出";
      error.where = "CreateNewForm(FormManager)";
      error.what = "フォームと紐づけられた「" + bookNumber + "-貸出」シートは既に存在しています。";
      InsertError(error);
      return;
    }
    if (triggerSheets[i].getName() == bookNumber +"-返却"){
      // Logger.log("in「5-返却」は既に存在しています")
      error.book = bookNumber +"-返却";
      error.where = "CreateNewForm(FormManager)";
      error.what = "フォームと紐づけられた「" + bookNumber + "-返却」シートは既に存在しています。";
      InsertError(error);
      return;
    }
  }

  let flag = 0;
  for (let i = 0; i < triggerSheets.length; i++) {
    
    if (triggerSheets[i].getName().indexOf("フォームの回答") >= 0) {
      if (flag > 0){
        error.book = bookNumber +"-貸出";
        error.where = "CreateNewForm(FormManager)";
        error.what = "（貸出シートを紐づけ）新しいシートが２枚以上あります";
        InsertError(error);
        break;
      }
      triggerSheets[i].setName(bookNumber + "-貸出");
      flag++;
    }
  }
  if (flag == 0){
    error.book = bookNumber +"-貸出";
    error.where = "CreateNewForm(FormManager)";
    error.what = "（貸出シートを紐づけ）新しいシートがありません";
    InsertError(error);
  }

  //貸出フォームとシートを紐づけ
  backForm.setDestination(FormApp.DestinationType.SPREADSHEET, TRIGGER_SS.getId());

  //紐づけされたシートの名前変更
  var triggerSheets = TRIGGER_SS.getSheets();

  flag = 0;
  for (let i = 0; i < triggerSheets.length; i++) {
    if (triggerSheets[i].getName().indexOf("フォームの回答") >= 0) {
      if (flag > 0){
        error.book = bookNumber +"-返却";
        error.where = "CreateNewForm(FormManager)";
        error.what = "（返却シートを紐づけ）新しいシートが２枚以上あります";
        InsertError(error);
        break;
      }
      triggerSheets[i].setName(bookNumber + "-返却");
      flag++;
    }
  }
  if (flag == 0){
    error.book = bookNumber +"-返却";
    error.where = "CreateNewForm(FormManager)";
    error.what = "（返却シートを紐づけ）新しいシートがありません（返却シートの名前が変更できませんでした）";
    InsertError(error);
  }
}
