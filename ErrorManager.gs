function InsertError(error) {
  
  const SS = SpreadsheetApp.openById("1d-DK2eNTH6iUVlj_kyNE6lvSp20eQiIR1ydu-6lf9RA");
  const ERROR_SHEET = SS.getSheetByName("エラー用");
  let lastRow = ERROR_SHEET.getLastRow();

  let range = ERROR_SHEET.getRange(lastRow + 1, 1, 1, 9);
  range.getCell(1,1).setValue("未");
  range.getCell(1,2).setValue(error.timestamp);
  range.getCell(1,3).setValue(error.book);
  range.getCell(1,4).setValue(error.employeeName);
  range.getCell(1,5).setValue(error.employeeNumber);
  range.getCell(1,6).setValue(error.formAnswer1);
  range.getCell(1,7).setValue(error.formAnswer2);
  range.getCell(1,8).setValue(error.where);
  range.getCell(1,9).setValue(error.what);

  MailApp.sendEmail("shiorit.53e@gmail.com", "【図書貸出管理システム】エラーのご報告" ,
                        "\n図書貸出管理システムにエラーがありました。\nエラー用シート" + lastRow +"行目をご確認ください。"
                        + "\n\n---------------\n"
                        + "エラー内容：" + error.where 
                        + "\n　　　　　　" + error.what); 


  return;
}
