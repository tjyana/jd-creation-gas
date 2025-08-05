function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('選択した回答をDocsに書き出す', 'exportToDocs')
    .addToUi();
}

function exportToDocs() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // 保存先フォルダのIDを指定
    const folderId = '1E9x3UB3To5zzcresi6b5MgbkO7LBC-ZU';
    const folder = DriveApp.getFolderById(folderId);
    
    // 選択された回答があるか確認
    const hasSelectedResponses = data.slice(1).some(row => row[1] === true);
    if (!hasSelectedResponses) {
      SpreadsheetApp.getUi().alert('選択された回答がありません。\nチェックボックスで回答を選択してください。');
      return;
    }

    // ドキュメントを作成
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    const doc = DocumentApp.create(`選択した回答_${timestamp}`);
    const body = doc.getBody();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === true) { // B列のチェックボックスがチェックされているか確認
        body.appendParagraph(`回答 ${i}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
        for (let j = 2; j < data[i].length; j++) { // 質問の内容をすべて書き出す
          const content = data[i][j].toString().trim();
          if (content) { // 空の回答は除外
            const questionPara = body.appendParagraph(`${headers[j]}:`);
            questionPara.setBold(true);
            body.appendParagraph(content).setIndentStart(20).setBold(false); // 回答部分は太字解除
          }
        }
        body.appendParagraph(''); // 空行を追加
      }
    }
    doc.saveAndClose();
    
    // 作成したドキュメントを指定フォルダに移動
    const docFile = DriveApp.getFileById(doc.getId());
    folder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile);
    
    // 生成されたドキュメントのURLを取得
    const url = doc.getUrl();
    SpreadsheetApp.getUi().alert('選択した回答を新しいドキュメントに書き出しました。\n\nドキュメントのURL: ' + url);

  } catch (error) {
    console.error('エラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました。\n' + error.toString());
  }
}