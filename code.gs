// ************************************
// メニューの追加
// ************************************
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('カレンダー作成')
     .addItem('カレンダーへ適用', 'createCal')
     .addToUi();

}

// ************************************
//
// ************************************
function createCal() {

  var calendar = CalendarApp.getCalendarById('カレンダーID');

  var spreadsheet = SpreadsheetApp.getActive();

  var sheets = spreadsheet.getSheets();

  for( j = 1; j < sheets.length; j++ ) {

    var target1 = "";
    var target2 = "";
    var target3 = "";
    var target4 = "";
    var target5 = "";
    var target6 = "";

    var targetRange = null;

    targetRange = sheets[j].getRange('C1');
    var year = targetRange.getValue();
    targetRange = sheets[j].getRange('F1');
    var month = targetRange.getValue();

    for(var i = 2; i <= 100; i++ ) {

      // 日付
      targetRange = sheets[j].getRange('A' + i);
      target1 = targetRange.getValue();

      if ( target1 == "" ) {
        break;
      }

      // タイトル
      targetRange = sheets[j].getRange('C' + i);
      target2 = targetRange.getValue();

      // タイトルが空白の場合のカレンダーイベントの削除
      if ( target2 == "" ) {
        // ID の取得
        targetRange = sheets[j].getRange('G' + i);
        target6 = targetRange.getValue();
        if ( target6 != "" ) {
          try {
            targetRange.setValue("");
            targetRange = sheets[j].getRange('F' + i);
            targetRange.setValue("");
            var curEvent = calendar.getEventById(target6);
            curEvent.deleteEvent();
          }
          catch(e){
          }
        }
        continue;
      }

      // 開始時間( 現在未使用 )
      targetRange = sheets[j].getRange('D' + i);
      target3 = targetRange.getValue();

      // 終了時間( 現在未使用 )
      targetRange = sheets[j].getRange('E' + i);
      target4 = targetRange.getValue();

      // 備考
      targetRange = sheets[j].getRange('F' + i);
      target5 = targetRange.getValue();

      // id
      targetRange = sheets[j].getRange('G' + i);
      target6 = targetRange.getValue();

      var startTime = new Date(year + '/' + month + '/' + target1);
      var endTime = startTime;

      // 新規登録
      if ( target6 + "" == "" ) {
          var curEvent = calendar.createAllDayEvent(target2, startTime);
          curEvent.setDescription(target5);
          targetRange.setValue(curEvent.getId());
      }
      // 既存修正
      else {
        var iCalId = target6;
        var curEvent = calendar.getEventById(iCalId);
        if ( curEvent.getTitle() != target2 ) {
          curEvent.setTitle(target2);
        }
        if ( curEvent.getDescription() != target5 ) {
          curEvent.setDescription(target5);
        }
      }
    }
  }
}
