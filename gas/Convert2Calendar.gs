// 作成日:2021/4/24
// 作成者:Tatsuya
// URL:https://kogyo-kyoiku.net

function Convert2Calender() {
  // 現在開いているシートを取得する
  const ActiveSheet = SpreadsheetApp.getActiveSheet();
  // 現在開いているシートの月を取得する
  const ActiveMonth = ActiveSheet.getName().replace("月","");
  // 現在開いているシートのデータを二次元配列で取得する
  let SheetData = ActiveSheet.getDataRange().getValues();
  // 年度（1行A列のデータ）を取得する
  let TargetYear = SheetData[0][0];
  // デフォルトのカレンダーを取得する
  let Calendar = CalendarApp.getDefaultCalendar();
  // カレンダーのタイムゾーンの設定
  Calendar.setTimeZone("Asia/Tokyo");
  // 日程情報を検索していく処理
  for(let i = 1 ; i < SheetData.length ; i++){
    for(let j = 2 ; j < SheetData[i].length ; j+=3){
      if(SheetData[i][j] != ""){
        // 予定があった場合に、開始時間を参照して予定を作成する
        Calendar.createEvent(SheetData[i][j],
                              new Date(TargetYear+"/"+ActiveMonth+"/"+SheetData[i][0]+" "+SheetData[i][j+1]),
                              new Date(TargetYear+"/"+ActiveMonth+"/"+SheetData[i][0]+" "+SheetData[i][j+2]));
      }
    }
  }
  // 処理完了表示
  Browser.msgBox("日程をカレンダーに出力しました。");
}