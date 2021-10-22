// 作成日:2021/5/2
// 作成者:Tatsuya
// URL:https://kogyo-kyoiku.net

function Convert2Calender() {
  // カレンダー名をメッセージボックスから取得する
  let CalendarName = Browser.inputBox("登録したいカレンダーの名前を入力してください\\n（空白の場合はデフォルトカレンダーになります）");
  // 現在開いているシートを取得する
  const ActiveSheet = SpreadsheetApp.getActiveSheet();
  // 現在開いているシートの月を取得する
  const ActiveMonth = ActiveSheet.getName().replace("月","");
  // 現在開いているシートのデータを二次元配列で取得する
  let SheetData = ActiveSheet.getDataRange().getValues();
  // 年度（1行A列のデータ）を取得する
  let TargetYear = SheetData[0][0];
  // カレンダーオブジェクト格納変数を宣言
  let Calendar;
  if(CalendarName == ""){
    // カレンダー名を指定されていなければ、デフォルトカレンダーを取得
    Calendar = CalendarApp.getDefaultCalendar();
  }else{
    // カレンダー名を指定されていれば、そのカレンダーを取得
    Calendar = CalendarApp.getOwnedCalendarsByName(CalendarName)[0];
  }
  if(Calendar==null){
    // カレンダーが取得できなかった場合の処理
    Browser.msgBox("カレンダー名「"+CalendarName+"」が見つかりませんでした。\\n正しい名前を入力し直してください。\\n処理を終了します。");
    // スクリプトの終了処理
    return;
  }
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
  Browser.msgBox("カレンダー名「"+Calendar.getName()+"」に日程を出力しました。");
}