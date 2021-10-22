'use strict'

//----------------------------------------------------------------------------
// チェック結果をクリア
function checkClear( ) {
  checkClear0( );
  checkClear1( );
}

//----------------------------------------------------------------------------
// チェック結果をクリア(列 A)
function checkClear0( ) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('カレンダー情報');

  // 見出しを再設定
  sheet.getRange( 'A1' ).setValue( "実行アカウント" );
  sheet.getRange( 'A4' ).setValue( "登録されているカレンダーの数" );

  // 可変の欄をクリア
  sheet.getRange( 'A2' ).setValue( "" );
  sheet.getRange( 'A5' ).setValue( "" );
}

//----------------------------------------------------------------------------
// チェック結果をクリア(列 B～I)
function checkClear1( ) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('カレンダー情報');
  
  // 一旦、範囲をすべて削除
  sheet.getRange( 'B:I' ).clearContent( );
  
  // 先頭行の見出しを再設定
  sheet.getRange( 'C1' ).setValue( "登録されているカレンダーの ID" );
  sheet.getRange( 'D1' ).setValue( "名称" );
  sheet.getRange( 'E1' ).setValue( "isHidden" );
  sheet.getRange( 'F1' ).setValue( "isMyPrimaryCalendar" );
  sheet.getRange( 'G1' ).setValue( "isOwnedByMe" );
  sheet.getRange( 'H1' ).setValue( "isSelected" );
  sheet.getRange( 'I1' ).setValue( "Timezone" );
  sheet.getRange( 'A4' ).setValue( "登録されているカレンダーの数" );

  sheet.getRange( 'A5' ).setValue( "" );
}

//----------------------------------------------------------------------------
// 現在のアカウントのカレンダー情報を取得
function getMyCalendars( ) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('カレンダー情報');

  // 現在の実行アカウントを確認    ※実行されているアカウントによって、動作が変わってくるので...
  var eMailID = Session.getActiveUser( ).getUserLoginId( );
  checkClear0( );
  sheet.getRange( 'A2' ).setValue( eMailID );
  
  // ユーザーが所有またはサブスクライブしているすべてのカレンダーを取得します。
  // Determines how many calendars the user can access.
  var calendars = CalendarApp.getAllCalendars( );
  checkClear1( );
  sheet.getRange( 'A5' ).setValue( calendars.length );     // カレンダーの総数
  
  for ( var i = 0 ; i < calendars.length ; i++ ) {
    sheet.getRange( i + 2, 2 ).setValue( i + 1 );
    sheet.getRange( i + 2, 3 ).setValue( calendars[ i ].getId( ) );                   // C
    sheet.getRange( i + 2, 4 ).setValue( calendars[ i ].getName( ) );                 // D
    sheet.getRange( i + 2, 5 ).setValue( calendars[ i ].isHidden( ) );                // E
    sheet.getRange( i + 2, 6 ).setValue( calendars[ i ].isMyPrimaryCalendar( ) );     // F
    sheet.getRange( i + 2, 7 ).setValue( calendars[ i ].isOwnedByMe( ) );             // G
    sheet.getRange( i + 2, 8 ).setValue( calendars[ i ].isSelected( ) );              // H
    sheet.getRange( i + 2, 9 ).setValue( calendars[ i ].getTimeZone() );              // I
  }  
}

//----------------------------------------------------------------------------
// 入力されているイベント情報をクリア
function eventClear( ) {
  var sheet = SpreadsheetApp.getActiveSheet( );
  var result = Browser.msgBox( "入力されているイベント情報をクリアしても構いませんか？\\n 【注意】 この操作は取り消せません！", Browser.Buttons.OK_CANCEL );
  if ( result == "ok" ) {
    sheet.getRange( 'C3:Z33' ).clearContent( );
  }
}

//----------------------------------------------------------------------------
// 作成日:2021/4/24
// 作成者:Tatsuya
// URL:https://kogyo-kyoiku.net    
//     https://kogyo-kyoiku.net/spreadsheet2calendar/?fbclid=IwAR3IdUnxCZ--_9woyNCH9H210zLE4pPwdwl1Im8TgFnFbHy6m-lo6mR1U2E
// 改変：T.Sakamoto 2021/05/01
function Convert2Calender( ) {
  var result = Browser.msgBox( "入力されているイベント情報で、イベントを出力しても構いませんか？\\n 【注意】 この操作は取り消せません！", Browser.Buttons.OK_CANCEL );
  if ( result != "ok" ) {
    // 以降の処理を行わない
    return;
  }
  
  // 現在開いているシートを取得する
  const ActiveSheet = SpreadsheetApp.getActiveSheet( );
  
  // 現在開いているシートのデータを二次元配列で取得する
  let SheetData = ActiveSheet.getRange( 'A1:Z33' ).getValues( );
  
  // 処理対象の月を取得する
  const ActiveMonth = SheetData[ 0 ][ 1 ];     // B1
  
  // 年度（1行A列のデータ）を取得する
  let TargetYear = SheetData[ 0 ][ 0 ];        // A1
  
  // デフォルトのカレンダーを取得する
  let Calendar = CalendarApp.getCalendarsByName( SheetData[ 0 ][ 5 ] );
  
  // カレンダーのタイムゾーンの設定
  Calendar[ 0 ].setTimeZone("Asia/Tokyo");

  // 日程情報を検索していく処理
  for( let i = 2 ; i < SheetData.length ; i++ ) {
    for( let j = 2 ; j < SheetData[ i ].length ; j += 3 ) {
      let sTitle = SheetData[ i ][ j ];
      if( sTitle != "" ) {
        // 予定に何か入力されていれば、イベント(予定)を作成する
        let startTime = SheetData[ i ][ j + 1 ];
        let endTime = SheetData[ i ][ j + 2 ];
        
        if ( ( startTime != "" ) && ( endTime != "" ) ) {
          // 開始時間・終了時間が指定されている
          Calendar[ 0 ].createEvent( sTitle, 
                               new Date( TargetYear + "/" + ActiveMonth + "/" + SheetData[ i ][ 0 ] + " " + startTime ),
                               new Date( TargetYear + "/" + ActiveMonth + "/" + SheetData[ i ][ 0 ] + " " + endTime ) );
        }
        else {
          // 終日イベントとして作成する
          if (startTime != "") {
            // 「開始時間」だけが指定されている場合は、イベントのタイトルに "（hh:mm～）" を追加
            sTitle = sTitle + "（" + startTime + "～）";
          }
          else {
            // 「終了時間」だけが指定されている場合は、イベントのタイトルに "（～hh:mm）" を追加
            sTitle = sTitle + "（～" + endTime + "）";
          }
          Calendar[ 0 ].createAllDayEvent( sTitle, new Date( TargetYear + "/" + ActiveMonth + "/" + SheetData[ i ][ 0 ] ) );
        }
        Utilities.sleep( 200 );
      }
    }
  }
  
  // 処理完了表示
  Browser.msgBox( "日程をカレンダーに出力しました。" );
}
