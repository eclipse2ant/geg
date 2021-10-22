'use strict'

//----------------------------------------------------------------------------
// �`�F�b�N���ʂ��N���A
function checkClear( ) {
  checkClear0( );
  checkClear1( );
}

//----------------------------------------------------------------------------
// �`�F�b�N���ʂ��N���A(�� A)
function checkClear0( ) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('�J�����_�[���');

  // ���o�����Đݒ�
  sheet.getRange( 'A1' ).setValue( "���s�A�J�E���g" );
  sheet.getRange( 'A4' ).setValue( "�o�^����Ă���J�����_�[�̐�" );

  // �ς̗����N���A
  sheet.getRange( 'A2' ).setValue( "" );
  sheet.getRange( 'A5' ).setValue( "" );
}

//----------------------------------------------------------------------------
// �`�F�b�N���ʂ��N���A(�� B�`I)
function checkClear1( ) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('�J�����_�[���');
  
  // ��U�A�͈͂����ׂč폜
  sheet.getRange( 'B:I' ).clearContent( );
  
  // �擪�s�̌��o�����Đݒ�
  sheet.getRange( 'C1' ).setValue( "�o�^����Ă���J�����_�[�� ID" );
  sheet.getRange( 'D1' ).setValue( "����" );
  sheet.getRange( 'E1' ).setValue( "isHidden" );
  sheet.getRange( 'F1' ).setValue( "isMyPrimaryCalendar" );
  sheet.getRange( 'G1' ).setValue( "isOwnedByMe" );
  sheet.getRange( 'H1' ).setValue( "isSelected" );
  sheet.getRange( 'I1' ).setValue( "Timezone" );
  sheet.getRange( 'A4' ).setValue( "�o�^����Ă���J�����_�[�̐�" );

  sheet.getRange( 'A5' ).setValue( "" );
}

//----------------------------------------------------------------------------
// ���݂̃A�J�E���g�̃J�����_�[�����擾
function getMyCalendars( ) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('�J�����_�[���');

  // ���݂̎��s�A�J�E���g���m�F    �����s����Ă���A�J�E���g�ɂ���āA���삪�ς���Ă���̂�...
  var eMailID = Session.getActiveUser( ).getUserLoginId( );
  checkClear0( );
  sheet.getRange( 'A2' ).setValue( eMailID );
  
  // ���[�U�[�����L�܂��̓T�u�X�N���C�u���Ă��邷�ׂẴJ�����_�[���擾���܂��B
  // Determines how many calendars the user can access.
  var calendars = CalendarApp.getAllCalendars( );
  checkClear1( );
  sheet.getRange( 'A5' ).setValue( calendars.length );     // �J�����_�[�̑���
  
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
// ���͂���Ă���C�x���g�����N���A
function eventClear( ) {
  var sheet = SpreadsheetApp.getActiveSheet( );
  var result = Browser.msgBox( "���͂���Ă���C�x���g�����N���A���Ă��\���܂��񂩁H\\n �y���Ӂz ���̑���͎������܂���I", Browser.Buttons.OK_CANCEL );
  if ( result == "ok" ) {
    sheet.getRange( 'C3:Z33' ).clearContent( );
  }
}

//----------------------------------------------------------------------------
// �쐬��:2021/4/24
// �쐬��:Tatsuya
// URL:https://kogyo-kyoiku.net    
//     https://kogyo-kyoiku.net/spreadsheet2calendar/?fbclid=IwAR3IdUnxCZ--_9woyNCH9H210zLE4pPwdwl1Im8TgFnFbHy6m-lo6mR1U2E
// ���ρFT.Sakamoto 2021/05/01
function Convert2Calender( ) {
  var result = Browser.msgBox( "���͂���Ă���C�x���g���ŁA�C�x���g���o�͂��Ă��\���܂��񂩁H\\n �y���Ӂz ���̑���͎������܂���I", Browser.Buttons.OK_CANCEL );
  if ( result != "ok" ) {
    // �ȍ~�̏������s��Ȃ�
    return;
  }
  
  // ���݊J���Ă���V�[�g���擾����
  const ActiveSheet = SpreadsheetApp.getActiveSheet( );
  
  // ���݊J���Ă���V�[�g�̃f�[�^��񎟌��z��Ŏ擾����
  let SheetData = ActiveSheet.getRange( 'A1:Z33' ).getValues( );
  
  // �����Ώۂ̌����擾����
  const ActiveMonth = SheetData[ 0 ][ 1 ];     // B1
  
  // �N�x�i1�sA��̃f�[�^�j���擾����
  let TargetYear = SheetData[ 0 ][ 0 ];        // A1
  
  // �f�t�H���g�̃J�����_�[���擾����
  let Calendar = CalendarApp.getCalendarsByName( SheetData[ 0 ][ 5 ] );
  
  // �J�����_�[�̃^�C���]�[���̐ݒ�
  Calendar[ 0 ].setTimeZone("Asia/Tokyo");

  // ���������������Ă�������
  for( let i = 2 ; i < SheetData.length ; i++ ) {
    for( let j = 2 ; j < SheetData[ i ].length ; j += 3 ) {
      let sTitle = SheetData[ i ][ j ];
      if( sTitle != "" ) {
        // �\��ɉ������͂���Ă���΁A�C�x���g(�\��)���쐬����
        let startTime = SheetData[ i ][ j + 1 ];
        let endTime = SheetData[ i ][ j + 2 ];
        
        if ( ( startTime != "" ) && ( endTime != "" ) ) {
          // �J�n���ԁE�I�����Ԃ��w�肳��Ă���
          Calendar[ 0 ].createEvent( sTitle, 
                               new Date( TargetYear + "/" + ActiveMonth + "/" + SheetData[ i ][ 0 ] + " " + startTime ),
                               new Date( TargetYear + "/" + ActiveMonth + "/" + SheetData[ i ][ 0 ] + " " + endTime ) );
        }
        else {
          // �I���C�x���g�Ƃ��č쐬����
          if (startTime != "") {
            // �u�J�n���ԁv�������w�肳��Ă���ꍇ�́A�C�x���g�̃^�C�g���� "�ihh:mm�`�j" ��ǉ�
            sTitle = sTitle + "�i" + startTime + "�`�j";
          }
          else {
            // �u�I�����ԁv�������w�肳��Ă���ꍇ�́A�C�x���g�̃^�C�g���� "�i�`hh:mm�j" ��ǉ�
            sTitle = sTitle + "�i�`" + endTime + "�j";
          }
          Calendar[ 0 ].createAllDayEvent( sTitle, new Date( TargetYear + "/" + ActiveMonth + "/" + SheetData[ i ][ 0 ] ) );
        }
        Utilities.sleep( 200 );
      }
    }
  }
  
  // ���������\��
  Browser.msgBox( "�������J�����_�[�ɏo�͂��܂����B" );
}
