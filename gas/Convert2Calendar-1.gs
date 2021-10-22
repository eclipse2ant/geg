// �쐬��:2021/5/2
// �쐬��:Tatsuya
// URL:https://kogyo-kyoiku.net

function Convert2Calender() {
  // �J�����_�[�������b�Z�[�W�{�b�N�X����擾����
  let CalendarName = Browser.inputBox("�o�^�������J�����_�[�̖��O����͂��Ă�������\\n�i�󔒂̏ꍇ�̓f�t�H���g�J�����_�[�ɂȂ�܂��j");
  // ���݊J���Ă���V�[�g���擾����
  const ActiveSheet = SpreadsheetApp.getActiveSheet();
  // ���݊J���Ă���V�[�g�̌����擾����
  const ActiveMonth = ActiveSheet.getName().replace("��","");
  // ���݊J���Ă���V�[�g�̃f�[�^��񎟌��z��Ŏ擾����
  let SheetData = ActiveSheet.getDataRange().getValues();
  // �N�x�i1�sA��̃f�[�^�j���擾����
  let TargetYear = SheetData[0][0];
  // �J�����_�[�I�u�W�F�N�g�i�[�ϐ���錾
  let Calendar;
  if(CalendarName == ""){
    // �J�����_�[�����w�肳��Ă��Ȃ���΁A�f�t�H���g�J�����_�[���擾
    Calendar = CalendarApp.getDefaultCalendar();
  }else{
    // �J�����_�[�����w�肳��Ă���΁A���̃J�����_�[���擾
    Calendar = CalendarApp.getOwnedCalendarsByName(CalendarName)[0];
  }
  if(Calendar==null){
    // �J�����_�[���擾�ł��Ȃ������ꍇ�̏���
    Browser.msgBox("�J�����_�[���u"+CalendarName+"�v��������܂���ł����B\\n���������O����͂������Ă��������B\\n�������I�����܂��B");
    // �X�N���v�g�̏I������
    return;
  }
  // �J�����_�[�̃^�C���]�[���̐ݒ�
  Calendar.setTimeZone("Asia/Tokyo");
  // ���������������Ă�������
  for(let i = 1 ; i < SheetData.length ; i++){
    for(let j = 2 ; j < SheetData[i].length ; j+=3){
      if(SheetData[i][j] != ""){
        // �\�肪�������ꍇ�ɁA�J�n���Ԃ��Q�Ƃ��ė\����쐬����
        Calendar.createEvent(SheetData[i][j],
                              new Date(TargetYear+"/"+ActiveMonth+"/"+SheetData[i][0]+" "+SheetData[i][j+1]),
                              new Date(TargetYear+"/"+ActiveMonth+"/"+SheetData[i][0]+" "+SheetData[i][j+2]));
      }
    }
  }
  // ���������\��
  Browser.msgBox("�J�����_�[���u"+Calendar.getName()+"�v�ɓ������o�͂��܂����B");
}