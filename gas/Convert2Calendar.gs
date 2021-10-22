// �쐬��:2021/4/24
// �쐬��:Tatsuya
// URL:https://kogyo-kyoiku.net

function Convert2Calender() {
  // ���݊J���Ă���V�[�g���擾����
  const ActiveSheet = SpreadsheetApp.getActiveSheet();
  // ���݊J���Ă���V�[�g�̌����擾����
  const ActiveMonth = ActiveSheet.getName().replace("��","");
  // ���݊J���Ă���V�[�g�̃f�[�^��񎟌��z��Ŏ擾����
  let SheetData = ActiveSheet.getDataRange().getValues();
  // �N�x�i1�sA��̃f�[�^�j���擾����
  let TargetYear = SheetData[0][0];
  // �f�t�H���g�̃J�����_�[���擾����
  let Calendar = CalendarApp.getDefaultCalendar();
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
  Browser.msgBox("�������J�����_�[�ɏo�͂��܂����B");
}