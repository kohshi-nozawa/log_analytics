Attribute VB_Name = "Module1"
Sub mkaccess_log()
  ' ���O�G�N�Z���t�@�C���̖��O���`
  Dim dates As String
  Dim NewxlsxName As String
  dates = Format(Now, "yyyy-mm-dd")
  NewxlsxName = "access_" & dates & ".xlsx"

  ' �e���v���[�g���R�s�[���Ė{���̃��O�t�@�C�����쐬
  Dim ret As Long
  Dim Current As String
  If Dir(Current & "\" & NewxlsxName) <> "" Then
        ret = MsgBox("�����̃t�@�C�������݂��܂��B" & vbCrLf & _
                  "�㏑�����܂����H", vbYesNo)
        If ret = vbNo Then Exit Sub
  End If
  Current = ActiveWorkbook.Path
  FileCopy Current & "\access_" & "temp" & ".xlsx", Current & "\" & NewxlsxName

  ' log�t�@�C�����J�����߂̃_�C�A���O���J��
  selectFileName = _
    Application.GetOpenFilename( _
      FileFilter:="���ׂẴt�@�C��(*),*.*", _
      FilterIndex:=1, _
      Title:="�ǂݍ��ރt�@�C����I�����Ă��������B", _
      MultiSelect:=True _
    )
  ' �I�������t�@�C���ɑ΂��鏈��
  If IsArray(selectFileName) Then
  ' ���O�G�N�Z���t�@�C����ϐ��Ɋi�[���ăA�N�e�B�u�ɂ���
  Dim wb1 As Workbook
  Dim n As Long
  Workbooks.Open ThisWorkbook.Path & "\" & NewxlsxName
  Set wb1 = ActiveWorkbook
  n = 1
    ' �S�Ẵt�@�C���ŌJ��Ԃ��������s��
      For Each oneFileName In selectFileName
        Open oneFileName For Input As #1
          Do Until EOF(1)
            Line Input #1, buf
            n = n + 1
            ThisWorkbook.Worksheets("access_log").Cells(n, 2) = buf
        Close
      Next
      Else
        MsgBox ("�t�@�C����I�����Ȃ��ŏI��")
      End If
End Sub
