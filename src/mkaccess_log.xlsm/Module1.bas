Attribute VB_Name = "Module1"
Sub mkaccess_log()
  ' ���O�t�@�C���̖��O���`
  Dim dates As String
  Dim NewxlsxName As String
  dates = Format(Now, "yyyy-mm-dd")
  NewxlsxName = "access_" & dates & ".xlsx"

  ' �e���v���[�g���R�s�[���Ė{���̃��O�t�@�C�����쐬
  Dim ret As Long
  Dim Current As String
  Dim TempFullPath As String
  Dim NewFullPath As String
  TempFullPath = Current & "\access_" & "temp" & ".xlsx"
  NewFullPath = Current & "\" & NewxlsxName
  If Dir(NewFullPath) <> "" Then
        ret = MsgBox("�����̃t�@�C�������݂��܂��B" & vbCrLf & _
                  "�㏑�����܂����H", vbYesNo)
        If ret = vbNo Then Exit Sub
  End If
  Current = ActiveWorkbook.Path
  FileCopy TempFullPath, NewFullPath

End Sub
