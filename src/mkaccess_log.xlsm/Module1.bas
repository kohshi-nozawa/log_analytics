Attribute VB_Name = "Module1"
Sub mkaccess_log()
  ' ���O�t�@�C���̖��O���`
  Dim dates As String
  dates = Format(Now, "yyyy-mm-dd")
  Dim NewxlsxName = "access_" & dates & ".xlsx"

  ' �e���v���[�g���R�s�[���Ė{���̃��O�t�@�C�����쐬
  Dim ret As Long
  Dim Current As String
  If Dir("C:\Work\Test.txt") <> "" Then
        ret = MsgBox("�����̃t�@�C�������݂��܂��B" & vbCrLf & _
                  "�㏑�����܂����H", vbYesNo)
        If ret = vbNo Then Exit Sub
  End If
  Current = ActiveWorkbook.Path
  FileCopy Current & "access_" & "temp" & ".xlsx", Current & NewxlsxName

  Dim OpenFileName As Variant
  OpenFileName = Application.GetOpenFilename(FileFilter:="���ׂẴt�@�C��,*.log?", _MultiSelect:=True)
End Sub
