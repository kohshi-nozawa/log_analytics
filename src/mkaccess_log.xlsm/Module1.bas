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
  ChDrive "C"
  ChDir Current
  selectFileName = _
    Application.GetOpenFilename( _
      FileFilter:="���ׂẴt�@�C��(*),*.*", _
      FilterIndex:=1, _
      Title:="�ǂݍ��ރt�@�C����I�����Ă��������B", _
      MultiSelect:=True _
    )
  ' �I�������t�@�C���ɑ΂��鏈��
  If IsArray(selectFileName) Then
  Dim n As Long
  Workbooks.Open (Current & "\" & NewxlsxName)
  n = 1
    ' �S�Ẵt�@�C���ŌJ��Ԃ��������s��
      For Each oneFileName In selectFileName
        Open oneFileName For Input As #1
        Dim fso As New FileSystemObject
        Dim ts As TextStream
        Dim line As String
        Dim items() As String

        ' �t�@�C�����J��
          Set ts = fso.OpenTextFile(oneFileName, ForRreading)
          Do Until ts.AtEndOfStream
            line = ts.ReadLine
            items = Split(line, ",")
            Debug.Print UBound(items)
            n = n + 1
            Cells(n, 2) = buf
            Cells(n, 2).WrapText = False
          Loop
          ts.Close
        Close #1
      Next
      Else
        MsgBox ("�t�@�C����I�����Ȃ��ŏI��")
      End If
End Sub
