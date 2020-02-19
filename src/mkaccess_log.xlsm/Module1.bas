Attribute VB_Name = "Module1"
Sub searchURL()
  With Application
    .Calculation = xlCalculationManual
    .EnableEvents = False
    .ScreenUpdating = False
  End With

  ' ���O�̍Ō�̍s���擾
  Dim lastRow_S1 As Long, lastRow_S2 As Long
  lastRow_S1 = Worksheets("accesslog").Cells(Rows.Count, "B").End(xlUp).Row - 1
  lastRow_S2 = Worksheets("url").Cells(Rows.Count, "A").End(xlUp).Row - 1

  ' url�V�[�g��A��̌��������������
  Dim i As Long, l As Long
  For i = 2 To lastRow_S1
    For l = 2 To lastRow_S2
    Dim searchTarget As String, searchWord As String, returnWord As String, returnCell As Range
    searchTarget = Worksheets("accesslog").Cells(i,10)
    searchWord = Worksheets("url").Cells(l,1)
    returnWord = Worksheets("url").Cells(l,2)
    Set returnCell = Worksheets("accesslog").Cells(i,9)
    If InStr(searchTarget, searchWord) > 0 Then
      returnCell = returnWord
    End If
    Next
  Next

  With Application
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
    .ScreenUpdating = True
  End With
End Sub

Sub margeLog()
  Dim Current As String
  Current = ActiveWorkbook.Path
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
  Dim inputText As String, buf As String, allFile As String
  If IsArray(selectFileName) Then
    allFile = Join(selectFileName, " + ")
    Debug.Print(allFile)
  Else
    MsgBox ("�t�@�C����I�����Ȃ��ŏI�����܂�")
  End If
End Sub

Sub Time_searchURL()
  Dim start As Date: start = Time
  With Application
    .Calculation = xlCalculationManual
    .EnableEvents = False
    .ScreenUpdating = False
  End With

  ' ���O�̍Ō�̍s���擾
  Dim lastRow_S1 As Long, lastRow_S2 As Long
  lastRow_S1 = Worksheets("accesslog").Cells(Rows.Count, "B").End(xlUp).Row - 1
  lastRow_S2 = Worksheets("url").Cells(Rows.Count, "A").End(xlUp).Row - 1

  ' url�V�[�g��A��̌��������������
  Dim i As Long, l As Long
  For i = 2 To lastRow_S1
    For l = 2 To lastRow_S2
    Dim searchTarget As String, searchWord As String, returnWord As String, returnCell As Range
    searchTarget = Worksheets("accesslog").Cells(i,10)
    searchWord = Worksheets("url").Cells(l,1)
    returnWord = Worksheets("url").Cells(l,2)
    Set returnCell = Worksheets("accesslog").Cells(i,9)
    If InStr(searchTarget, searchWord) > 0 Then
      returnCell = returnWord
    End If
    Next
  Next

  With Application
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
    .ScreenUpdating = True
  End With

  Dim finish As Date: finish = Time
  MsgBox "���s���Ԃ� " & Format(finish - start, "nn��ss�b") & " �ł���", vbInformation + vbOKOnly  
End Sub