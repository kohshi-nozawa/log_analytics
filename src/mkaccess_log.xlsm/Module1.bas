Attribute VB_Name = "Module1"
Sub searchURL()
  With Application
    .Calculation = xlCalculationManual
    .EnableEvents = False
    .ScreenUpdating = False
  End With

  ' ログの最後の行を取得
  Dim lastRow_S1 As Long, lastRow_S2 As Long
  lastRow_S1 = Worksheets("accesslog").Cells(Rows.Count, "B").End(xlUp).Row - 1
  lastRow_S2 = Worksheets("url").Cells(Rows.Count, "A").End(xlUp).Row - 1

  ' urlシートのA列の検索文字列を検索
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