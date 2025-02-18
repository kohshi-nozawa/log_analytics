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
    searchTarget = Worksheets("accesslog").Cells(i, 10)
    searchWord = Worksheets("url").Cells(l, 1)
    returnWord = Worksheets("url").Cells(l, 2)
    Set returnCell = Worksheets("accesslog").Cells(i, 9)
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
  ' logファイルを開くためのダイアログを開く
  ChDrive "C"
  ChDir Current
  selectFileName = _
    Application.GetOpenFilename( _
      FileFilter:="すべてのファイル(*),*.*", _
      FilterIndex:=1, _
      Title:="読み込むファイルを選択してください。", _
      MultiSelect:=True _
    )

  ' 選択したファイルに対する処理
  Dim inputText As String, buf As String, in_allFile As String
  If IsArray(selectFileName) Then
    in_allFile = Join(selectFileName, " + ")
    Debug.Print (in_allFile)
  Else
    MsgBox ("ファイルを選択しないで終了します")
  End If

  'コマンドプロンプトを使うためのオブジェクト
  Dim wsh As New IWshRuntimeLibrary.WshShell
  'コマンド結果を格納する変数
  Dim result As WshExec
 
  Dim cmd As String
  Dim filedata() As String
  Dim i As Integer
  Dim out_allFile As String

  out_allFile = "marge_log.csv"
  Debug.Print (out_allFile)
 
  '実行したいコマンド
  cmd = "copy" & allFile & out_allFile
 
  'コマンドを実行
  Set result = wsh.Exec("%ComSpec% /c " & cmd)
  'コマンドの実行が終わるまで待機
  Do While result.Status = 0
    DoEvents
  Loop

End Sub

Sub Time_searchURL()
  Dim start As Date: start = Time
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
    searchTarget = Worksheets("accesslog").Cells(i, 10)
    searchWord = Worksheets("url").Cells(l, 1)
    returnWord = Worksheets("url").Cells(l, 2)
    Set returnCell = Worksheets("accesslog").Cells(i, 9)
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
  MsgBox "実行時間は " & Format(finish - start, "nn分ss秒") & " でした", vbInformation + vbOKOnly
End Sub
