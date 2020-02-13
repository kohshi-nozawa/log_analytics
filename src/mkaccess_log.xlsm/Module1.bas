Attribute VB_Name = "Module1"
Sub mkaccess_log()
  ' ログエクセルファイルの名前を定義
  Dim dates As String
  Dim NewxlsxName As String
  dates = Format(Now, "yyyy-mm-dd")
  NewxlsxName = "access_" & dates & ".xlsx"

  ' テンプレートをコピーして本日のログファイルを作成
  Dim ret As Long
  Dim Current As String
  If Dir(Current & "\" & NewxlsxName) <> "" Then
        ret = MsgBox("同名のファイルが存在します。" & vbCrLf & _
                  "上書きしますか？", vbYesNo)
        If ret = vbNo Then Exit Sub
  End If
  Current = ActiveWorkbook.Path
  FileCopy Current & "\access_" & "temp" & ".xlsx", Current & "\" & NewxlsxName

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
  If IsArray(selectFileName) Then
  ' ログエクセルファイルを変数に格納してアクティブにする
  Dim wb1 As Workbook 
  Dim ws1 As WorkSheet
  Dim n As Long
  Set wb1 = Workbooks(Current & "\" & NewxlsxName)
  Set ws1 = Worksheets("access_log")
  n = 1
    ' 全てのファイルで繰り返し処理を行う
      For Each oneFileName In selectFileName
        Open oneFileName For Input As #1
          Do Until EOF(1)
            Line Input #1, buf
            n = n + 1
            wb1.ws1.Range(Cells(n, 2)) = buf
          Loop
        Close
      Next
      Else
        MsgBox ("ファイルを選択しないで終了")
      End If
End Sub
