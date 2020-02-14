Attribute VB_Name = "Module1"
Sub mkaccess_log()
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
  Dim LF2CRLF As String ,inputText As String
  If IsArray(selectFileName) Then
    ' 全てのファイルで繰り返し処理を行う
    For Each oneFileName In selectFileName
      Open oneFileName For Input As #1
        Do Until EOF(1)
          Line Input #1, buf
          LF2CRLF = buf
          inputText = inputText & LfToCrlf(LF2CRLF)
        Loop
      Close #1
    Next
  Else
    MsgBox ("ファイルを選択しないで終了します")
  End If

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

End Sub

'// LF→CRLF
Function LfToCrlf(a_sSrc As String) As String
    LfToCrlf = Replace(a_sSrc, vbLf, vbCrLf)
End Function
