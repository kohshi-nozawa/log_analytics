Attribute VB_Name = "Module1"
Sub mkaccess_log()
  ' ログファイルの名前を定義
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

  Dim OpenFileName As Variant
  OpenFileName = Application.GetOpenFilename(FileFilter:="Microsoft Excelブック,*.xls?", _MultiSelect:=True)
    If IsArray(OpenFileName) Then
        ''選択されたファイルへの処理
    Else
        MsgBox "キャンセルされました"
    End If
End Sub
