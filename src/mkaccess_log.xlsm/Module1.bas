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
  Dim result As Boolean, filePath As String
  filePath = ActiveWorkbook.Path & "\output-date" & ".log"
  result = saveText(filePath, inputText)
  If result Then
    MsgBox sprintf("%sに出力しました。", filePath)
  Else
    MsgBox "出力に失敗しました。システム管理者にお問い合わせください。"
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

' LF→CRLF
Function LfToCrlf(a_sSrc As String) As String
    LfToCrlf = Replace(a_sSrc, vbLf, vbCrLf)
End Function

' Save file
Function saveText(filePath As String, text As String, Optional encoding = "UTF-8") As Boolean

    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2

    Dim Position
    Dim Charset As String
    Dim Bytes
    Position = 0

    Select Case UCase(encoding)
        Case "UTF-8"
            Charset = "utf-8"
            Position = 3
        Case Else
            Charset = encoding
    End Select

    On Error Resume Next

    With CreateObject("ADODB.Stream")
        .Type = adTypeText
        .Charset = Charset
        .Open
        .WriteText text
        .Position = 0
        .Type = adTypeBinary
        .Position = Position
        Bytes = .Read
        .Close
    End With

    With CreateObject("ADODB.Stream")
        .Type = adTypeBinary
        .Open
        .Position = 0
        .Write Bytes
        .SaveToFile filePath, adSaveCreateOverWrite
        .Close
    End With

    If Err.Number = 0 Then
        saveText = True
    Else
        saveText = False
    End If

End Function

' sprintf
'
' @param fmt String フォーマット
' @param prmary ParamArray
' @return result String 文字列変換結果
'
Private Function sprintf(fmt As String, ParamArray prmary()) As String

    Dim i As Long: i = 1
    Dim j As Long: j = LBound(prmary)
    Dim result As String

    Do Until i > Len(fmt)
        If Mid(fmt, i, 1) = "%" Then
            i = i + 1
            Select Case Mid(fmt, i, 1)
                Case "d":
                    result = result & CInt(prmary(j))
                    j = j + 1
                Case "f":
                    result = result & CDbl(prmary(j))
                    j = j + 1
                Case "s":
                    result = result & CStr(prmary(j))
                    j = j + 1
                Case "%":
                    result = result & "%"
                Case Else:
                    Debug.Print "無効な識別子"
            End Select
        Else
            result = result & Mid(fmt, i, 1)
        End If
        i = i + 1
    Loop
    sprintf = result
End Function