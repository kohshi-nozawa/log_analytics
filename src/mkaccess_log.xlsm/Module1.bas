Attribute VB_Name = "Module1"
Sub searchURL()
  ' ログの最後の行を取得
  Dim lastRow_S1 As Long, lastRow_S2 As Long
  lastRow_S1 = Worksheets("accesslog").Cells(Rows.Count, "B").End(xlUp).Row - 1
  lastRow_S1 = Worksheets("url").Cells(Rows.Count, "A").End(xlUp).Row - 1
  Debug.Print(lastRow)

  ' B列の検索文字列を検索
  Dim i As Long, l As Long
  For i = 2 To lastRow_S1
    For l = 2 To lastRow_S2
    Dim searchTarget As String, searchWord As String, returnWord As String, returnCell As Range
    searchTarget = Worksheets("accesslog").Cells(i,10)
    searchWord = Worksheets("url").Cells(l,1)
    returnWord = Worksheets("url").Cells(l,2)
    Set returnCell = Worksheets("accesslog").Cells(i,9)
    If nStr(searchTarget, searchWord) > 0 Then
      returnCell = returnWord
    End If
    Next
  Next
End Sub

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
