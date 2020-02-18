Attribute VB_Name = "Module1"
Sub searchURL()
  Dim l As Long
  l = Cells(Rows.Count, "B").End(xlUp).Row
  Debug.Print(l)
End Sub

' sprintf
'
' @param fmt String �t�H�[�}�b�g
' @param prmary ParamArray
' @return result String ������ϊ�����
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
                    Debug.Print "�����Ȏ��ʎq"
            End Select
        Else
            result = result & Mid(fmt, i, 1)
        End If
        i = i + 1
    Loop
    sprintf = result
End Function
