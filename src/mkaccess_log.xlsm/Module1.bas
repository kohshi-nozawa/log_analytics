Attribute VB_Name = "Module1"
Sub mkaccess_log()
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
  Dim LF2CRLF As String ,inputText As String
  If IsArray(selectFileName) Then
    ' �S�Ẵt�@�C���ŌJ��Ԃ��������s��
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
    MsgBox ("�t�@�C����I�����Ȃ��ŏI�����܂�")
  End If
  Dim result As Boolean, filePath As String
  filePath = ActiveWorkbook.Path & "\output-date" & ".log"
  result = saveText(filePath, inputText)
  If result Then
    MsgBox sprintf("%s�ɏo�͂��܂����B", filePath)
  Else
    MsgBox "�o�͂Ɏ��s���܂����B�V�X�e���Ǘ��҂ɂ��₢���킹���������B"
  End If

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

End Sub

' LF��CRLF
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