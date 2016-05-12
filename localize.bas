Attribute VB_Name = "localize"
Type genString
    gsOriginal As String
    gsTranslated As String
End Type

Dim langCache() As String
Dim genString() As genString

Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFilename As String) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)



Sub CacheTranslateEx(path As String)
    
    On Error Resume Next
    

    ReDim genString(0) As genString
    
    Dim i As Integer
    Dim tmpString As String
    Dim trsString As String
    
    i = FreeFile
    Open path For Input As #i
    If FreeFile = i Then Exit Sub

findmark:
    
    ' Seeking for string
    Do
        Line Input #i, tmpString
        If EOF(i) Then GoTo theend
    Loop While Not (Left(tmpString, 1) = "[" And Right(tmpString, 1) = "]")
    
    Line Input #i, trsString
    
    ReDim Preserve genString(0 To UBound(genString) + 1) As genString
  
    genString(UBound(genString)).gsOriginal = Mid(tmpString, 2, Len(tmpString) - 2)
    genString(UBound(genString)).gsTranslated = trsString
    
    GoTo findmark
    
theend:
    Close #i
    
End Sub

Sub TranslateForm(inForm As Form)

    On Error Resume Next
    
    Dim a As Control
    Dim b As String
    
    inForm.Caption = TranslateEx(inForm.Caption)
    
    For Each a In inForm.Controls
        If TypeOf a Is CommandButton _
        Or TypeOf a Is Label _
        Or TypeOf a Is CheckBox _
        Or TypeOf a Is OptionButton _
        Or TypeOf a Is Frame _
        Or TypeOf a Is Menu Then
            a.ToolTipText = TranslateEx(a.ToolTipText)
            a.Caption = TranslateEx(a.Caption)
        End If
    Next a

End Sub


Function TranslateEx(Expression As String) As String
    
    
    On Error Resume Next

    Dim Index As Integer, tmpText As genString
    Dim Jndex As Integer
    For Index = 0 To UBound(genString)
        If genString(Index).gsOriginal = Expression Then
            TranslateEx = Replace(genString(Index).gsTranslated, "\\n", vbCrLf)
            tmpText = genString(Index)
            For Jndex = Index To 1 Step -1
                genString(Jndex) = genString(Jndex - 1)
            Next Jndex
            genString(0) = tmpText
            Exit Function
        End If
        
    Next Index
    
    TranslateEx = Replace(Expression, "\\n", vbCrLf)
    
End Function

Sub CacheStrings(FileName As String)

Dim tmpStr As String
tmpStr = LoadFile(FileName)
langCache = Split(tmpStr & vbCrLf, vbCrLf)

End Sub

Function gl(inMASK As String, Optional inDefault As String = "???") As String

On Error Resume Next

Dim Index As Integer, tmpText As String
Dim Jndex As Integer
For Index = LBound(langCache) To UBound(langCache)
    If UCase$(Mid$(langCache(Index), 1, Len(inMASK & "="))) = UCase(inMASK & "=") Then
      gl = Mid$(langCache(Index), Len(inMASK & "=") + 1)
      tmpText = langCache(Index)
      For Jndex = Index To 1 Step -1
        langCache(Jndex) = langCache(Jndex - 1)
      Next Jndex
      langCache(0) = tmpText
      Exit Function
    End If
Next Index
gl = inDefault

End Function

Function TR(Expression As String) As String

End Function

Function GetLanguage2(inMASK As String, Optional inDefault As String = "???") As String

Dim Index As Integer
For Index = LBound(langCache) To UBound(langCache)
If UCase$(Right$(langCache(Index), Len("=" & inMASK))) = UCase("=" & inMASK) Then
   GetLanguage2 = Mid$(langCache(Index), 1, Len(langCache(Index)) - Len("=" & inMASK))
   Exit Function
End If
Next Index

GetLanguage2 = inDefault

End Function

Private Function LoadFile(FileName As String) As String
Dim f As Long, b() As Byte, IC As Long
f = FreeFile
If CheckFile(FileName) <> 1 Then Exit Function
Open FileName For Binary As f
  IC = LOF(f)
  If Not IC = 0 Then
    ReDim b(1 To IC) As Byte
    Get #f, 1, b()
    LoadFile = String(IC, " ")
    CopyMemory ByVal LoadFile, b(1), IC
  End If
Close f
End Function

Private Function CheckFile(Name As String) As Integer
Dim s As Long
s = GetFileAttributes(Name)
If s = -1 Then CheckFile = 0: Exit Function
If s And &H10 Then CheckFile = 2: Exit Function
CheckFile = 1
End Function
