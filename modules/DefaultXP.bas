Attribute VB_Name = "Default_Subs"
Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Const SPI_GETWORKAREA = 48
Const None = ""

Global IDXPosIn(32000) As Long
Global IDXPosOut(32000) As Long
Function Biggest(a As Integer, b As Integer) As Integer
If a > b Then Biggest = a Else Biggest = b
End Function

Function MkDirEx(path As String) As Boolean

    On Error Resume Next
    
    Dim i As Integer, j As Integer
    j = 1
    
    Do
        j = InStr(j + 1, path, "\")
        If j > 0 Then
            MkDirOnDemand Mid(path, 1, j - 1)
        Else
            MkDirOnDemand path
        End If
    Loop While Not InStr(j, path, "\") = 0

End Function

Function DirExists(path As String) As Boolean

    Select Case Len(Dir(path, vbDirectory))
    Case 0
        DirExists = False
    Case Else
        DirExists = True
    End Select
    
End Function

Sub MkDirOnDemand(path As String)
    
    On Error Resume Next
    
    If Len(Dir(path, vbDirectory)) = 0 Then MkDir path

End Sub

Function NoExe(iStr As String) As String
    Dim Y
    Y = InStrRev(iStr, ".")
    If Y > 0 Then
        NoExe = Mid(iStr, 1, Y - 1)
    Else
        NoExe = iStr
    End If
End Function

Function KillNull(Expression As String) As String

    N = InStr(Expression, Chr(0))

    If N Then KillNull = Mid(Expression, 1, N - 1) Else KillNull = Expression

End Function

Function ValW(Expression)
    Dim i As Integer, j As String
    For i = 1 To Len(Expression)
        If Asc(Mid(Expression, i, 1)) >= vbKey0 And Asc(Mid(Expression, i, 1)) <= vbKey9 Then _
                j = j & Mid(Expression, i, 1)
    Next i
    ValW = j
End Function

Function TopValue(inArray())
    Dim a, b
    For a = LBound(inArray) To UBound(inArray)
        If b < inArray(a) Then b = inArray(a)
    Next a
    TopValue = b
End Function

Function DateX(inSeconds As Long) As String
    Dim xDy, xHr, xMn, xSc
    
    xDy = Fix(inSeconds / 86400)
    xHr = Fix(inSeconds / 3600) Mod 24
    xMn = Fix(inSeconds / 60) Mod 60
    xSc = inSeconds Mod 60
    
    DateX = IIf(xDy, Format(xDy, "0д") & ". ", "") & IIf(xHr, Format(xHr, "0ч") & ". ", "") & IIf(xMn, Format(xMn, "0м") & ". ", "") & IIf(xSc, Format(xSc, "0с") & ".", "")
    If DateX = "" Then DateX = "0c."
        
End Function


Function DateBorn(inBorn As String) As String
    
    Dim X As Integer, Y As Integer, z As Integer
    z = DateDiff("d", inBorn, Now)
    
    GetDateFromLong CInt(z), X, Y, Year(inBorn)
    
    DateBorn = Format(Fix(z / 365), "0") & "г. " & Format(Y, "0") & "мес. " & Format(X, "0") & "дн. "
        
End Function


Function FormatEx(xExpression, xFormat) As String
Dim T As String
T = Trim(TrimEx(Format(xExpression, xFormat)))
FormatEx = Replace(T, ",", ".")
End Function

Function LeadingEx(inString As String, inZeros As Integer) As String

Dim T As Integer
T = inZeros - Len(inString)
If T >= 0 Then LeadingEx = Space(T) & inString Else LeadingEx = inString

End Function

Function Par(Expression)
 If Expression Mod 2 = 0 Then Par = Expression Else Par = Expression - 1
End Function

Function tppX()
 tppX = Screen.TwipsPerPixelX
End Function

Function tppY()
 tppY = Screen.TwipsPerPixelY
End Function

Function RGBBright(inRGB As Long, inBright As Integer) As Long

Dim iG As Byte, IB As Byte

Dim IC(2) As Byte

Call CopyMemory(IC(0), inRGB, 3)

RGBBright = RGB(ColorLimit(IC(0) / 255 * inBright), ColorLimit(IC(1) / 255 * inBright), ColorLimit(IC(2) / 255 * inBright))

End Function

Function ColorLimit(inColor As Integer) As Byte
 ColorLimit = IIf(inColor >= 255, 255, inColor)
End Function


Function GetDistanceToControl(inForm As Form, inControl As Control, Optional X As Long = -1, Optional Y As Long = -1, Optional sX As Long, Optional sY As Long) As Long

Dim mA As POINTAPI
Dim MouseCoords As POINTAPI
Dim d(1 To 4) As Long
Dim s(1 To 4) As Long
Dim tmpDIST As Long
Dim X1, Y1, X2, Y2

' Если координаты точки не заданы используем координаты мышки
If X = -1 And Y = -1 Then _
  GetCursorPos MouseCoords: _
  X = 15 * MouseCoords.X: _
  Y = 15 * MouseCoords.Y
  
  
' Определение углов элемента управления
X1 = inForm.Left + inControl.Left
Y1 = inForm.Top + inControl.Top
X2 = inForm.Left + inControl.Left + inControl.Width
Y2 = inForm.Top + inControl.Top + inControl.Height

' Нахождение ближайшей стороны или угла
If X >= X1 And X <= X2 And Y <= Y1 And Y <= Y2 Then
    tmpDIST = (Y1 - Y)
    sX = X
    sY = Y1
ElseIf X >= X1 And X <= X2 And Y > Y1 And Y >= Y2 Then
    tmpDIST = (Y - Y2)
    sX = X
    sY = Y2
ElseIf Y >= Y1 And Y <= Y2 And X < X1 And X <= X2 Then
    tmpDIST = (X1 - X)
    sX = X1
    sY = Y
ElseIf Y >= Y1 And Y <= Y2 And X > X1 And X >= X2 Then
    tmpDIST = (X - X2)
    sX = X2
    sY = Y
ElseIf X < X1 And Y < Y1 Then
    tmpDIST = Sqr((X1 - X) ^ 2 + (Y1 - Y) ^ 2)
    sX = X1
    sY = Y1
ElseIf X > X2 And Y < Y1 Then
    tmpDIST = Sqr((X2 - X) ^ 2 + (Y1 - Y) ^ 2)
    sX = X2
    sY = Y1
ElseIf X < X1 And Y > Y2 Then
    tmpDIST = Sqr((X1 - X) ^ 2 + (Y2 - Y) ^ 2)
    sX = X1
    sY = Y2
ElseIf X > X2 And Y > Y2 Then
    tmpDIST = Sqr((X2 - X) ^ 2 + (Y2 - Y) ^ 2)
    sX = X2
    sY = Y2
Else
    sX = X
    sY = Y
    tmpDIST = 0
End If

GetDistanceToControl = tmpDIST

End Function

Function DistanceToWindow(inhandle As Long, Optional X As Long = -1, Optional Y As Long = -1, Optional sX As Long, Optional sY As Long) As Long

Dim mA As POINTAPI
Dim MouseCoords As POINTAPI
Dim uw As RECT
Dim d(1 To 4) As Long
Dim s(1 To 4) As Long
Dim tmpDIST As Long
Dim X1, Y1, X2, Y2

' Если координаты точки не заданы используем координаты мышки
If X = -1 And Y = -1 Then _
  GetCursorPos MouseCoords: _
  X = tppX * MouseCoords.X: _
  Y = tppY * MouseCoords.Y
  
Call GetWindowRect(inhandle, uw)
  
' Определение углов элемента управления
X1 = uw.Left * tppX
Y1 = uw.Top * tppY
X2 = uw.Right * tppX
Y2 = uw.Bottom * tppY

' Нахождение ближайшей стороны или угла
If X >= X1 And X <= X2 And Y <= Y1 And Y <= Y2 Then
    tmpDIST = (Y1 - Y)
    sX = X
    sY = Y1
ElseIf X >= X1 And X <= X2 And Y > Y1 And Y >= Y2 Then
    tmpDIST = (Y - Y2)
    sX = X
    sY = Y2
ElseIf Y >= Y1 And Y <= Y2 And X < X1 And X <= X2 Then
    tmpDIST = (X1 - X)
    sX = X1
    sY = Y
ElseIf Y >= Y1 And Y <= Y2 And X > X1 And X >= X2 Then
    tmpDIST = (X - X2)
    sX = X2
    sY = Y
ElseIf X < X1 And Y < Y1 Then
    tmpDIST = Sqr((X1 - X) ^ 2 + (Y1 - Y) ^ 2)
    sX = X1
    sY = Y1
ElseIf X > X2 And Y < Y1 Then
    tmpDIST = Sqr((X2 - X) ^ 2 + (Y1 - Y) ^ 2)
    sX = X2
    sY = Y1
ElseIf X < X1 And Y > Y2 Then
    tmpDIST = Sqr((X1 - X) ^ 2 + (Y2 - Y) ^ 2)
    sX = X1
    sY = Y2
ElseIf X > X2 And Y > Y2 Then
    tmpDIST = Sqr((X2 - X) ^ 2 + (Y2 - Y) ^ 2)
    sX = X2
    sY = Y2
Else
    sX = X
    sY = Y
    tmpDIST = 0
End If

DistanceToWindow = tmpDIST

End Function

Public Function GetIniRecord(Record As String, INIFile As String, Optional rDefault = "") As String
Dim CfgLine As String, g As Integer
On Error Resume Next
g = FreeFile
Open INIFile For Input As #g
Do
Line Input #g, CfgLine
If UCase$(Mid$(CfgLine, 1, Len(Record))) = UCase(Record) Then
   GetIniRecord = Mid$(CfgLine, Len(Record) + 1)
   Close g: Exit Function
End If
Loop While Not EOF(g)
GetIniRecord = Format(rDefault)
Close g
End Function


Function TrimEx(xExpression As String) As String
Dim T As Integer
For T = 1 To Len(xExpression)
    If Asc(Mid(xExpression, T, 1)) > 32 Then
     If Mid(xExpression, Len(xExpression), 1) = "," Or Mid(xExpression, Len(xExpression), 1) = "." Then
      TrimEx = Mid(xExpression, T)
      TrimEx = Left(TrimEx, Len(TrimEx) - 1)
      Exit Function
     Else
      TrimEx = Mid(xExpression, T)
      Exit Function
     End If
    End If
Next T
End Function

Function BeginsWith(inString As String, inInclude As String) As Boolean

If Mid(inString, 1, Len(inInclude)) = inInclude Then BeginsWith = True Else BeginsWith = False

End Function

Function NotInteger(inValue As Variant) As Boolean
If Fix(inValue) = inValue Then NotInteger = True Else NotInteger = False
End Function

Sub Modulate(Expression As Variant)
 If Expression < 0 Then Expression = 0
End Sub
Function ModulateEx(Expression As Variant) As Variant
 If Expression < 0 Then ModulateEx = 0 Else ModulateEx = Expression
End Function
Sub Summ(inValue, Optional inAdd = 1)
 inValue = inValue + inAdd
End Sub

Function GetDaysInMonth(inMonth As Integer) As Integer
 Select Case inMonth
 Case 1: GetDaysInMonth = 31
 Case 2: GetDaysInMonth = 28.25
 Case 3: GetDaysInMonth = 31
 Case 4: GetDaysInMonth = 30
 Case 5: GetDaysInMonth = 31
 Case 6: GetDaysInMonth = 30
 Case 7: GetDaysInMonth = 31
 Case 8: GetDaysInMonth = 30
 Case 9: GetDaysInMonth = 30
 Case 10: GetDaysInMonth = 31
 Case 11: GetDaysInMonth = 30
 Case 12: GetDaysInMonth = 31
 End Select
End Function

Sub FillIn(fform As Form)
SetLayeredWindowAttributes fform.hWnd, 0, 0, LWA_ALPHA

fform.Visible = True
For Y = 0 To 200 Step 2
  DoEvents
  NormalWindowStyle = GetWindowLong(fform.hWnd, GWL_EXSTYLE)
  SetWindowLong fform.hWnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
  SetLayeredWindowAttributes fform.hWnd, 0, Y, LWA_ALPHA
Next Y

End Sub

Function CWF(logic As Boolean, a As Variant, b As Variant) As Variant
If logic Then CWF = a Else CWF = b
End Function

Sub FillOut(fform As Form)

For Y = 200 To 0 Step -2
  DoEvents
  NormalWindowStyle = GetWindowLong(fform.hWnd, GWL_EXSTYLE)
  SetWindowLong fform.hWnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
  SetLayeredWindowAttributes fform.hWnd, 0, Y, LWA_ALPHA
Next Y
fform.Visible = False

End Sub



Function Trim32(inString As String)

inString = Trim(inString)

For X = Len(inString) To 1 Step -1
 If Asc(Mid(inString, X, 1)) > 32 Then inString = Mid(inString, 1, X): Exit For
Next

For X = 1 To Len(inString) Step 1
 If Asc(Mid(inString, X, 1)) < 32 Then Trim32 = Mid(inString, 1, X - 1): Exit Function
Next

Trim32 = inString

End Function

Function Scaler(SA As Long, sB As Long, sStep As Long, sSteps As Long) As Long

Dim sC As Long

sC = sB - SA

Scaler = sC + (SA / sSteps * sStep)

End Function

Function dsRes(Expression, StepSize)
    dsRes = IIf(Fix(Expression / StepSize) = Expression / StepSize, Expression, StepSize + Fix(Expression / StepSize) * StepSize)
End Function

Sub Задержка(Миллисекунд As Long)

Dim ВремЗнач As Long

ВремЗнач = GetTickCount

Do: DoEvents: Loop While Not GetTickCount - ВремЗнач > Миллисекунд

End Sub
Sub Dream(ms As Long)

Dim vz As Long

vz = GetTickCount

Do: DoEvents: Loop While Not GetTickCount - vz > ms

End Sub
Function MMod(mLong)
  If mLong < 0 Then MMod = -mLong Else MMod = mLong
End Function

Function MaxVal(inVal1, inVal2)
If inVal1 > inVal2 Then MaxVal = inVal1 Else MaxVal = inVal2
End Function

Function MaxValue(ByRef inVal1 As Long, ByRef inVal2 As Long) As Long

If inVal1 > inVal2 Then MaxValue = inVal1 Else MaxValue = inVal2

End Function

Function CountChars(inChar As String, inString As String) As Integer
k = 0
For X = 1 To Len(inString)
 If Mid(inString, X, 1) = inChar Then k = k + 1
Next
CountChars = k
End Function

Function EnPass(inText) As String
Dim inTmp, X

inTmp = String(Len(inText), 32)

For X = 1 To Len(inText)
 Mid(inTmp, X, 1) = Chr(255 - Asc(Mid(inText, X, 1)))
Next

EnPass = inTmp

End Function


Function GetLongFromData(inDay As Integer, inMonth As Integer, inYear As Integer) As Currency

If inYear Mod 4 = 0 Then N = 1 Else N = 0
If inMonth = 1 Then temp = inDay
If inMonth = 2 Then temp = inDay + 31
If inMonth = 3 Then temp = inDay + 31 + 28 + N
If inMonth = 4 Then temp = inDay + 31 + 28 + 31 + N
If inMonth = 5 Then temp = inDay + 31 + 28 + 31 + 30 + N
If inMonth = 6 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + N
If inMonth = 7 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + 30 + N
If inMonth = 8 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + N
If inMonth = 9 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + N
If inMonth = 10 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + N
If inMonth = 11 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + N
If inMonth = 12 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30 + N
GetLongFromData = temp

End Function
Function GetLongFromDataEx(inDay As Integer, inMonth As Integer, inYear As Integer) As Currency

If inYear Mod 4 = 0 Then N = 1 Else N = 0
If inMonth = 1 Then temp = inDay
If inMonth = 2 Then temp = inDay + 31
If inMonth = 3 Then temp = inDay + 31 + 28 + N
If inMonth = 4 Then temp = inDay + 31 + 28 + 31 + N
If inMonth = 5 Then temp = inDay + 31 + 28 + 31 + 30 + N
If inMonth = 6 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + N
If inMonth = 7 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + 30 + N
If inMonth = 8 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + N
If inMonth = 9 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + N
If inMonth = 10 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + N
If inMonth = 11 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + N
If inMonth = 12 Then temp = inDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30 + N
GetLongFromDataEx = inYear * (365 + N) + temp

End Function
Sub GetDateFromLong(InLong As Integer, inDay As Integer, inMonth As Integer, inYear As Integer)

Dim temp As Integer, a As Integer, b As Integer, N As Integer

If inYear Mod 4 = 0 Then N = 1 Else N = 0
temp = InLong Mod 365 + N
If temp > 0 Then a = temp: b = 1
If temp > 31 Then a = temp - 31: b = 2
If temp > 59 + N Then a = temp - 59 - N: b = 3
If temp > 90 + N Then a = temp - 90 - N: b = 4
If temp > 120 + N Then a = temp - 120 - N: b = 5
If temp > 151 + N Then a = temp - 151 - N: b = 6
If temp > 181 + N Then a = temp - 181 - N: b = 7
If temp > 212 + N Then a = temp - 212 - N: b = 8
If temp > 243 + N Then a = temp - 243 - N: b = 9
If temp > 273 + N Then a = temp - 273 - N: b = 10
If temp > 304 + N Then a = temp - 304 - N: b = 11
If temp > 334 + N Then a = temp - 334 - N: b = 12

inDay = a
inMonth = b - 1

End Sub


Function GetFileSize(FName As String) As Long
 On Error Resume Next
 i = FreeFile
 Open FName For Input As #i
 GetFileSize = LOF(i)
 Close i
End Function


Function GetTimeFromMinutes(vMinutes As Long)
If vMinutes < 60 * 60 Then GetTimeFromMinutes = Format$(Fix(vMinutes / 60), "00") & ":" & Format$(Fix(vMinutes Mod 60), "00")
If vMinutes >= 60 * 60 Then GetTimeFromMinutes = Format$(Fix(vMinutes / 3600), "00") & ":" & Format$(Fix(vMinutes / 60) Mod 60, "00") & ":" & Format$(Fix(vMinutes Mod 60), "00")
End Function

Function GetMinutesFromTime(vTime As String)
Dim MinS, Hors
Hors = Val(Mid$(vTime, 1, 2))
MinS = Val(Mid(vTime, 4, 2))
GetMinutesFromTime = (Hors * 60) + MinS
End Function

Public Function GetVersion() As String
GetVersion = Format$(App.Major, "0") + "." + Format$(App.Minor, "0") + "." + Format$(App.Revision, "000")
End Function

Public Function Get2Version() As String
Get2Version = Format$(App.Major, "0") + "." + Format$(App.Minor, "0")
End Function

Function PathHead$(FileName As String)
Dim Names As Integer
For Names = Len(FileName) To 1 Step -1
 If Mid$(FileName, Names, 1) = "\" Then
  PathHead$ = Mid$(FileName, 1, (Names) - 1)
  If PathHead$ = "$APPDIR$" Then PathHead$ = App.path
  Exit For
 End If
Next

End Function

Function FileExists(path$) As Boolean
    Dim X As Integer

    X = FreeFile
    Err.Clear
    On Error Resume Next
    Open path$ For Input As X
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close X
    Err.Clear

End Function

Public Function FileHead$(FileName As String)
Dim Names As Integer
For Names = Len(FileName) To 1 Step -1
If Mid$(FileName, Names, 1) = "\" Then FileHead$ = Right$(FileName, Len(FileName) - (Names)): Exit Function
Next
End Function

Public Function LowPath(inPath As String) As String
If Right$(inPath, 1) = "\" Then LowPath = inPath
If Right$(inPath, 1) <> "\" Then LowPath = inPath + "\"
End Function



Public Function Загрузить_Настройки2(Опция As String, Имя_Файла As String, Optional По_Умолчанию = "") As String
Dim CfgLine As String, g As Integer
On Error Resume Next
g = FreeFile
Open Имя_Файла For Input As #g
Do
Line Input #g, CfgLine

Loop While Not EOF(g)
Загрузить_Настройки2 = Format(По_Умолчанию)
Close g
End Function

Public Function GetNetCard(Record As String)
Dim CfgLine As String, g As Integer
g = InStr(Record, " - ")
If g > 0 Then
 CfgLine = Trim(Mid(Record, 1, g))
Else
 CfgLine = Record
End If

GetNetCard = CfgLine
End Function

Function ReadCommand(ByRef GetCommand As String, ByRef GetValue As Boolean)
 If GetValue = True Then ReadCommand = Right$(GetCommand, Len(GetCommand) - 12)
 If GetValue = False Then ReadCommand = Mid$(GetCommand, 1, 11)
End Function

Function FilterName(Text As String) As String

Dim LS, Bs, Variants, Bizer
On Error Resume Next

For LS = 1 To Len(Text)
Bs = Mid$(Text, LS, 1)

 For Variants = 0 To 47
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 For Variants = 91 To 96
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 For Variants = 58 To 63
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 For Variants = 123 To 191
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 
Mid$(Text, LS, 1) = Bs

Next

If Text = "" Then Text = "Unnamed"
FilterName = Text

End Function

Function CBol(Value) As Boolean
If Not Val(Format(Value)) = 0 Then CBol = True: Exit Function
If Format(Value) = "True" Then CBol = True: Exit Function
CBol = False
End Function

Function xStr(Value As Boolean) As String
If Value = True Then xStr = "True" Else xStr = "False"
End Function

Function Bol2Int(inVal As Boolean) As Integer
  Bol2Int = 0
  If inVal = True Then Bol2Int = 1
  If inVal = False Then Bol2Int = 0
End Function

Function CountDubs(Expression As String, Chars As String) As Long
Dim Q As Long: Q = 0
Dim W As Long: W = 0
Do
 Q = InStr(Q + 1, Expression, Chars)
 If Q > 0 Then W = W + 1
Loop Until Q = 0

CountDubs = W
End Function

Sub ExchangeFiles(SI As Integer, DI As Integer, Sources As ListBox)
On Error Resume Next
Dim a, b, ASel As Boolean, BSel As Boolean
a = Sources.List(DI)
b = Sources.List(SI)
ASel = Sources.Selected(DI)
BSel = Sources.Selected(SI)
Sources.List(DI) = b
Sources.List(SI) = a
Sources.Selected(DI) = BSel
Sources.Selected(SI) = ASel

End Sub

' Sub Sleep(TM)
' tm1 = Timer
' Do: DoEvents: Loop While Not Timer >= tm1 + TM
' End Sub

Sub ExtractData(inFile As String, outFile As String, fstByte As Long, lenByte As Long)
On Error Resume Next


Open inFile For Binary As #11
Open outFile For Binary As #12

Const BUFLEN = 32666
LessData = lenByte Mod BUFLEN
OkedData = Fix(lenByte / BUFLEN)
Dim BufferString As String * BUFLEN

For N = fstByte To lenByte + 1 Step BUFLEN
 Get #11, N, BufferString
 Put #12, , BufferString
Next N

Close #11, #12

End Sub

Sub CopyFile(sourcepath As String, destinationpath As String)

    On Error Resume Next

    Dim i As String
    Dim a As Integer, b As Integer
    Dim fullblocks As Integer
    Dim overblock As Integer
    Dim Index As Integer
    
    a = FreeFile
    Open sourcepath For Binary As #a
    fullblocks = Fix(LOF(a) / 32766)
    overblock = LOF(a) Mod 32766
    
    b = FreeFile
    Kill destinationpath
    Open destinationpath For Binary As #b
    
    i = String(32766, 32)
    
    For Index = 1 To fullblocks
        Get #a, , i
        Put #b, , i
    Next Index
    
    If overblock > 0 Then
        
        i = String(overblock, 32)
        Get #a, , i
        Put #b, , i
        
    End If
    
    Close #a, #b
        
End Sub
