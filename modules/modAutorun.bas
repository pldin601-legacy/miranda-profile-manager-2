Attribute VB_Name = "modAutorun"
Function IfAutorun(Optional inAppName As String = "", Optional inExeName As String = "") As Boolean
On Error GoTo errores

If inAppName = "" Then inAppName = App.ProductName
If inExeName = "" Then inExeName = App.EXEName + ".exe"

Dim nKey As String: nKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
Dim ret&, vz As String
       
vz = QueryValue(HKEY_LOCAL_MACHINE, nKey, inAppName)
If LCase(vz) = LCase(LowPath(App.path) + inExeName & vbNullChar) Then IfAutorun = True: Exit Function

errores:
IfAutorun = False

End Function

Function IfAutorunUser(Optional inAppName As String = "", Optional inExeName As String = "") As Boolean
On Error GoTo errores

If inAppName = "" Then inAppName = App.ProductName
If inExeName = "" Then inExeName = App.EXEName + ".exe"

Dim nKey As String: nKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
Dim ret&, vz As String
       
vz = QueryValue(HKEY_CURRENT_USER, nKey, inAppName)
If LCase(vz) = LCase(LowPath(App.path) + inExeName & vbNullChar) Then IfAutorunUser = True: Exit Function

errores:
IfAutorunUser = False

End Function

Sub SetAutorun(Optional inAppName As String = "", Optional inExeName As String = "")

Dim nKey As String: nKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
Dim ret&, vz As Long

If inAppName = "" Then inAppName = App.ProductName
If inExeName = "" Then inExeName = App.EXEName + ".exe"

SetKeyValue HKEY_LOCAL_MACHINE, nKey, inAppName, LowPath(App.path) + inExeName, REG_SZ

End Sub

Sub SetAutorunUser(Optional inAppName As String = "", Optional inExeName As String = "")

Dim nKey As String: nKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
Dim ret&, vz As Long

If inAppName = "" Then inAppName = App.ProductName
If inExeName = "" Then inExeName = App.EXEName + ".exe"

SetKeyValue HKEY_CURRENT_USER, nKey, inAppName, LowPath(App.path) + inExeName, REG_SZ

End Sub


Sub KillAutorun(Optional inAppName As String = "")

Dim nKey As String: nKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
Dim ret&, vz As Long
 
If inAppName = "" Then inAppName = App.ProductName
 
DeleteValue HKEY_LOCAL_MACHINE, nKey, inAppName

End Sub

Sub KillAutorunUser(Optional inAppName As String = "")

Dim nKey As String: nKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
Dim ret&, vz As Long
 
If inAppName = "" Then inAppName = App.ProductName
DeleteValue HKEY_CURRENT_USER, nKey, inAppName

End Sub


