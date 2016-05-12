Attribute VB_Name = "Folders"

Public Const PMANAGER_SYSDIR = "PManager"
Public Const PMANAGER_PATTERNS = PMANAGER_SYSDIR & "\Patterns"
Public Const PMANAGER_GRAPH = PMANAGER_SYSDIR & "\Logo"

Dim ProfileCache As String
Dim CreateOK As Boolean

Public Type PMOptions
    optionFilename As String
    optionName As String
    optionDescription As String
    optionEnable As Boolean
End Type

Public OptRecord() As PMOptions
Public PrevListIndex As Integer

Public pm_root As String
Public pm_sysdir As String
Public pm_patterns As String
Public pm_pictures As String

Public pm_mirandaexe As String
Public pm_profiles As String


Sub InitPathes()
    pm_root = App.path
    pm_sysdir = pm_root + "\pmanager"
    pm_patterns = pm_sysdir + "\patterns"
    pm_pictures = pm_sysdir + "\logo"
    pm_mirandaexe = DecPath(GetSettingINI("Settings", "MirandaEXE", "%pmroot%\miranda32.exe", pm_sysdir & "\pmanager.ini"))
    pm_profiles = DecPath(GetSettingINI("Settings", "Profiles", "%pmroot%\profiles", pm_sysdir & "\pmanager.ini"))
End Sub

Function DecPath(inpath As String) As String
    Dim tmppath As String
    tmppath = Replace(inpath, "%pmroot%", App.path)
    DecPath = tmppath
End Function

Public Function MirandaEXE() As String
Dim MPath As String
MirandaEXE = LowPath(MirandaPATH) + GetSettingINI("Settings", "MirandaEXE", "miranda32.exe", LowPath(AbsolutePath(PMANAGER_SYSDIR)) & "pmanager.ini")
End Function

Public Function MirandaINI() As String
Dim MPath As String
MirandaINI = LowPath(MirandaPATH) + "mirandaboot.ini"
End Function

Public Function MirandaPATH() As String
Dim MPath As String
MirandaPATH = LowPath(App.path) + LowPath(GetSettingINI("Settings", "MirandaPath", "", LowPath(AbsolutePath(PMANAGER_SYSDIR)) & "pmanager.ini"))
End Function

Public Function ManagerINI() As String
Dim MPath As String
ManagerINI = LowPath(AbsolutePath(PMANAGER_SYSDIR)) & "pmanager.ini"
End Function

Public Function AbsolutePath(relative As String) As String
    
    If Mid(relative, 2, 1) = ":" Then
        AbsolutePath = relative
    Else
        AbsolutePath = IIf(Right(App.path, 1) = "\", App.path, App.path & "\") & relative
    End If

End Function

Public Function GetDefaultProfile() As String

    On Error Resume Next
    GetDefaultProfile = GetSettingINI("Settings", "DefaultProfile", "", ManagerINI)

End Function

Public Function SaveDefaultProfile(profile As String)

    On Error Resume Next
    Call SaveSettingINI("Settings", "DefaultProfile", profile, ManagerINI)
    Call SaveSettingINI("Database", "DefaultProfile", profile, MirandaINI)
    
End Function

Public Function PMANAGER_PROFILES() As String
    
    On Error Resume Next
    

        ProfileCache = LowPath(MirandaPATH) + GetSettingINI("Database", "ProfileDir", ".", MirandaINI)
        PMANAGER_PROFILES = ProfileCache

    
End Function


Function RequestProfileName(Optional default As String) As String
    frmName.txtName.Text = default
    If Len(default) > 0 Then
        frmName.txtName.SelStart = 0
        frmName.txtName.SelLength = Len(default)
    End If
    frmName.Show vbModal
    RequestProfileName = frmName.txtName.Text
    Unload frmName
End Function
