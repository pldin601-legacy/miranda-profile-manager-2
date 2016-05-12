VERSION 5.00
Begin VB.Form mdlSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Miranda Profile Manager"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "mdlSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   277
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Profile settings"
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   3915
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   3735
         TabIndex        =   13
         Top             =   240
         Width           =   3735
         Begin VB.CheckBox chkRestore 
            Caption         =   "Restore original settings"
            Height          =   435
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   3615
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   -60
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   1
      Top             =   -60
      Width           =   4275
      Begin VB.Image Image1 
         Height          =   1440
         Left            =   1320
         Top             =   120
         Width           =   1860
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2475
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   3915
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   120
         ScaleHeight     =   1755
         ScaleWidth      =   3735
         TabIndex        =   2
         Top             =   240
         Width           =   3735
         Begin VB.CommandButton Command1 
            Caption         =   "New..."
            Height          =   315
            Left            =   2700
            TabIndex        =   11
            ToolTipText     =   "Create new profile"
            Top             =   240
            Width           =   915
         End
         Begin VB.ComboBox cmdProfiles 
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   240
            Width           =   2595
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "&Open"
            Default         =   -1  'True
            Enabled         =   0   'False
            Height          =   375
            Left            =   60
            TabIndex        =   7
            Top             =   1380
            Width           =   1695
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            TabIndex        =   6
            Top             =   1380
            Width           =   1695
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Automatically start with &Windows"
            Height          =   195
            Left            =   60
            TabIndex        =   4
            Top             =   1020
            Width           =   3615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Don`t &show profile manager"
            Height          =   195
            Left            =   60
            TabIndex        =   3
            Top             =   720
            Width           =   3615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Profile:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   9
            Top             =   0
            Width           =   585
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "About profile manager"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         MouseIcon       =   "mdlSelect.frx":57E2
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   2100
         Width           =   1605
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profile Settings..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   120
      MouseIcon       =   "mdlSelect.frx":63A4
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   4380
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://miranda-planet.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   2400
      MouseIcon       =   "mdlSelect.frx":6F66
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4380
      Width           =   1590
   End
End
Attribute VB_Name = "mdlSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const WND_MIN = 5025
Const WND_MAX = 6645

Sub EnumProfiles()
    cmdProfiles.Clear
    Dim i As String
    Dim temppath As String
    i = Dir(LowPath(Folders.pm_profiles) & "*", vbDirectory)
    If Len(i) > 0 Then
        Do While Len(i) > 0
            If Len(i) > 0 And i <> "." And i <> ".." Then
                temppath = Folders.pm_profiles & "\" & i & "\" & i & ".dat"
                If FileExists(temppath) = True Then
                    cmdProfiles.AddItem i
                End If
            End If
            i = Dir
        Loop
    End If
    cmdProfiles_Change
End Sub

Sub SelectNext(Index As Integer)

    On Error Resume Next
    
    If Index + 1 <= cmdProfiles.ListCount Then
        cmdProfiles.ListIndex = Index
    ElseIf Index + 1 = cmdProfiles.ListCount Then
    Else
        cmdProfiles.ListIndex = Index - 1
    End If
    
    
End Sub

Sub SelectDefault()
    
    On Error Resume Next
    
    Dim i As String
    Dim pn As String
    
    i = GetDefaultProfile
    pn = LowPath(Folders.pm_profiles) & i & "\" & i & ".dat"
    
    If Len(Trim(i)) > 0 Then
        If FileExists(pn) = False Then cmdProfiles.Text = "": Exit Sub
        If Check1.Value > 0 And InStr(UCase(Command), "/FORCESHOW") = 0 And Not IsMirandaRunning Then
            StartMiranda i
        Else
            cmdProfiles.Text = i
        End If
    End If
    
End Sub

Function IsMirandaRunning() As Boolean
    IsMirandaRunning = IsProgramRunning(MirandaEXE)
End Function

Private Sub cmdDelete_Click()

    Call DeleteProfile(cmdProfiles.Text, cmdProfiles.ListIndex)

End Sub

Private Sub cmdOpen_Click()

    Call OpenProfile(cmdProfiles.Text)

End Sub

Private Sub cmdProfiles_Change()

    Dim i As String
    i = LowPath(AbsolutePath(Folders.PMANAGER_PROFILES)) & cmdProfiles.Text & ".dat"
    
    For u = 0 To cmdProfiles.ListCount - 1
        If LCase(cmdProfiles.List(u)) = LCase(cmdProfiles.Text) Then _
            cmdProfiles.ListIndex = u: _
            Exit For

    Next u
    
    If Trim(cmdProfiles.Text) = "" Then
        cmdOpen.Enabled = False
        cmdDelete.Enabled = False
    Else
        cmdOpen.Enabled = True
        If FileExists(i) Then
            cmdDelete.Enabled = True
        Else
            cmdDelete.Enabled = False
        End If
    End If

End Sub


Sub OpenProfile(profilename As String)

    Dim i As String
    Dim sn As String
    
    i = LowPath(AbsolutePath(Folders.PMANAGER_PROFILES)) & profilename & ".dat"
    
    If FileExists(i) Then
'        If chkRestore.Value > 0 Then
'            If MsgBox(TranslateEx("Are you sure want to restore original profile settings?\\nHistory, contacts and account information will not be erased."), vbQuestion + vbYesNo) = vbYes Then
'                sn = LowPath(AbsolutePath(Folders.PMANAGER_PATTERNS)) & GetSettingINI("Profiles", profilename, "", LowPath(AbsolutePath(PMANAGER_SYSDIR)) & "pmanager.ini") & ".ini"
'                If FileExists(sn) Then
'                    CopyFile sn, LowPath(App.path) & "autoexec_restore.ini"
'                Else
'                    MsgBox TranslateEx("Can`t restore settings. Settings file not found!"), vbCritical
'                    Exit Sub
'                End If
'            Else
'                Exit Sub
'            End If
'        End If
        SaveDefault profilename
        SaveSettingsX
        StartMiranda profilename
    Else
        If MsgBox(TranslateEx("Profile not exists. Want to create new?"), vbQuestion + vbYesNo) = vbYes Then
            CreateProfile profilename
        End If
    End If
    
        
    
End Sub


Sub SaveDefault(profilename As String)
    
   
    Call SaveDefaultProfile(profilename)

End Sub

Private Sub cmdProfiles_Click()
    cmdProfiles_Change
End Sub

Private Sub cmdProfiles_KeyPress(KeyAscii As Integer)
    Select Case Chr(KeyAscii)
    Case "\", "/", ":", "*", "?", """", "<", ">", "|"
        KeyAscii = 0
    End Select
End Sub

Private Sub Command1_Click()
    
    On Error Resume Next
    
    CreateProfile
    
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    cmdProfiles.SetFocus
End Sub

Private Sub Form_Initialize()
    InitCommonControlsXP
End Sub

Private Sub Form_Load()

    If App.PrevInstance Then End
    CacheTranslateEx FindLNG(App.path)
    ' is miranda present?

    If Not FileExists(MirandaEXE) Or Not FileExists(MirandaINI) Then
        MsgBox TranslateEx("Miranda IM is not installed or installed not correctly.\\nReinstall please."), vbCritical
        'End
    End If
    
    Call Folders.InitPathes
    Call TranslateForm(Me)
    
    ' Me.Height = WND_MIN
    Image1.Picture = LoadPicture(LowPath(Folders.pm_pictures) & "logo.jpg")
    Image1.Left = Picture1.ScaleLeft + (Me.ScaleWidth / 2) - (Image1.Width / 2)
    Image1.Top = (Picture1.ScaleHeight / 2) - (Image1.Height / 2)
    
    Check1.Value = CInt(GetSettingINI("Settings", "ShowManager", "0", LowPath(AbsolutePath(PMANAGER_SYSDIR)) & "pmanager.ini"))
    Check2.Value = CInt(-IfAutorunUser("Miranda IM"))

    EnumProfiles
    SelectDefault
    
    
End Sub

Sub SaveAll(ppname As String)
                mdlSelect.SaveDefault ppname
                mdlSelect.SaveSettingsX

End Sub

Sub StartMiranda(profilename As String)
    On Error Resume Next
    Dim pmUseExtension As Boolean
    Me.Hide
    pmUseExtension = CBol(GetSettingINI("Options", "UseExtension", "True", LowPath(AbsolutePath(PMANAGER_SYSDIR)) & "pmanager.ini"))
    RunWEB MirandaEXE, profilename & IIf(pmUseExtension, ".dat", "")
    End
End Sub

Sub SaveSettingsX()
    
    On Error Resume Next
    
    If Check2.Value = 1 Then
        If Not IfAutorunUser("Miranda IM") Then SetAutorunUser ("Miranda IM")
    Else
        If IfAutorunUser("Miranda IM") Then KillAutorunUser ("Miranda IM")
    End If

    Call SaveSettingINI("Settings", "ShowManager", CStr(Check1.Value), LowPath(AbsolutePath(PMANAGER_SYSDIR)) & "pmanager.ini")

End Sub

Sub DeleteProfile(profilename As String, listindx As Integer)
    
    On Error Resume Next
    
    Dim i As String, j As Integer
    i = LowPath(AbsolutePath(Folders.PMANAGER_PROFILES)) & profilename & ".dat"
    
    If FileExists(i) Then
        If MsgBox(TranslateEx("Are you sure want to delete selected profile?"), vbQuestion + vbYesNo) = vbYes Then
            Kill i
            If Err Then
                Call MsgBox(TranslateEx("Profile could not be deleted!"), vbCritical)
            End If
            EnumProfiles
            SelectNext listindx
        End If
    Else
        Call MsgBox(TranslateEx("Profile not exists!"), vbExclamation)
    End If
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Label1_Click()
    RunWEB "http://miranda-planet.com"
End Sub

Private Sub Label2_Click()
    frmAboutPM.Show vbModal, Me
End Sub

Private Sub Label3_Click()
    cmdProfiles.SetFocus
End Sub


Function FindLNG(path As String) As String
    
    On Error Resume Next
    
    Dim FN As String
    Dim pn As String
    
    pn = LowPath(AbsolutePath(Folders.PMANAGER_SYSDIR))
    
    FN = Dir(pn & "pman_*.lng")
    
    If Len(FN) > 0 Then FindLNG = pn & FN
    
End Function

Sub CreateProfile(Optional profilename As String)

    On Error Resume Next
    
    Dim phn As String
    

    
    If Len(profilename) = 0 Then
        phn = IIf(Not FileExists(LowPath(AbsolutePath(Folders.PMANAGER_PROFILES)) & cmdProfiles.Text & ".dat"), cmdProfiles.Text, "")
        profilename = RequestProfileName(phn)
    End If
    
    If Len(profilename) = 0 Then Exit Sub
    
    Dim i As String
    i = LowPath(AbsolutePath(Folders.PMANAGER_PROFILES)) & profilename & ".dat"
    
    If Not FileExists(i) Then
        frmPCreate.Main profilename
    Else
        MsgBox TranslateEx("Can`t create profile with this name! It already exists!")
    End If
    
End Sub

Private Sub Label4_Click()
    Me.Height = IIf(Me.Height <> WND_MAX, WND_MAX, WND_MIN)
End Sub


Sub uuu()
    Dim lne As String
    Open LowPath(App.path) & "kolibris.ini" For Input As #1
    
    Do
        Line Input #1, lne
        If InStr(lne, "=") > 0 Then
            Select Case Mid(lne, InStr(lne, "=") + 1, 1)
            Case "b", "w", "d", "l", "s", "n"
            Case Else
                Debug.Print lne
            End Select
        End If
    Loop While Not EOF(1)
    
    Close #1
    
End Sub
