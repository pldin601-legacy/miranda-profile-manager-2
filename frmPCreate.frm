VERSION 5.00
Begin VB.Form frmPCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create new profile"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4860
      ScaleHeight     =   495
      ScaleWidth      =   2715
      TabIndex        =   6
      Top             =   5100
      Width           =   2715
      Begin VB.CommandButton cmdDelete 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   60
         Width           =   1155
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Create"
         Default         =   -1  'True
         Height          =   375
         Left            =   180
         TabIndex        =   7
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Preview"
      Height          =   4035
      Left            =   3180
      TabIndex        =   5
      Top             =   1020
      Width           =   4275
      Begin VB.PictureBox pView 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2715
         Left            =   120
         ScaleHeight     =   2715
         ScaleWidth      =   4035
         TabIndex        =   10
         Top             =   240
         Width           =   4035
         Begin VB.PictureBox pPic 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   8115
            Left            =   2820
            ScaleHeight     =   8115
            ScaleWidth      =   10080
            TabIndex        =   13
            Top             =   1980
            Visible         =   0   'False
            Width           =   10080
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   3675
         Left            =   120
         ScaleHeight     =   3675
         ScaleWidth      =   4095
         TabIndex        =   9
         Top             =   240
         Width           =   4095
         Begin VB.Label lPDesc 
            Height          =   615
            Left            =   0
            TabIndex        =   12
            Top             =   3120
            Width           =   3975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Description"
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
            Left            =   0
            TabIndex        =   11
            Top             =   2880
            Width           =   960
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Available"
      Height          =   4035
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   2895
      Begin VB.ListBox lstProfiles 
         Height          =   3570
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   -60
      ScaleHeight     =   915
      ScaleWidth      =   8295
      TabIndex        =   1
      Top             =   -60
      Width           =   8355
      Begin VB.PictureBox pMiniLogo 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   6500
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   62
         TabIndex        =   14
         Top             =   100
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select what kind of profile you want to create"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   3840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create profile"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   180
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmPCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ppname As String

    


Sub SetProfileName(pname As String)
    ppname = pname
    Label1.Caption = TranslateEx("Create profile") & " """ & pname & """"
End Sub

Private Sub cmdDelete_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    If frmOptions.lstOpts.ListCount > 0 Then
        frmOptions.Show vbModal, Me
    Else
        CreateIt
    End If
End Sub

Function SelectedPattern() As String
    If lstProfiles.ListIndex > -1 Then SelectedPattern = lstProfiles.List(lstProfiles.ListIndex) Else selectedprofile = ""
End Function

Sub LoadOptions()
    
    On Error Resume Next
    
    Dim FN As String
    Dim indx As Integer
    Dim boud As Integer
    Dim npt As Integer
    ' Scan pattern's options
    boud = -1
    FN = Dir(LowPath(AbsolutePath(Folders.PMANAGER_PATTERNS)) & SelectedPattern & "_options\autoexec_*.ini")
    If Trim(FN) > "" Then
        boud = boud + 1
        ReDim OptRecord(boud) As PMOptions
        OptRecord(boud).optionFilename = LowPath(AbsolutePath(Folders.PMANAGER_PATTERNS)) & SelectedPattern & "_options\" & FN

        Do
            FN = Dir
            If FN > "" Then
                boud = boud + 1
                ReDim Preserve OptRecord(0 To boud) As PMOptions
                OptRecord(boud).optionFilename = LowPath(AbsolutePath(Folders.PMANAGER_PATTERNS)) & SelectedPattern & "_options\" & FN
            End If
        Loop While Not FN = ""
    End If
    
    ' Scan global options
    FN = Dir(LowPath(AbsolutePath(Folders.PMANAGER_PATTERNS)) & "options\autoexec_*.ini")
    
    If FN > "" Then
        boud = boud + 1
        ReDim Preserve OptRecord(0 To boud) As PMOptions
        OptRecord(boud).optionFilename = LowPath(AbsolutePath(Folders.PMANAGER_PATTERNS)) & "options\" & FN
    
        Do
            FN = Dir
            If FN > "" Then
                boud = boud + 1
                ReDim Preserve OptRecord(0 To boud) As PMOptions
                OptRecord(boud).optionFilename = LowPath(AbsolutePath(Folders.PMANAGER_PATTERNS)) & "options\" & FN
            End If
        Loop While Not FN = ""
    
    End If
    
    
    Dim tmpName As String, tmpDef As Boolean
    ' Analysing options
    npt = 0
    frmOptions.lstOpts.Clear
    frmOptions.lblDescr.Caption = ""
    
    For indx = 0 To boud
        tmpName = GetSettingINI("PManager", "OptionName", "", OptRecord(indx).optionFilename & ".pm")
        If tmpName > "" Then
            OptRecord(indx).optionName = tmpName
            OptRecord(indx).optionDescription = GetSettingINI("PManager", "OptionDescription", "", OptRecord(indx).optionFilename & ".pm")
            OptRecord(indx).optionEnable = CBol(GetSettingINI("PManager", "OptionDefault", "False", OptRecord(indx).optionFilename & ".pm"))
            frmOptions.lstOpts.AddItem tmpName, npt
            frmOptions.lstOpts.Selected(npt) = OptRecord(indx).optionEnable
            npt = npt + 1
        End If
    Next indx
    
    If frmOptions.lstOpts.ListCount > 0 Then
        frmOptions.lstOpts.ListIndex = 0
        Call frmOptions.lstOpts_Click
        cmdShOpt.Enabled = True
    Else
        cmdShOpt.Enabled = False
    End If
    
End Sub

Sub Main(profilename As String)

    If profilename = "" Then Unload Me
    SetProfileName profilename
    pMiniLogo.Picture = LoadPicture(LowPath(AbsolutePath(Folders.PMANAGER_GRAPH)) & "logo_mini.jpg")
    TranslateForm Me
    EnumPatterns
    LoadOptions
    SkipOnDemand
    Me.Show vbModal, mdlSelect
    
End Sub

Sub LoadPreview(pname As String)
    
    On Error Resume Next
    
    Dim pFolder As String
    pFolder = LowPath(AbsolutePath(Folders.PMANAGER_PATTERNS))
    
    If FileExists(pFolder & pname & "_prevw.jpg") Then
        pPic.Picture = LoadPicture(pFolder & pname & "_prevw.jpg")
        SetStretchBltMode pView.hdc, STRETCH_HALFTONE
        StretchBlt pView.hdc, 0, 0, pView.Width, pView.Height, pPic.hdc, 0, 0, pPic.Width, pPic.Height, vbSrcCopy
    Else
        pView.Cls
    End If
    
    pView.Refresh
    
    If FileExists(pFolder & pname & "_descr.txt") Then
        lPDesc.Caption = LoadFile(pFolder & pname & "_descr.txt")
    Else
        lPDesc.Caption = ""
    End If
    
    
End Sub

Sub EnumPatterns()

    On Error Resume Next
    Dim iDir As String, iFile As String
    lstProfiles.Clear
    
    iDir = LowPath(AbsolutePath(Folders.PMANAGER_PATTERNS))
    iFile = Dir(iDir & "*.pp")
    If Len(iFile) > 0 Then
        lstProfiles.AddItem Mid(iFile, 1, InStr(iFile, ".") - 1)
        Do
            iFile = Dir()
            If Len(iFile) > 0 Then lstProfiles.AddItem Mid(iFile, 1, InStr(iFile, ".") - 1)
        Loop While Not Trim(iFile) = ""
    End If
    
    If lstProfiles.ListCount > 0 Then
        lstProfiles.ListIndex = 0
        cmdOpen.Enabled = True
        
    Else
        pView.Cls
        lPDesc.Caption = ""
        cmdOpen.Enabled = False
    End If
    
    LoadOptions
    
End Sub






Private Sub Form_Load()
ReDim OptRecord(0) As PMOptions
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Unload frmOptions
End Sub

Private Sub lstProfiles_Click()
    LoadPreview SelectedPattern
    If lstProfiles.ListIndex <> PrevListIndex Then LoadOptions
    PrevListIndex = lstProfiles.ListIndex
End Sub

Sub SkipOnDemand()

    On Error Resume Next
    
    If frmOptions.lstOpts.ListCount = 0 And lstProfiles.ListCount = 1 Then
        lstProfiles.ListIndex = 0
        CreateIt
    End If
    
End Sub
Function CountSelectedOptions() As Integer
    
    Dim N, M As Integer
    For N = 0 To frmOptions.lstOpts.ListCount - 1
        If frmOptions.lstOpts.Selected(N) = True Then M = M + 1
    Next N

    CountSelectedOptions = M

End Function


Sub CreateIt()
    
    Dim u As String, i As String, pFolder As String
    
    
    
    If CountSelectedOptions > 0 Then
        If IsProgramRunning(MirandaEXE) Then
            MsgBox TranslateEx("Can`t create new profile with options while Miranda is running.\\nClose Miranda and try again.")
            Exit Sub
        End If
    End If
    
    pFolder = LowPath(AbsolutePath(Folders.PMANAGER_PATTERNS))
    
    i = LowPath(AbsolutePath(Folders.PMANAGER_PROFILES)) & ppname & ".dat"
    
    MkDirEx AbsolutePath(Folders.PMANAGER_PROFILES)
    
        If FileExists(pFolder & lstProfiles.List(lstProfiles.ListIndex) & ".pp") Then
            CopyFile pFolder & lstProfiles.List(lstProfiles.ListIndex) & ".pp", i
            Dim indx As Integer
            For indx = 0 To UBound(OptRecord)
                If OptRecord(indx).optionEnable Then
                    CopyFile OptRecord(indx).optionFilename, LowPath(MirandaPATH) & FileHead(OptRecord(indx).optionFilename)
                End If
            Next indx
            
            ' Save profile parent
            ' Call SaveSettingINI("Profiles", ppname, lstProfiles.List(lstProfiles.ListIndex), LowPath(AbsolutePath(PMANAGER_SYSDIR)) & "pmanager.ini")
            ' Save setting and GO
            mdlSelect.SaveAll ppname
            mdlSelect.StartMiranda ppname
            

        Else
            MsgBox TranslateEx("Can`t create new profile. Pattern file not found."), vbCritical
            Main ppname
        End If
    
End Sub

Private Sub lstProfiles_DblClick()
    Call cmdOpen_Click
End Sub
