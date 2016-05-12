VERSION 5.00
Begin VB.Form frmName 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creating profile"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4545
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
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
      Left            =   -120
      ScaleHeight     =   915
      ScaleWidth      =   5055
      TabIndex        =   5
      Top             =   -60
      Width           =   5115
      Begin VB.PictureBox pMiniLogo 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   3540
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   62
         TabIndex        =   6
         Top             =   120
         Width           =   930
      End
      Begin VB.Label Label2 
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
         TabIndex        =   8
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type name for the new profile"
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
         TabIndex        =   7
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   60
      ScaleHeight     =   1695
      ScaleWidth      =   4575
      TabIndex        =   1
      Top             =   1020
      Width           =   4575
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Next >>"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   1980
         TabIndex        =   4
         Top             =   1260
         Width           =   1155
      End
      Begin VB.CommandButton cmdDelete 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3180
         TabIndex        =   3
         Top             =   1260
         Width           =   1155
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   420
         TabIndex        =   0
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Profile name:"
         Height          =   195
         Left            =   420
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDelete_Click()
    txtName.Text = ""
    Me.Hide
End Sub

Private Sub cmdOpen_Click()
    If Len(Trim(txtName.Text)) > 0 Then
        If FileExists(LowPath(AbsolutePath(Folders.PMANAGER_PROFILES)) & Trim(txtName.Text) & ".dat") = False Then
            Me.Hide
        Else
            MsgBox TranslateEx("Can`t create profile with this name! Profile already exists!"), vbCritical
        End If
    End If
End Sub



Private Sub Form_Load()
    TranslateForm Me
    pMiniLogo.Picture = LoadPicture(LowPath(AbsolutePath(Folders.PMANAGER_GRAPH)) & "logo_mini.jpg")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    txtName.Text = ""
End Sub

Private Sub txtName_Change()
    If Len(Trim(txtName.Text)) > 0 Then
        cmdOpen.Enabled = True
    Else
        cmdOpen.Enabled = False
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    Select Case Chr(KeyAscii)
    Case "\", "/", ":", "*", "?", """", "<", ">", "|"
        KeyAscii = 0
    End Select
End Sub
