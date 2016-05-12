VERSION 5.00
Begin VB.Form frmAboutPM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About programm"
   ClientHeight    =   2700
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAboutPM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3060
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   2160
      Width           =   1635
      Begin VB.CommandButton OKButton 
         Caption         =   "&OK"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   60
         Width           =   1215
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
      Height          =   2175
      Left            =   -60
      ScaleHeight     =   2115
      ScaleWidth      =   5295
      TabIndex        =   0
      Top             =   -60
      Width           =   5355
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://woobind.org.ua"
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
         Left            =   2940
         MouseIcon       =   "frmAboutPM.frx":57E2
         MousePointer    =   99  'Custom
         TabIndex        =   8
         ToolTipText     =   "Visit website"
         Top             =   1200
         Width           =   1620
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "woobind@ukr.net"
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
         Left            =   240
         MouseIcon       =   "frmAboutPM.frx":63A4
         MousePointer    =   99  'Custom
         TabIndex        =   7
         ToolTipText     =   "Write e-mail"
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4200
         Picture         =   "frmAboutPM.frx":6F66
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2008 Roman Gemini. This software is released under the terms if the GNU General Public License."
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0 build ¹8 Unicode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   900
         Width           =   2460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Profile Manager"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Miranda IM"
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
         TabIndex        =   1
         Top             =   180
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmAboutPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    TranslateForm Me
    Label3.Caption = TranslateEx("Version") & " " & Format(App.Major, "0") & "." & Format(App.Minor, "0") & " " & TranslateEx("build") & " #" & Format(App.Revision, "0") & " ANSI"
End Sub

Private Sub Label5_Click()
    RunWEB ("mailto:woobind@ukr.net")
End Sub

Private Sub Label6_Click()
    RunWEB ("http://woobind.org.ua/forum/index.php?board=13.0")
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
