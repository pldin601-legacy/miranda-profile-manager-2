VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Additional features"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      Begin VB.ListBox lstOpts 
         Height          =   2310
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   300
         Width           =   5295
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   6
         Top             =   2700
         Width           =   960
      End
      Begin VB.Label lblDescr 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   2940
         Width           =   5295
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3060
      ScaleHeight     =   495
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   3900
      Width           =   2655
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   60
         Width           =   1155
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Complete"
         Default         =   -1  'True
         Height          =   375
         Left            =   180
         TabIndex        =   1
         Top             =   60
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdOpen_Click()
    Me.Hide
    frmPCreate.CreateIt
End Sub



Private Sub Form_Load()
    TranslateForm Me
    lstOpts_Click
End Sub

Sub lstOpts_Click()
    Dim i As Integer
    On Error Resume Next
    For i = 0 To UBound(OptRecord)
        If OptRecord(i).optionName = lstOpts.List(lstOpts.ListIndex) Then
            If Len(Trim(OptRecord(i).optionDescription)) > 0 Then
                lblDescr.Caption = OptRecord(i).optionDescription
            Else
                lblDescr.Caption = TranslateEx("No description found for this option.")
            End If
            Exit Sub
        End If
    Next i
End Sub

Private Sub lstOpts_ItemCheck(item As Integer)
    Dim i As Integer
    On Error Resume Next
    For i = 0 To UBound(OptRecord)
        If OptRecord(i).optionName = lstOpts.List(item) Then
            OptRecord(i).optionEnable = lstOpts.Selected(item)
            Exit For
        End If
    Next i
End Sub
