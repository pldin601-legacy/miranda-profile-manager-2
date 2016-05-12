Attribute VB_Name = "imgButton"
Option Explicit

Private Const IMAGE_BITMAP As Long = 0
Private Const BM_SETIMAGE As Long = &HF7&
Private Const LR_LOADFROMFILE As Long = &H10

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long

Dim img As Long


Public Sub SetImageIcon(iconpath As String, btnHandle As Long)
    Dim img As Long
    img = ExtractAssociatedIcon(App.hInstance, iconpath, 0)
    SendMessage btnHandle, BM_SETIMAGE, IMAGE_ICON, img
End Sub
