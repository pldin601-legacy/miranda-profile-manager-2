Attribute VB_Name = "Manifest"
Option Explicit
Public Col As Long

Public Type INITCOMMONCONTROLSEX
    dwSize As Long
    dwICC As Long
End Type

Public Const ICC_USEREX_CLASSES As Long = &H200

Public Declare Function INITCOMMONCONTROLSEX Lib "comctl32.dll" Alias "InitCommonControlsEx" (ByRef TLPINITCOMMONCONTROLSEX As INITCOMMONCONTROLSEX) As Long

Public Function InitCommonControlsXP() As Boolean
    On Error Resume Next
    Dim ICCEx As INITCOMMONCONTROLSEX
    With ICCEx
        .dwSize = Len(ICCEx)
        .dwICC = ICC_USEREX_CLASSES
    End With
    Call INITCOMMONCONTROLSEX(ICCEx)
    InitCommonControlsXP = CBool(Err = 0)
End Function
