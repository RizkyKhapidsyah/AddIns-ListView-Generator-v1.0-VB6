Attribute VB_Name = "modAPI"
Option Explicit

Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_SHOWDROPDOWN As Long = &H14F

'---------------------------------------------------------------------------------
Public Sub ShowDropDown(pobjCombo As ComboBox)
'---------------------------------------------------------------------------------

    If pobjCombo.ListCount > 0 Then
        SendMessage pobjCombo.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&
    End If

End Sub


