Attribute VB_Name = "modComboBox"
Option Explicit

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function MoveWindow Lib "USER32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ScreenToClient Lib "USER32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETITEMHEIGHT = &H154

Public Sub ChangeComboDropDownHeight(frm As Form, cbo As ComboBox, iToDisplay As Integer)

    Dim pt As POINTAPI
    Dim rc As RECT
    Dim cWidth As Long
    Dim newHeight As Long
    Dim oldScaleMode As Long
    Dim numItemsToDisplay As Long
    Dim itemHeight As Long

    'how many items should appear in the dropdown?
    numItemsToDisplay = iToDisplay

    'Save the current form scalemode, then
    'switch to pixels
    oldScaleMode = frm.ScaleMode
    frm.ScaleMode = vbPixels

    'the width of the combo, used below
    cWidth = cbo.Width

    'get the system height of a single
    'combo box list item
    itemHeight = SendMessage(cbo.hWnd, CB_GETITEMHEIGHT, 0, ByVal 0)

    'Calculate the new height of the combo box. This
    'is the number of items times the item height
    'plus two. The 'plus two' is required to allow
    'the calculations to take into account the size
    'of the edit portion of the combo as it relates
    'to item height. In other words, even if the
    'combo is only 21 px high (315 twips), if the
    'item height is 13 px per item (as it is with
    'small fonts), we need to use two items to
    'achieve this height.
    newHeight = itemHeight * (numItemsToDisplay + 2)

    'Get the co-ordinates of the combo box
    'relative to the screen
    Call GetWindowRect(cbo.hWnd, rc)
    pt.X = rc.Left
    pt.Y = rc.Top

    'Then translate into co-ordinates
    'relative to the form.
    Call ScreenToClient(frm.hWnd, pt)

    'Using the values returned and set above,
    'call MoveWindow to reposition the combo box
    Call MoveWindow(cbo.hWnd, pt.X, pt.Y, cbo.Width, newHeight, True)

    'Its done, so show the new combo height
    '   Call SendMessage(cbo.hWnd, CB_SHOWDROPDOWN, True, ByVal 0)

    'restore the original form scalemode
    'before leaving
    frm.ScaleMode = oldScaleMode

End Sub

    


