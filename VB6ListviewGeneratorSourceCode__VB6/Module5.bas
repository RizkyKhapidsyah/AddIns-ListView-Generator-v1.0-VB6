Attribute VB_Name = "modRegistry"
Option Explicit

'Public VBInstance As VBIDE.VBE
Private CodeModuleObject As CodeModule
Private VBComponentObject As VBComponent
Private VBFormObject As VBForm
Public VBInstance As VBIDE.VBE

Private Declare Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Long) As Long
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17

Private Const strKey As String = "HKEY_CURRENT_USER\Software\"

Dim oWSHShell As WshShell

Private Function RegWrite(sKey As String, sFilepath As String) As Boolean
    On Error GoTo Err
    Set oWSHShell = New WshShell
    oWSHShell.RegWrite sKey, sFilepath
    Set oWSHShell = Nothing
    RegWrite = True
    Exit Function
Err:
    RegWrite = False
End Function

Private Function RegDelete(sKey As String) As Boolean
    On Error GoTo Err
    Set oWSHShell = New WshShell
    oWSHShell.RegDelete sKey
    Set oWSHShell = Nothing
    RegDelete = True
    Exit Function
Err:
    RegDelete = False
End Function

Private Function RegRead(strKey)
    On Error Resume Next
    Set oWSHShell = New WshShell
    RegRead = oWSHShell.RegRead(strKey)
    Set oWSHShell = Nothing
End Function

Public Function SavePositionsInRegistry(frm As Form)

    If frm.WindowState = vbMaximized Or frm.WindowState = vbMinimized Then Exit Function

    Dim KeyReg As String, k As String

    KeyReg = strKey & App.Title & "\" & frm.Name & "\"
    RegWrite KeyReg & "FormLeft", frm.Left
    RegWrite KeyReg & "FormTop", frm.Top
    RegWrite KeyReg & "FormWidth", frm.Width
    RegWrite KeyReg & "FormHeight", frm.Height

End Function

Public Function GetPositionsFromRegistry(frm As Form)

    If frm.WindowState = vbMaximized Or frm.WindowState = vbMinimized Then Exit Function

    Dim KeyReg As String
    Dim ileft, itop, iwidth, iheight
    Dim lCenterLeft As Long, lCenterTop As Long
    GetFormCenter frm, lCenterLeft, lCenterTop
    KeyReg = strKey & App.Title & "\" & frm.Name & "\"

    ileft = IIf(IsEmpty(RegRead(KeyReg & "FormLeft")), lCenterLeft, RegRead(KeyReg & "FormLeft"))
    itop = IIf(IsEmpty(RegRead(KeyReg & "FormTop")), lCenterTop, RegRead(KeyReg & "FormTop"))
    iwidth = IIf(IsEmpty(RegRead(KeyReg & "FormWidth")), frm.Width, RegRead(KeyReg & "FormWidth"))
    iheight = IIf(IsEmpty(RegRead(KeyReg & "FormHeight")), frm.Height, RegRead(KeyReg & "FormHeight"))

    frm.Move ileft, itop, iwidth, iheight

End Function

Private Function GetFormCenter(frm As Form, lLeft As Long, lTop As Long)
    With frm
        lLeft = (Screen.TwipsPerPixelX * (GetSystemMetrics(SM_CXFULLSCREEN) / 2)) - (.Width / 2)
        lTop = (Screen.TwipsPerPixelY * (GetSystemMetrics(SM_CYFULLSCREEN) / 2)) - (.Height / 2)
    End With
End Function

Public Function AddOptionExplicit() As String
    If IsExistOptionExplicit Then
        AddOptionExplicit = vbNullString
    Else
        AddOptionExplicit = "Option Explicit 'Add by Project Builder 2.0" & vbCrLf & vbCrLf
    End If
End Function

Private Function IsExistOptionExplicit() As Boolean
    Set VBComponentObject = VBInstance.SelectedVBComponent
    Set CodeModuleObject = VBComponentObject.CodeModule
    IsExistOptionExplicit = CodeModuleObject.Find("Option Explicit", 1, 1, -1, -1, True, True)
End Function

Public Sub DeleteRegForm(frmName As String)
    RegDelete strKey & VBInstance.ActiveVBProject.Name & "\" & frmName & "\"
End Sub

Public Function FormCenter(frm As Form)
    Dim lLeft As Long
    Dim lTop As Long
    With frm
        lLeft = (Screen.TwipsPerPixelX * (GetSystemMetrics(SM_CXFULLSCREEN) / 2)) - (.Width / 2)
        lTop = (Screen.TwipsPerPixelY * (GetSystemMetrics(SM_CYFULLSCREEN) / 2)) - (.Height / 2)
    End With
    frm.Move lLeft, lTop
End Function

Public Function AddModule(ModulName As String, Optional strCode As String) As Boolean

    Dim newModule As VBComponent

    On Error GoTo ErrHandler

    Set newModule = VBInstance.ActiveVBProject.VBComponents.Add(vbext_ct_StdModule)
    With newModule
        .Name = ModulName
        .CodeModule.AddFromString strCode
    End With
    Exit Function

ErrHandler:

    MsgBox Err.Description

End Function

Public Function IsOCXRegistered(sGUID As String) As Boolean
    On Error GoTo ErrHandler
    Set oWSHShell = New WshShell
    oWSHShell.RegRead "HKEY_CLASSES_ROOT\TypeLib\" & sGUID & "\"
    Set oWSHShell = Nothing
    IsOCXRegistered = True
    Exit Function
ErrHandler:
    Debug.Print Err.Description & " " & sGUID
    Set oWSHShell = Nothing
    IsOCXRegistered = False
End Function


