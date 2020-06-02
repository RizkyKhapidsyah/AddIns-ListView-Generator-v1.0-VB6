VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------------------------------------------------
' Coder: Asep Hibban
'--------------------------------------------------------------------------------------------------

Private mCol As Collection

Public Function Add(ctl As Control, _
    Optional TipText As String = "Text To Show", _
    Optional Title As String = "Title", _
    Optional eStyle As toolStyleEnum, _
    Optional Icon As Integer, _
    Optional strKey As String) As clsTooltip

    Dim objNewMember As clsTooltip
    Set objNewMember = New clsTooltip

    If eStyle = styleBalloon Then
        objNewMember.CreateBalloon ctl, TipText, Title, Icon
    Else
        objNewMember.CreateTip ctl, TipText, Title, Icon
    End If

    If Len(strKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, strKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsTooltip
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    Me.Item(vntIndexKey).Remove 'detroy TipText Windows
    mCol.Remove vntIndexKey 'remove from collections
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Sub Clear()
    Dim i As Long
    For i = 1 To mCol.Count
        Remove mCol.Count
    Next
End Sub

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Clear
    Set mCol = Nothing
End Sub

    



