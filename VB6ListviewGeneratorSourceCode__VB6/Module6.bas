Attribute VB_Name = "modCode"
Option Explicit

Public Function strCodeFormLookUp() As String
    Dim sMsg As String
    sMsg = sMsg & AddOptionExplicit
    sMsg = sMsg & "'---------------------------------------------------------------------------------------------" & vbCrLf
    sMsg = sMsg & "' http://khoiriyyah.blogspot.com" & vbCrLf
    sMsg = sMsg & "'---------------------------------------------------------------------------------------------" & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    sMsg = sMsg & "Public blnDontSetData As Boolean" & vbCrLf
    sMsg = sMsg & "Public ID As Variant" & vbCrLf
    sMsg = sMsg & "Public MODE As Integer" & vbCrLf
    sMsg = sMsg & "Private Sub cmdActions_Click(Index As Integer)" & vbCrLf
    sMsg = sMsg & "    On Error GoTo ErrHandler" & vbCrLf
    sMsg = sMsg & "    Select Case Index" & vbCrLf
    sMsg = sMsg & "        Case COMMAND_FORM_SAVE" & vbCrLf
    sMsg = sMsg & "            If MODE = 0 Then" & vbCrLf
    sMsg = sMsg & "                 InsertData" & vbCrLf
    sMsg = sMsg & "            ElseIf MODE = 1 Then" & vbCrLf
    sMsg = sMsg & "                 UpdateData" & vbCrLf
    sMsg = sMsg & "            End If" & vbCrLf
    sMsg = sMsg & "            Me.Hide" & vbCrLf
    sMsg = sMsg & "        Case COMMAND_FORM_CANCEL" & vbCrLf
    sMsg = sMsg & "            blnDontSetData = True" & vbCrLf
    sMsg = sMsg & "            Me.Hide" & vbCrLf
    sMsg = sMsg & "    End Select" & vbCrLf
    sMsg = sMsg & "    Exit Sub" & vbCrLf
    sMsg = sMsg & "ErrHandler:" & vbCrLf
    sMsg = sMsg & "    MsgBox " & Chr(34) & "Error Number: " & Chr(34) & " & Err.Number & " & Chr(34) & " " & Chr(34) & " & " & Chr(34) & "Error Description: " & Chr(34) & " & Err.Description, vbCritical, " & Chr(34) & "Error" & Chr(34) & "" & vbCrLf
    sMsg = sMsg & "End Sub" & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    sMsg = sMsg & "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)" & vbCrLf
    sMsg = sMsg & "    If UnloadMode = vbFormControlMenu Then 'from close button (x)" & vbCrLf
    sMsg = sMsg & "        blnDontSetData = True" & vbCrLf
    sMsg = sMsg & "        Cancel = True" & vbCrLf
    sMsg = sMsg & "        Me.Hide" & vbCrLf
    sMsg = sMsg & "    End If" & vbCrLf
    sMsg = sMsg & "End Sub" & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    strCodeFormLookUp = sMsg
End Function

Public Function strCodeControlDefinitions() As String
    Dim sMsg As String
    sMsg = sMsg & "'Defini-definisi CommandButton pada Form" & vbCrLf
    sMsg = sMsg & "Public Const COMMAND_FORM_ADD As Integer = 0" & vbCrLf
    sMsg = sMsg & "Public Const COMMAND_FORM_EDIT As Integer = 1" & vbCrLf
    sMsg = sMsg & "Public Const COMMAND_FORM_SAVE As Integer = 2" & vbCrLf
    sMsg = sMsg & "Public Const COMMAND_FORM_CANCEL As Integer = 3" & vbCrLf
    sMsg = sMsg & "Public Const COMMAND_FORM_DELETE As Integer = 4" & vbCrLf
    sMsg = sMsg & "Public Const COMMAND_FORM_EXIT As Integer = 5" & vbCrLf
    sMsg = sMsg & "Public Const COMMAND_FORM_COLLAPSE As Integer = 6" & vbCrLf
    sMsg = sMsg & "Public Const COMMAND_FORM_REFRESH As Integer = 7" & vbCrLf
    strCodeControlDefinitions = sMsg
End Function

Public Function strModDatabase() As String
    Dim sMsg As String
    sMsg = sMsg & AddOptionExplicit
    sMsg = sMsg & "Public conn As New ADODB.Connection" & vbCrLf & vbCrLf
    sMsg = sMsg & "Public Function CloseRecordset(r As ADODB.Recordset, Optional SetNothing As Boolean = False) As Boolean" & vbCrLf
    sMsg = sMsg & "    On Error GoTo ErrHandler" & vbCrLf
    sMsg = sMsg & "        If Not r Is Nothing Then" & vbCrLf
    sMsg = sMsg & "            If r.State <> adStateClosed Then" & vbCrLf
    sMsg = sMsg & "                r.Close" & vbCrLf
    sMsg = sMsg & "            End If" & vbCrLf
    sMsg = sMsg & "        End If" & vbCrLf
    sMsg = sMsg & "        If SetNothing Then" & vbCrLf
    sMsg = sMsg & "            Set r = Nothing" & vbCrLf
    sMsg = sMsg & "        End If" & vbCrLf
    sMsg = sMsg & "        CloseRecordset = True" & vbCrLf
    sMsg = sMsg & "        Exit Function" & vbCrLf
    sMsg = sMsg & "ErrHandler:" & vbCrLf
    sMsg = sMsg & "        CloseRecordset = False" & vbCrLf
    sMsg = sMsg & "End Function" & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    sMsg = sMsg & "Public Function CloseDatabase(c As ADODB.Connection, Optional SetNothing As Boolean = False) As Boolean" & vbCrLf
    sMsg = sMsg & "    On Error GoTo ErrHandler" & vbCrLf
    sMsg = sMsg & "        If Not c Is Nothing Then" & vbCrLf
    sMsg = sMsg & "            If c.State <> adStateClosed Then" & vbCrLf
    sMsg = sMsg & "                c.Close" & vbCrLf
    sMsg = sMsg & "            End If" & vbCrLf
    sMsg = sMsg & "        End If" & vbCrLf
    sMsg = sMsg & "        If SetNothing Then" & vbCrLf
    sMsg = sMsg & "            Set c = Nothing" & vbCrLf
    sMsg = sMsg & "        End If" & vbCrLf
    sMsg = sMsg & "        CloseDatabase = True" & vbCrLf
    sMsg = sMsg & "        Exit Function" & vbCrLf
    sMsg = sMsg & "ErrHandler:" & vbCrLf
    sMsg = sMsg & "        CloseDatabase = False" & vbCrLf
    sMsg = sMsg & "End Function" & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    sMsg = sMsg & "Public Function nn(var As Variant) As String" & vbCrLf
    sMsg = sMsg & "    nn = IIf(IsNull(var), " & Chr(34) & "" & Chr(34) & ", var)" & vbCrLf
    sMsg = sMsg & "End Function" & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    sMsg = sMsg & "Public Function SetNullIfEmpty(ByVal sData As String) As Variant" & vbCrLf
    sMsg = sMsg & "    If Trim(sData) = " & Chr(34) & "" & Chr(34) & " Then" & vbCrLf
    sMsg = sMsg & "        SetNullIfEmpty = Null" & vbCrLf
    sMsg = sMsg & "    Else" & vbCrLf
    sMsg = sMsg & "        SetNullIfEmpty = sData" & vbCrLf
    sMsg = sMsg & "    End If" & vbCrLf
    sMsg = sMsg & "End Function" & vbCrLf
    strModDatabase = sMsg
End Function


