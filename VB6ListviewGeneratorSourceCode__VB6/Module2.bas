Attribute VB_Name = "modDatabase"
Option Explicit

Public DBase As New ADODB.Connection
Public cat As ADOX.Catalog

Function OpenDataBase(sFilename As String) As Boolean
    ' Membuat koneksi ke database
    Set DBase = New ADODB.Connection
    With DBase
        .CursorLocation = adUseClient
        .Open sFilename
    End With
    OpenDataBase = True
End Function

Public Sub CloseAll()
    Dim frm As Form
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
End Sub

Public Sub CekDataBase()
    If Not DBase Is Nothing Then
        If DBase.State = adStateOpen Then
            DBase.Close
        End If
        Set DBase = Nothing
    End If
End Sub

    ' BeginCreateDatabseVB
Sub CreateDatabase(Filename As String)
    On Error GoTo CreateDatabaseError
    Dim con As ADODB.Connection
    Dim cat As New ADOX.Catalog
    Set con = cat.Create("Provider='Microsoft.Jet.OLEDB.4.0';Data Source='" & Filename & "'")
    CreateTables con
    'Clean up
    Set cat = Nothing
    Exit Sub

CreateDatabaseError:
    Set cat = Nothing

    If Err <> 0 Then
        MsgBox Err.Source & "-->" & Err.Description, , "Error"
    End If
End Sub

    ' EndCreateDatabaseVB

Public Sub CreateTables(ActiveConnection As ADODB.Connection)
    '    On Error GoTo CreateTableError

    Dim tbl(2) As New Table
    Dim cat As New ADOX.Catalog

    cat.ActiveConnection = ActiveConnection
    With tbl(0)
        .Name = "Menu"
        .Columns.Append "menu_id", adSmallInt
        .Columns.Append "menu_level", adSmallInt
        .Columns.Append "menu_name", adVarWChar, 25
        .Columns.Append "menu_caption", adVarWChar, 40
        cat.Tables.Append tbl(0)
    End With
    With tbl(1)
        .Name = "Operator"
        .Columns.Append "Jabatan_ID", adVarWChar, 40
        .Columns.Append "Hak_Akses", adVarWChar, 40
        cat.Tables.Append tbl(1)
    End With
    With tbl(2)
        .Name = "Kedudukan"
        .Columns.Append "Operator_ID", adSmallInt
        .Columns.Append "Jabatan"
        .Columns.Append "Nama"
        .Columns.Append "Password"
        .Columns.Append "Alamat"
        cat.Tables.Append tbl(2)
    End With
    MsgBox IsExistTable("Kedudukan", ActiveConnection)
    MsgBox IsExistTable("eeee", ActiveConnection)

    'Clean up
    Set cat.ActiveConnection = Nothing
    Set cat = Nothing
    Set tbl(0) = Nothing
    Set tbl(1) = Nothing
    Set tbl(2) = Nothing
    Exit Sub

CreateTableError:

    Set cat = Nothing
    Set tbl(0) = Nothing
    Set tbl(1) = Nothing
    Set tbl(2) = Nothing

    If Err <> 0 Then
        MsgBox Err.Source & "-->" & Err.Description, , "Error"
    End If
End Sub

Public Function IsExistTable(strName As String, ActiveConnection As ADODB.Connection) As Boolean
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = ActiveConnection
    If cat.Tables.Item(strName).Name <> "" Then
        IsExistTable = True
    End If
    Set cat = Nothing
End Function

    


