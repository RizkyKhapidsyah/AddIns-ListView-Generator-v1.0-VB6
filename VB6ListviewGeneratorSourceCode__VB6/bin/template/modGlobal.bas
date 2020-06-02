Attribute VB_Name = "modGlobal"
Option Explicit

Public Const DB_TYPE_MYSQL As String = "MYSQL"
Public Const DB_TYPE_ACCESS As String = "MS JET"
'Public Const DB_TYPE_ etc .... add here

Global gstrDbType As String

Public Function GetDBDate(dtDate As Date) As String
    Select Case UCase$(gstrDbType)
        Case DB_TYPE_MYSQL
            GetDBDate = "'" & Format$(dtDate, "YYYY-MM-DD") & "'"
        Case DB_TYPE_ACCESS
            GetDBDate = dtDate
        Case Else 'etc .... add here
            GetDBDate = dtDate
    End Select
End Function

Public Function GetDatabaseType() As String
    Dim i As Integer
    For i = 1 To conn.Properties.Count - 1
        If conn.Properties(i).Name = "DBMS Name" Then
            GetDatabaseType = conn.Properties(i).Value
            Exit For
        End If
    Next
End Function

'sq = handling single quote
Public Function sq(Text As String) As String
    sq = Replace(Text, "'", "''")
End Function
