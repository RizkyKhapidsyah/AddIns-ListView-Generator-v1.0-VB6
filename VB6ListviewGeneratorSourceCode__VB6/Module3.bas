Attribute VB_Name = "modConnectionString"
Option Explicit

Public Function getADOConnectionString(Optional ByVal cnStringToEdit As String = "", Optional sPrePromptUserMessage As String = "") As String

    Dim sActivity As String
    Dim dl As Object    ' DataLinks
    Dim cn As Object ' ADODB.Connection

    On Error GoTo ErrGetAdoConnectionString
    sActivity = "Creating Datalinks object."

    ' we're using CreateObject so that we don't add a project reference dependdency.
    Set dl = CreateObject("DataLinks")

    If Not ("" = cnStringToEdit) Then
        ' Edit / Append Specified Connect String

        ' Check for user prompt message and display first if specified
        If Not ("" = sPrePromptUserMessage) Then
            MsgBox sPrePromptUserMessage, vbInformation, "Connecting to Database..."
        End If
        sActivity = "Creating ADODB.Connection object"
        Set cn = CreateObject("ADODB.Connection")

        cn.ConnectionString = "Provider=SQLOLEDB.1;Initial Catalog=PUBS"

        ' Call prompt new to have user build connect string; returns a connection object
        sActivity = "Prompting user to  edit  connect string"
        dl.PromptEdit cn

    Else
        sActivity = "Prompting user for  new  connect string"
        Set cn = dl.PromptNew()

    End If

    ' They clicked cancel
    If cn Is Nothing Then
        getADOConnectionString = ""
        Exit Function
    End If

    ' retrieve connection string
    getADOConnectionString = cn.ConnectionString

    ' Immediately release the connection object
    Set cn = Nothing

    Exit Function

ErrGetAdoConnectionString:

    Dim sMsg As String

    ' Immediately release the connection object
    Set cn = Nothing

    sMsg = "Error While [" + sActivity + "].  Details are below: " + vbCrLf
    sMsg = sMsg + "Description:[" + Err.Description + "]." + vbCrLf
    sMsg = sMsg + "Source:[" & Err.Source & "]." + vbCrLf
    sMsg = sMsg + "Number:[" & Err.Number & "]." + vbCrLf
    sMsg = sMsg + "Help File:[" & Err.HelpFile & "]." + vbCrLf
    MsgBox sMsg, vbCritical, "Error Connecting to Database."

End Function

    


