Attribute VB_Name = "modStringTable"
Option Explicit

Public Type tField
    FieldName As String
    MaxLength As Byte
Type As Variant
End Type

Public Type tTabel
    TabelName As String
    aField() As tField
End Type

Public aTabel() As tTabel

Public Function GetValidName(strName As String) As String
    '----------------------------------------------------------------------------
    'Fungsi Untuk merubah sebuah string dari:
    '"Jenis Bangunan" menjadi "JenisBangunan"
    '----------------------------------------------------------------------------

    Dim v As Variant
    Dim i As Integer
    Dim strNameToModify As String

    strNameToModify = strName
    strNameToModify = StrConv(strNameToModify, vbProperCase)

    v = Array(" ", "-", "/", ">", ".", "*", "&")
    For i = LBound(v) To UBound(v)
        strNameToModify = Replace(strNameToModify, v(i), "")
    Next
    GetValidName = strNameToModify

End Function

Public Function BlokUCase(StringCaption As String, Optional StartCaption As Integer) As String
    '----------------------------------------------------------------------------
    'Fungsi Untuk merubah sebuah string dari:
    '"tbl_form_bangunan" menjadi "Form Banguanan"
    'digunakan untuk penamaan Caption Form atau Label
    '----------------------------------------------------------------------------
    Dim i, s
    Dim sCaption As String
    sCaption = StringCaption

    '    MsgBox Left(sCaption, 4)
    If Left(sCaption, 4) = "tbl_" Then
        sCaption = Mid(sCaption, 5, Len(sCaption))
    ElseIf Left(sCaption, 2) = "t_" Then
        sCaption = Mid(sCaption, 3, Len(sCaption))
    ElseIf Left(sCaption, 3) = "tb_" Then
        sCaption = Mid(sCaption, 4, Len(sCaption))
    ElseIf Left(sCaption, 6) = "table_" Then
        sCaption = Mid(sCaption, 7, Len(sCaption))
    ElseIf Left(sCaption, 3) = "tbl" Then
        sCaption = Mid(sCaption, 4, Len(sCaption))
    ElseIf Left(sCaption, 2) = "tb" Then
        If Asc(Mid(sCaption, 3, 1)) > 64 And Asc(Mid(sCaption, 3, 1)) < 90 Then
            sCaption = Mid(sCaption, 3, Len(sCaption))
        End If
    ElseIf Left(sCaption, 1) = "t" Then
        If Asc(Mid(sCaption, 2, 1)) > 64 And Asc(Mid(sCaption, 2, 1)) < 90 Then
            sCaption = Mid(sCaption, 2, Len(sCaption))
        End If
    End If

    If InStr(sCaption, " ") > 0 Then
        sCaption = Replace(sCaption, "_", " ")
        sCaption = StrConv(sCaption, vbProperCase)
    End If

    For i = 1 + StartCaption To Len(sCaption)
        If i <> 1 Then
            Select Case Asc(Mid(sCaption, i, 1))
                Case 97 To 122 'Huruf Kecil
                    'Apabila huruf kecil maka lanjutkan ...
                    s = s & Mid(sCaption, i, 1)
                Case 65 To 90 'Hurup Besar
                    'Apabila ada huruf besar maka pisah dengan spasi
                    s = s & " " & Mid(sCaption, i, 1)
            End Select
        Else
            'Apabila huruf pertama maka tidak dipisah dengan spasi
            s = s & Mid(sCaption, i, 1)
        End If
    Next

    BlokUCase = StrConv(s, vbProperCase)

End Function

    


