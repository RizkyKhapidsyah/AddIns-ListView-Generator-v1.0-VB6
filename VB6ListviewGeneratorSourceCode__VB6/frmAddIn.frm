VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listview - Generator 3.0"
   ClientHeight    =   8130
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTables 
      Height          =   315
      Left            =   960
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   420
      Width           =   4635
   End
   Begin VB.Frame Frame3 
      Height          =   4515
      Left            =   120
      TabIndex        =   14
      Top             =   180
      Width           =   6195
      Begin VB.ComboBox cmbUniqID 
         Height          =   315
         Left            =   4020
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton OKButton 
         Caption         =   "&Generate Code"
         Height          =   375
         Left            =   3900
         TabIndex        =   25
         Top             =   4020
         Width           =   2175
      End
      Begin VB.CommandButton cmdDown 
         Height          =   375
         Left            =   5640
         Picture         =   "frmAddIn.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1710
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdUp 
         Height          =   375
         Left            =   5625
         Picture         =   "frmAddIn.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1185
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdLeftAll 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00000000&
         TabIndex        =   22
         Top             =   2220
         Width           =   495
      End
      Begin VB.CommandButton cmdLeftOne 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00000000&
         TabIndex        =   21
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton cmdRightAll 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00000000&
         TabIndex        =   20
         Top             =   1380
         Width           =   495
      End
      Begin VB.CommandButton cmdRightOne 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00000000&
         TabIndex        =   19
         Top             =   960
         Width           =   495
      End
      Begin VB.ListBox lstSelected 
         Height          =   1200
         IntegralHeight  =   0   'False
         Left            =   3240
         TabIndex        =   18
         Top             =   915
         Width           =   2220
      End
      Begin VB.ListBox lstAll 
         Height          =   1620
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   17
         Top             =   915
         Width           =   2220
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "C&onnection String"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         TabIndex        =   16
         Top             =   4020
         Width           =   2175
      End
      Begin VB.TextBox txtCon 
         Height          =   915
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2880
         Width           =   5955
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Primary"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3240
         TabIndex        =   33
         Tag             =   "2407"
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tab&les:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Tag             =   "2406"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label lblSelected 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Selected Items:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3180
         TabIndex        =   28
         Tag             =   "2407"
         Top             =   660
         Width           =   1125
      End
      Begin VB.Label lblAll 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&All Items:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Tag             =   "2406"
         Top             =   660
         Width           =   675
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Connection String:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Tag             =   "2406"
         Top             =   2640
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1035
      Index           =   0
      Left            =   2820
      ScaleHeight     =   975
      ScaleWidth      =   3075
      TabIndex        =   8
      Top             =   6600
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   960
         MaxLength       =   1
         TabIndex        =   10
         Text            =   "2"
         Top             =   480
         Width           =   555
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Alternate Color"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "per Rows"
         Height          =   195
         Left            =   1620
         TabIndex        =   11
         Top             =   540
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1035
      Index           =   2
      Left            =   3240
      ScaleHeight     =   975
      ScaleWidth      =   3075
      TabIndex        =   13
      Top             =   6960
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   1035
      Index           =   1
      Left            =   3060
      ScaleHeight     =   975
      ScaleWidth      =   3075
      TabIndex        =   12
      Top             =   6780
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Environment"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   6195
      Begin VB.CheckBox chkHideBalloonTips 
         Caption         =   "&Hide BalloonTip Text"
         Height          =   195
         Left            =   3600
         TabIndex        =   3
         Top             =   1080
         Width           =   2235
      End
      Begin VB.CheckBox chkMDIChild 
         Caption         =   "&MDI Child Form"
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkIncludeEditForm 
         Caption         =   "&Include Edit Form"
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "http://khoiriyyah.blogspot.com"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   660
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listview Type"
      Height          =   1395
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   6195
      Begin VB.OptionButton optListviewType 
         Caption         =   "&ComCtlLib"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   7
         Top             =   960
         Width           =   1995
      End
      Begin VB.OptionButton optListviewType 
         Caption         =   "&MSComCtlLib"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Top             =   660
         Width           =   1995
      End
      Begin VB.OptionButton optListviewType 
         Caption         =   "&Codejock"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit

Dim strCon As String
Dim BalloonTips As New clsToolTips
Dim strBetweenSQL As String

Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub cboTables_Click()
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = DBase
    Dim c As ADOX.Column
    lstAll.Clear
    lstSelected.Clear
    For Each c In cat.Tables(GetTabelValidName(cboTables.Text)).Columns
        lstAll.AddItem c.Name
    Next
    If lstAll.ListCount > 0 Then lstAll.ListIndex = 0
    Set cat = Nothing
End Sub

Private Sub chkHideBalloonTips_Click()
    If chkHideBalloonTips.Value = vbChecked Then
        HideBalloonTips
    Else
        ShowBalloonTips
    End If
End Sub

Private Sub cmbUniqID_Click()
    Dim TableName As String
    Dim strPrimaryKey As String
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = DBase
    TableName = GetValidName(GetTabelValidName(cboTables.Text))
    strPrimaryKey = SignBetweenAndValue(cat.Tables(TableName).Columns(cmbUniqID.Text).Type)
    Set cat = Nothing
End Sub

Private Function IsAutoIncrement(TableName As String, FieldName As String) As Boolean
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM " & TableName, DBase, adOpenKeyset, adLockOptimistic
    IsAutoIncrement = rs.Fields(FieldName).Properties("IsAutoIncrement").Value
End Function

Private Sub cmdConnect_Click()

    On Error GoTo ErrHandler

    strCon = getADOConnectionString()
    If strCon = "" Then Exit Sub
    txtCon.Text = Replace(strCon, Chr(34), "")

    If OpenDataBase(strCon) = True Then
        FillComboWithTables
        ChangeComboDropDownHeight Me, cboTables, 35
    End If

    Exit Sub

ErrHandler:

    MsgBox Err.Number & vbNewLine & Err.Description, vbExclamation + vbOKOnly, "Connection Error"

    lstAll.Clear
    cboTables.Clear
    lstSelected.Clear
End Sub

Private Sub Form_Initialize()
    ShowBalloonTips
End Sub

Private Sub ShowBalloonTips()
    With BalloonTips
        .Add OKButton, "Melakukan generate code dan interface setelah melakukan koneksi database" & vbCrLf & _
        " dan berhasil, setelah memilih tabel dan memilih field", "Generate Code & Interface", styleBalloon, 3
        .Add Text1, "Jumlah perbedaan warna pada row Listview, apabila CheckBox Alternate Color dicentang", "Perbedaan Warna", styleBalloon, 3
        .Add optListviewType(0), "Jenis Listview yang dipilih.", "Codejock Listview", styleBalloon, 3
        .Add optListviewType(1), "Jenis Listview yang dipilih.", "MSComCtlLib Listview", styleBalloon, 3
        .Add optListviewType(2), "Jenis Listview yang dipilih.", "ComCtlLib Listview", styleBalloon, 3
    End With
End Sub

Private Sub HideBalloonTips()
    BalloonTips.Clear
End Sub

Private Sub Form_Load()
    optListviewType_Click LISTVIEW_CODEJOCK
    Dim i As Integer
    For i = 0 To 2
        Picture1(i).BorderStyle = vbBSNone
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set BalloonTips = Nothing
End Sub

Private Sub OKButton_Click()

    If DBase.State <> adStateOpen Then
        If MsgBox("Lakukan koneksi database terlebih dahulu!, Apakah Anda ingin melakukan koneksi database?", vbInformation + vbYesNo, "Peringatan") = vbYes Then
            cmdConnect_Click
        End If
        Exit Sub
    End If

    If cboTables.ListIndex < 0 Then
        If MsgBox("Tabel belum dipilih, pilih tabel dari ComboBox di atas!", vbInformation + vbYesNo, "Peringatan") = vbYes Then
            ShowDropDown cboTables
        End If
        Exit Sub
    End If

    If lstSelected.ListCount = 0 Then
        MsgBox "Belum ada item/fields terpilih, minimal harus ada 1 (satu) field terpilih", vbInformation, "Peringatan"
        Exit Sub
    End If

    If cmbUniqID.Text = "" Then
        If MsgBox("Pilih primary key terlebih dahulu!", vbInformation + vbYesNo, "Peringatan") = vbYes Then
            ShowDropDown cmbUniqID
        End If
        Exit Sub
    End If

    Dim strPrimaryKey As String
    Dim f As VBForm
    Dim c As VBControl
    Dim v As VBComponent
    Dim sData As String

    If IsOCXRegistered(ListviewType.GUID) = False Then
        MsgBox "Anda tidak memiliki komponen " & ListviewType.ProgID & " atau komponen tersebut belum teregister", vbCritical, "Error"
        Exit Sub
    Else
        InsertOCX ListviewType.GUID
    End If

    InsertReferences "{2A75196C-D9EB-4129-B803-931327F72D5C}", 2, 8 'Microsoft ActiveX Data Objects 2.8 Library

    Dim frm As VBIDE.VBForm
    Dim ctl As VBControl
    Dim k, i, sTextName As String
    Dim TableName As String
    Dim sBody As String
    Dim sBodyData As String
    
    TableName = GetValidName(GetTabelValidName(cboTables.Text))
    Set modRegistry.VBInstance = Me.VBInstance

    If IsExistVBComponent("frmLv" & TableName) Then
        VBInstance.ActiveVBProject.VBComponents("frmLv" & TableName).Activate
        MsgBox "Dalam project telah ada nama frmLv" & TableName, vbInformation, "Dobel VBComponent"
        Set v = Nothing
        Set modRegistry.VBInstance = Nothing
        Exit Sub
    End If

    Set v = VBInstance.ActiveVBProject.VBComponents.Add(vbext_ct_VBForm)

    With v
        .Properties("Name") = "frmLv" & TableName
        .Properties("Caption") = BlokUCase(GetTabelValidName(cboTables.Text))
        .Properties("Height") = 5800
        .Properties("WindowState") = vbMaximized
        If chkMDIChild.Value = vbChecked Then
            .Properties("MDIChild") = True
        End If
    End With

    Set frm = VBInstance.SelectedVBComponent.Designer

    Do While ctl Is Nothing
        Set ctl = frm.VBControls.Add(ListviewType.ProgID)
    Loop

    With ctl.ControlObject
        .Name = "lv" & TableName
        .Left = 50
        .Top = 50
        If optListviewType(0).Value = True Then .IconSize = 20
        .HideSelection = False
        .Height = v.Properties("ScaleHeight") - 1240
        .Width = v.Properties("ScaleWidth") - 240
        .View = 3 'xtpListViewReport
        .Appearance = 0 'xtpAppearanceStandard
        If optListviewType(0).Value = True Or optListviewType(1).Value = True Then
            .FullRowSelect = True
            .GridLines = True
        End If
    End With

    Dim sColomHeader As String

    For i = 0 To lstSelected.ListCount
        With ctl.ControlObject
            If i > 0 Then
                sBody = sBody & "                .SubItems(" & i & ") = nn(rs" & TableName & "!" & AutoBracket(lstSelected.List(i - 1)) & ")" & vbCrLf
                sBodyData = sBodyData & "        .SubItems(" & i & ") = nn(rs!" & AutoBracket(lstSelected.List(i - 1)) & ")" & vbCrLf
                sData = sData & "                    !" & lstSelected.List(i - 1) & " = txt" & lstSelected.List(i - 1) & ".Text" & vbCrLf
            End If
            If i = 0 Then
                sColomHeader = sColomHeader & "        .Add , , " & Chr(34) & "No." & Chr(34) & ", 900" & vbCrLf
            Else
                sColomHeader = sColomHeader & "        .Add , , " & Chr(34) & lstSelected.List(i - 1) & Chr(34) & ", 1500" & vbCrLf
            End If
        End With
    Next

    Dim sMsg As String
    sMsg = AddOptionExplicit
    sMsg = sMsg & "Dim rs" & TableName & " As New ADODB.Recordset" & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    
    sMsg = sMsg & "Private Sub Form_Load()" & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    sMsg = sMsg & "    Dim list_item As " & ListviewType.ListItemName & vbCrLf
    sMsg = sMsg & "    conn.Open " & Chr(34) & txtCon.Text & Chr(34) & vbCrLf
    sMsg = sMsg & "    gstrDbType = GetDatabaseType" & vbCrLf
    sMsg = sMsg & "    rs" & TableName & ".Open " & Chr(34) & "SELECT * FROM " & AutoBracket(GetTabelValidName(cboTables.Text)) & Chr(34) & ", conn, adOpenKeyset, adLockOptimistic" & vbCrLf
    sMsg = sMsg & "" & vbCrLf

    sMsg = sMsg & "    With lv" & TableName & ".ColumnHeaders" & vbCrLf
    sMsg = sMsg & sColomHeader
    sMsg = sMsg & "    End With" & vbCrLf

    sMsg = sMsg & "" & vbCrLf
    sMsg = sMsg & "    Do While Not rs" & TableName & ".EOF" & vbCrLf
    sMsg = sMsg & "        Set list_item = lv" & TableName & ".ListItems.Add(, " & Chr(34) & "a" & Chr(34) & " & rs" & TableName & "!" & cmbUniqID.Text & " , rs" & TableName & ".AbsolutePosition)" & vbCrLf
    sMsg = sMsg & "            With list_item" & vbCrLf
    sMsg = sMsg & sBody
    sMsg = sMsg & "            End With" & vbCrLf
    sMsg = sMsg & "        rs" & TableName & ".MoveNext" & vbCrLf
    sMsg = sMsg & "    Loop" & vbCrLf
    sMsg = sMsg & "    rs" & TableName & ".Close" & vbCrLf
    sMsg = sMsg & "" & vbCrLf

    sMsg = sMsg & "End Sub" & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    
    sMsg = sMsg & "Private Sub Form_Resize()" & vbCrLf
    sMsg = sMsg & "    On Error Resume Next" & vbCrLf
    sMsg = sMsg & "    With lv" & TableName & vbCrLf
    sMsg = sMsg & "        .Left = 0" & vbCrLf
    sMsg = sMsg & "        .Top = 0" & vbCrLf
    sMsg = sMsg & "        .Width = Me.ScaleWidth" & vbCrLf
    sMsg = sMsg & "        .Height = Me.ScaleHeight - (.Top + Picture1.Height)" & vbCrLf
    sMsg = sMsg & "    End With" & vbCrLf
    sMsg = sMsg & "End Sub" & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    
    sMsg = sMsg & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    sMsg = sMsg & "    CloseRecordset rs" & TableName & ", True" & vbCrLf
    sMsg = sMsg & "End Sub" & vbCrLf
    sMsg = sMsg & "" & vbCrLf

    sMsg = sMsg & "Private Function nn(var As Variant) As String" & vbCrLf
    sMsg = sMsg & "    nn = IIf(IsNull(var), " & Chr(34) & "" & Chr(34) & ", var)" & vbCrLf
    sMsg = sMsg & "End Function" & vbCrLf

    If chkIncludeEditForm.Value = 1 Then
    
        sMsg = sMsg & "Private Sub cmdCommand_Click(Index As Integer)" & vbCrLf
        sMsg = sMsg & "        DoCommand Index" & vbCrLf
        sMsg = sMsg & "End Sub" & vbCrLf
        sMsg = sMsg & "" & vbCrLf

        sMsg = sMsg & "" & vbCrLf

        sMsg = sMsg & "Private Sub lv" & TableName & "_DblClick()" & vbCrLf
        sMsg = sMsg & "    DoCommand COMMAND_FORM_EDIT" & vbCrLf
        sMsg = sMsg & "End Sub" & vbCrLf
        sMsg = sMsg & "" & vbCrLf

        sMsg = sMsg & "Public Sub DoCommand(Index As Integer)" & vbCrLf
        sMsg = sMsg & "" & vbCrLf
        sMsg = sMsg & "On Error Goto ErrHandler" & vbCrLf
        sMsg = sMsg & "" & vbCrLf
        sMsg = sMsg & "    Dim lv As " & ListviewType.ProgID & vbCrLf
        sMsg = sMsg & "    Dim List_Item As " & ListviewType.ListItemName & vbCrLf
        sMsg = sMsg & "    Dim i As Integer" & vbCrLf
        sMsg = sMsg & "    Dim frm As Form" & vbCrLf
        sMsg = sMsg & "    Dim key As Variant" & vbCrLf
        sMsg = sMsg & "    " & vbCrLf
        sMsg = sMsg & "    Set lv = lv" & TableName & vbCrLf
        sMsg = sMsg & "    Set frm = New " & "frmEdit" & TableName & vbCrLf
        sMsg = sMsg & "    " & vbCrLf
        sMsg = sMsg & "    If lv.ListItems.Count <= 0 Then Exit Sub" & vbCrLf
        sMsg = sMsg & "" & vbCrLf
        sMsg = sMsg & "    ReDim a(lv.ColumnHeaders.Count - 1)" & vbCrLf
        sMsg = sMsg & "" & vbCrLf
        sMsg = sMsg & "    If Not lv.SelectedItem Is Nothing Then" & vbCrLf
        sMsg = sMsg & "        key = right(lv.SelectedItem.key,len(lv.SelectedItem.key)-1)" & vbCrLf
        sMsg = sMsg & "    End If" & vbCrLf
        sMsg = sMsg & "    " & vbCrLf
        
        sMsg = sMsg & "    Select Case Index" & vbCrLf
        sMsg = sMsg & "        Case COMMAND_FORM_ADD, COMMAND_FORM_EDIT" & vbCrLf
        sMsg = sMsg & "                If Index = COMMAND_FORM_EDIT Then" & vbCrLf
        sMsg = sMsg & "                    If lv.SelectedItem Is Nothing Then" & vbCrLf
        sMsg = sMsg & "                        Set frm = Nothing" & vbCrLf
        sMsg = sMsg & "                        Exit Sub" & vbCrLf
        sMsg = sMsg & "                    End If" & vbCrLf
        sMsg = sMsg & "                    frm.ID = key" & vbCrLf
        sMsg = sMsg & "                    frm.MODE = 1" & vbCrLf
        sMsg = sMsg & "                Else" & vbCrLf
        sMsg = sMsg & "                    frm.MODE = 0" & vbCrLf
        sMsg = sMsg & "                End If" & vbCrLf
        sMsg = sMsg & "                frm.Show vbModal" & vbCrLf
        sMsg = sMsg & "                If Not frm.blnDontSetData Then" & vbCrLf
        sMsg = sMsg & "                    GetDataForm frm.txt" & cmbUniqID.Text & ".Text, frm.MODE" & vbCrLf
        sMsg = sMsg & "                End If" & vbCrLf
        sMsg = sMsg & "        Case COMMAND_FORM_DELETE" & vbCrLf
        sMsg = sMsg & "            If Not lv.SelectedItem Is Nothing Then" & vbCrLf
        sMsg = sMsg & "                If MsgBox(" & Chr(34) & "Are you sure?" & Chr(34) & ", vbQuestion + vbYesNo, " & Chr(34) & "Delete Confirm" & Chr(34) & ") = vbNo Then Exit Sub" & vbCrLf
        sMsg = sMsg & "                conn.Execute " & Chr(34) & "DELETE FROM " & TableName & " WHERE " & cmbUniqID.Text & " =" & Chr(34) & " & key" & vbCrLf
        sMsg = sMsg & "                lv.ListItems.Remove lv.SelectedItem.Index" & vbCrLf
        sMsg = sMsg & "                lv.ListItems(lv.SelectedItem.Index).Selected = True" & vbCrLf
        sMsg = sMsg & "            End If" & vbCrLf
        sMsg = sMsg & "        Case COMMAND_FORM_EXIT" & vbCrLf
        sMsg = sMsg & "            Unload Me" & vbCrLf
        sMsg = sMsg & "    End Select" & vbCrLf
    
        sMsg = sMsg & "" & vbCrLf
        sMsg = sMsg & "    Unload frm" & vbCrLf
        sMsg = sMsg & "    Set frm = Nothing" & vbCrLf
        sMsg = sMsg & "    Set lv = Nothing" & vbCrLf
        sMsg = sMsg & "" & vbCrLf
        sMsg = sMsg & "    Exit Sub" & vbCrLf
        sMsg = sMsg & "" & vbCrLf
        sMsg = sMsg & "ErrHandler:" & vbCrLf
        sMsg = sMsg & "" & vbCrLf
        sMsg = sMsg & "    MsgBox " & Chr(34) & "Error Number: " & Chr(34) & " & Err.Number & " & Chr(34) & " " & Chr(34) & " & " & Chr(34) & "Error Description: " & Chr(34) & " & Err.Description, vbCritical, " & Chr(34) & "Error" & Chr(34) & "" & vbCrLf
        sMsg = sMsg & "    Unload frm" & vbCrLf
        sMsg = sMsg & "    Set frm = Nothing" & vbCrLf
        sMsg = sMsg & "    Set lv = Nothing" & vbCrLf
        sMsg = sMsg & "" & vbCrLf
        sMsg = sMsg & "End Sub" & vbCrLf

        sMsg = sMsg & vbCrLf & "Private Function GetDataForm(ID As Variant, Mode As Integer)" & vbCrLf
        sMsg = sMsg & "    Dim rs As ADODB.Recordset" & vbCrLf
        sMsg = sMsg & "    Set rs = conn.Execute(" & Chr(34) & "SELECT * FROM " & TableName & " " & GetWhereString & ")" & vbCrLf
        sMsg = sMsg & "    Dim List_Item As " & ListviewType.ListItemName & vbCrLf
        sMsg = sMsg & "    If Mode = 1 Then" & vbCrLf
        sMsg = sMsg & "        Set List_Item = lv" & TableName & ".SelectedItem" & vbCrLf
        sMsg = sMsg & "    ElseIf Mode = 0 Then" & vbCrLf
        sMsg = sMsg & "        Set List_Item = lv" & TableName & ".ListItems.Add(, " & Chr(34) & "a" & Chr(34) & " & ID)" & vbCrLf
        sMsg = sMsg & "    End If" & vbCrLf
        sMsg = sMsg & "    With List_Item" & vbCrLf
        sMsg = sMsg & "        .key = " & Chr(34) & "a" & Chr(34) & " & rs!" & cmbUniqID.Text & vbCrLf
        sMsg = sMsg & sBodyData
        sMsg = sMsg & "    End With" & vbCrLf
        sMsg = sMsg & "End Function" & vbCrLf

    End If

    v.CodeModule.AddFromString sMsg
    
    If Not IsExistVBComponent("modGlobal") Then
        AddVBComponentFromFile App.Path & "\template\modGlobal.bas"
    End If
    
    'Insert macro script ----
    If chkIncludeEditForm.Value = 1 Then
        OpenTextFileAndTranslate App.Path & "\template\controls.robotscript"
    Else
        OpenTextFileAndTranslate App.Path & "\template\readonly.robotscript"
    End If
    '------------------------
    
    Dim X As Integer
    Dim itop As Integer
    Dim lbl As VBControl

    If chkIncludeEditForm.Value = 1 Then
        Set v = VBInstance.ActiveVBProject.VBComponents.Add(vbext_ct_VBForm)
        With v
            .Properties("Name") = "frmEdit" & TableName
            .Properties("Caption") = "Form Edit - " & BlokUCase(GetTabelValidName(cboTables.Text))
            .Properties("Height") = 4800
            .Properties("Width") = 6000
            .Properties("StartUpPosition") = vbCenter
        End With

        Dim h As Integer
        Set f = v.Designer
        With f
            For i = 0 To lstSelected.ListCount - 1
                Set c = f.VBControls.Add("VB.TextBox")
                With c.ControlObject
                    .Width = 3000
                    .Height = 315
                    .Left = 2500
                    If i = 0 Then
                        .Top = 300
                        h = 300 + 315 + 10
                    Else
                        .Top = h + X
                        X = X + 325
                    End If
                    itop = c.Properties("Top")
                    .Name = "txt" & lstSelected.List(i)
                    .Text = "txt" & lstSelected.List(i)
                    .MaxLength = GetDefinedSize(lstSelected.List(i))
                    '.Index = i
                End With

                Set lbl = f.VBControls.Add("VB.Label")
                With lbl.ControlObject
                    .Width = 2000
                    .Height = 315
                    .Left = 280
                    .Top = itop
                    .Name = "lblData"
                    .Caption = lstSelected.List(i)
                    .Index = i
                End With
            Next

            Set lbl = f.VBControls.Add("VB.CommandButton")
            With lbl.ControlObject
                .Width = 1335
                .Height = 435
                .Left = 2760
                .Top = X + 800
                .Default = True
                .Name = "cmdActions"
                .Caption = "&Save"
                .Index = COMMAND_FORM_SAVE
            End With

            Set lbl = f.VBControls.Add("VB.CommandButton")
            With lbl.ControlObject
                .Width = 1335
                .Height = 435
                .Left = 4200
                .Top = X + 800
                .Cancel = True
                .Name = "cmdActions"
                .Caption = "&Cancel"
                .Index = COMMAND_FORM_CANCEL
            End With

            v.Properties("Height") = X + 2000
            v.Properties("BorderStyle") = vbFixedDialog
            v.CodeModule.AddFromString strCodeFormLookUp
            v.CodeModule.AddFromString GetUpdate & SQLUpdate & GetExecute
            v.CodeModule.AddFromString GetInsert & SQLInsert & GetExecute(True)
            v.CodeModule.AddFromString GetFormLoad
        End With
        

        If Not IsExistVBComponent("modControlDefinitions") Then
            AddModule "modControlDefinitions", strCodeControlDefinitions
        End If
    
    End If
    
    If Not IsExistVBComponent("modDatabase") Then
        AddModule "modDatabase", strModDatabase
    End If
    
    Set modRegistry.VBInstance = Nothing

End Sub

Private Function GetFormLoad() As String

    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = DBase
    Dim c As ADOX.Column
    Dim newTable As ADOX.Table
    
    Dim i As Integer
    Dim s As String
    Dim X As Integer
    Dim r As String
    Dim tType As Integer
    Dim TableName As String

    TableName = GetValidName(GetTabelValidName(cboTables.Text))

    X = lstSelected.ListCount - 1
    Set newTable = cat.Tables("tbModem")
    For i = 0 To lstSelected.ListCount - 1
        r = r & "        txt" & lstSelected.List(i) & ".Text = " & Chr(34) & "" & Chr(34) & "" & vbCrLf
        s = s & "        txt" & lstSelected.List(i) & ".Text = nn(rs!" & lstSelected.List(i) & ")" & vbCrLf
    Next

    Dim sMsg As String
    sMsg = sMsg & vbCrLf & "Private Sub Form_Load()" & vbCrLf
    sMsg = sMsg & "    If MODE = 0 Then" & vbCrLf
    sMsg = sMsg & r
    sMsg = sMsg & "    Else" & vbCrLf
    sMsg = sMsg & "        Dim rs As ADODB.Recordset" & vbCrLf
    sMsg = sMsg & "        Set rs = conn.Execute(" & Chr(34) & "SELECT * FROM " & TableName & " " & GetWhereString & ")" & vbCrLf
    sMsg = sMsg & s
    sMsg = sMsg & "    End If" & vbCrLf
    sMsg = sMsg & "End Sub" & vbCrLf
    
    GetFormLoad = sMsg
    
End Function

Private Function GetExecute(Optional WithID As Boolean) As String
    Dim sMsg As String
    sMsg = sMsg & "    conn.Execute strSQL" & vbCrLf
    If WithID Then
        sMsg = sMsg & "    ID = conn.Execute(" & Chr(34) & "SELECT ID From Tbinoutsms ORDER BY ID DESC LIMIT 1" & Chr(34) & ")!ID" & vbCrLf
    End If
    sMsg = sMsg & "End Sub" & vbCrLf
    GetExecute = sMsg
End Function

Private Function GetUpdate() As String
    Dim sMsg As String
    sMsg = sMsg & "Private Sub UpdateData()" & vbCrLf
    sMsg = sMsg & "    Dim strSQL As String" & vbCrLf
    GetUpdate = sMsg
End Function

Private Function GetInsert() As String
    Dim sMsg As String
    sMsg = sMsg & "Private Sub InsertData()" & vbCrLf
    sMsg = sMsg & "    Dim strSQL As String" & vbCrLf
    GetInsert = sMsg
End Function

Private Function InsertOCX(ProgID As String) As Boolean
    On Error GoTo ErrHandler
    'Add OCX
    VBInstance.ActiveVBProject.AddToolboxProgID ProgID
    InsertOCX = True
    Exit Function
ErrHandler:
    InsertOCX = False
End Function

Private Function FillComboWithTables()
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = DBase
    Dim i As Integer
    cboTables.Clear
    lstAll.Clear
    lstSelected.Clear
    For i = 0 To cat.Tables.Count - 1
        If cat.Tables(i).Type <> "SYSTEM TABLE" And cat.Tables(i).Type <> "ACCESS TABLE" Then
            If cat.Tables(i).Type = "TABLE" Then
                cboTables.AddItem "Table : " & cat.Tables(i).Name
            Else
                cboTables.AddItem "query : " & cat.Tables(i).Name
            End If
        End If
    Next i
    Set cat = Nothing
End Function

Private Function GetTabelValidName(strName As String) As String
    Dim s() As String
    s = Split(strName, " : ")
    GetTabelValidName = s(1)
End Function

Private Sub cmdUp_Click()
    On Error Resume Next
    Dim nItem As Integer

    With lstSelected
        If .ListIndex < 0 Then Exit Sub
        nItem = .ListIndex
        If nItem = 0 Then Exit Sub  'can't move 1st item up
        .AddItem .Text, nItem - 1
        .RemoveItem nItem + 1
        .Selected(nItem - 1) = True
    End With
End Sub

Private Sub cmdDown_Click()
    On Error Resume Next
    Dim nItem As Integer

    With lstSelected
        If .ListIndex < 0 Then Exit Sub
        nItem = .ListIndex
        If nItem = .ListCount - 1 Then Exit Sub 'can't move last item down
        'move item down
        .AddItem .Text, nItem + 2
        'remove old item
        .RemoveItem nItem
        'select the item that was just moved
        .Selected(nItem + 1) = True
    End With
End Sub

Private Sub cmdRightOne_Click()
    On Error Resume Next
    Dim i As Integer

    If lstAll.ListCount = 0 Then Exit Sub

    lstSelected.AddItem lstAll.Text
    i = lstAll.ListIndex

    lstAll.RemoveItem lstAll.ListIndex
    If lstAll.ListCount > 0 Then
        If i > lstAll.ListCount - 1 Then
            lstAll.ListIndex = i - 1
        Else
            lstAll.ListIndex = i
        End If
    End If
    lstSelected.ListIndex = lstSelected.NewIndex
    FillComboUniqID
End Sub

Private Sub cmdRightAll_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To lstAll.ListCount - 1
        lstSelected.AddItem lstAll.List(i)
    Next
    lstAll.Clear
    lstSelected.ListIndex = 0
    FillComboUniqID
End Sub

Private Sub cmdLeftOne_Click()
    On Error Resume Next
    Dim i As Integer

    If lstSelected.ListCount = 0 Then Exit Sub

    lstAll.AddItem lstSelected.Text
    i = lstSelected.ListIndex
    lstSelected.RemoveItem i

    lstAll.ListIndex = lstAll.NewIndex
    If lstSelected.ListCount > 0 Then
        If i > lstSelected.ListCount - 1 Then
            lstSelected.ListIndex = i - 1
        Else
            lstSelected.ListIndex = i
        End If
    End If
    FillComboUniqID
End Sub

Private Sub cmdLeftAll_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To lstSelected.ListCount - 1
        lstAll.AddItem lstSelected.List(i)
    Next
    lstSelected.Clear
    lstAll.ListIndex = lstAll.NewIndex
    FillComboUniqID
End Sub

Private Sub FillComboUniqID()
    Dim i As Integer
    cmbUniqID.Clear
    For i = 0 To lstSelected.ListCount - 1
        cmbUniqID.AddItem lstSelected.List(i)
    Next
End Sub

Private Sub lstAll_DblClick()
    cmdRightOne_Click
End Sub

Private Sub lstSelected_DblClick()
    cmdLeftOne_Click
End Sub

Private Function AutoBracket(ByVal str As String) As String
    Dim s As String
    s = str
    If InStr(1, s, " ") Then
        s = "[" & s & "]"
    Else
        s = s
    End If
    AutoBracket = s
End Function

Private Function InsertReferences(GUID As String, Mayor As Long, Minor As Long) As Boolean
    On Error GoTo ErrHandler
    VBInstance.ActiveVBProject.References.AddFromGuid GUID, Mayor, Minor
    InsertReferences = True
ErrHandler:
    InsertReferences = False
End Function

Private Function IsExistVBComponent(strName As String) As Boolean
    Dim vc As VBComponent
    For Each vc In VBInstance.ActiveVBProject.VBComponents
        If LCase(vc.Name) = LCase(strName) Then
            IsExistVBComponent = True
            Exit For
        End If
    Next
End Function

Private Sub optListviewType_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Picture1(i).Visible = False
    Next
    With Picture1(Index)
        .Visible = True
        .Top = 6600
        .Left = 2820
    End With
    Select Case Index
        Case LISTVIEW_CODEJOCK
            With ListviewType
                .GUID = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}"
                .ProgID = "XtremeSuiteControls.Listview"
                .ListItemName = "XtremeSuiteControls.ListViewItem"
            End With
        Case LISTVIEW_MSCOMCTLLIB
            With ListviewType
                .GUID = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}"
                .ProgID = "MSComCtlLib.Listview"
                .ListItemName = "MSComCtlLib.ListItem"
            End With
        Case LISTVIEW_COMCTLLIB
            With ListviewType
                .GUID = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}"
                .ProgID = "ComCtlLib.Listview"
                .ListItemName = "ComCtlLib.ListItem"
            End With
    End Select
End Sub

Private Function FieldType(intType As Integer) As String

    Select Case intType
        Case adBigInt: FieldType = "adBigInt"
        Case adBinary: FieldType = "adBinary"
        Case adBoolean: FieldType = "adBoolean"
        Case adBSTR: FieldType = "adBSTR"
        Case adChapter: FieldType = "adChapter"
        Case adChar: FieldType = "adChar"
        Case adCurrency: FieldType = "adCurrency"
        Case adDate: FieldType = "adDate"
        Case adDBDate: FieldType = "adDBDate"
        Case adDBTime: FieldType = "adDBTime"
        Case adDBTimeStamp: FieldType = "adDBTimeStamp"
        Case adDecimal: FieldType = "adDecimal"
        Case adDouble: FieldType = "adDouble"
        Case adEmpty: FieldType = "adEmpty"
        Case adError: FieldType = "adError"
        Case adFileTime: FieldType = "adFileTime"
        Case adGUID: FieldType = "adGUID"
        Case adIDispatch: FieldType = "adIDispatch"
        Case adInteger: FieldType = "adInteger"
        Case adIUnknown: FieldType = "adIUnknown"
        Case adLongVarBinary: FieldType = "adLongVarBinary"
        Case adLongVarChar: FieldType = "adLongVarChar"
        Case adLongVarWChar: FieldType = "adLongVarWChar"
        Case adNumeric: FieldType = "adNumeric"
        Case adPropVariant: FieldType = "adPropVariant"
        Case adSingle: FieldType = "adSingle"
        Case adSmallInt: FieldType = "adSmallInt"
        Case adTinyInt: FieldType = "adTinyInt"
        Case adUnsignedBigInt: FieldType = "adUnsignedBigInt"
        Case adUnsignedInt: FieldType = "adUnsignedInt"
        Case adUnsignedSmallInt: FieldType = "adUnsignedSmallInt"
        Case adUnsignedTinyInt: FieldType = "adUnsignedTinyInt"
        Case adUserDefined: FieldType = "adUserDefined"
        Case adVarBinary: FieldType = "adVarBinary"
        Case adVarChar: FieldType = "adVarChar"
        Case adVariant: FieldType = "adVariant"
        Case adVarNumeric: FieldType = "adVarNumeric"
        Case adVarWChar: FieldType = "adVarWChar"
        Case adWChar: FieldType = "adWChar"
    End Select

End Function

Private Function SignBetweenAndValue(iListItem As Integer) As String
    Dim s As String
    Dim X As String
    s = FieldType(iListItem)
    Select Case s
        Case "adChar", "adWChar", "adVarWChar", "adVarChar", "adLongVarChar", "adGUID", _
            "adLongVarWChar", "adChar"
            strBetweenSQL = "'"
        Case "adDate", "adDBDate", "adDBTime", "adDBTimeStamp"
            strBetweenSQL = "#"
        Case Else
            strBetweenSQL = ""
    End Select
    SignBetweenAndValue = strBetweenSQL
End Function

'fungsi yang digunakan untuk membaca template baris demi baris
Private Function OpenTextFileAndTranslate(Filename As String) As String
    Dim nFileNum As Integer, sText As String
    Dim sNextLine As String, lLineCount As Long
    nFileNum = FreeFile
    Open Filename For Input As nFileNum
        lLineCount = 1
        Do While Not EOF(nFileNum)
            Line Input #nFileNum, sNextLine
            Compile sNextLine
            sNextLine = sNextLine & vbCrLf
            sText = sText & sNextLine
        Loop
        OpenTextFileAndTranslate = sText
    Close nFileNum
End Function

'fungsi yang digunakan untuk menterjemahkan template
Private Sub Compile(Code As String)

    Dim ObjectOrMethod As String
    Dim strCode As String
    Dim strObject As String
    Dim strProperty As Variant
    Dim ctl As VBControl
    Dim f As VBForm
    Dim i As Integer

    If (Trim(Left(Code, 1)) = "/") Or (Trim(Code) = "") Then
        Exit Sub 'do nothing with comments or blank line
    End If

    strObject = Split(Code, "=")(0)
    ObjectOrMethod = Trim$(Split(Code, "=")(0))
    Set f = VBInstance.SelectedVBComponent.Designer

    Select Case ObjectOrMethod
        Case "Form"
            'Do with Form
        Case Else
            Set ctl = f.VBControls.Add(ObjectOrMethod)
    End Select

    strCode = Trim$(Right$(Code, Len(Code) - Len(strObject)))
    strCode = Right$(strCode, Len(strCode) - 1)
    strProperty = Split(strCode, ",")

    For i = LBound(strProperty) To UBound(strProperty)
        Select Case ObjectOrMethod
            Case "InsertOCX"
                InsertOCX Trim$(Split(strProperty(i), "=")(0))
            Case "Form"
                With VBInstance.SelectedVBComponent
                    .Properties(Trim$(Split(strProperty(i), "=")(0))) = CVar((Trim$(Split(strProperty(i), "=")(1))))
                End With
            Case "SetContainer"
            Case Else
                With ctl
                    .Properties(Trim$(Split(strProperty(i), "=")(0))) = CVar((Trim(Split(strProperty(i), "=")(1))))
                End With
        End Select
    Next

    Select Case ObjectOrMethod
        Case "SetContainer"
            Call SetContaier(Code)
    End Select

End Sub

Private Sub SetContaier(Code As String)
    Dim f As VBForm
    Dim i As Integer
    Dim cContainer As VBControl
    Set f = VBInstance.SelectedVBComponent.Designer
    Dim sCode As String
    Dim sContainer() As String
    sCode = Trim$(Split(Code, "=")(1))
    sContainer = Split(sCode, ",")
    Dim sObject As String
    Dim iIndex As Integer

    If IsHaveIndex(sContainer(0)) Then
        sObject = Split(sContainer(0), "(")(0)
        iIndex = CInt(Replace(Split(sContainer(0), "(")(1), ")", ""))
        Set cContainer = f.VBControls(sObject, iIndex)
    Else
        Set cContainer = f.VBControls(sContainer(0))
    End If
    For i = LBound(sContainer) To UBound(sContainer)
        If i > 0 Then
            If IsHaveIndex(sContainer(i)) Then
                sObject = Split(sContainer(i), "(")(0)
                iIndex = CInt(Replace(Split(sContainer(i), "(")(1), ")", ""))
                Set f.VBControls(sObject, iIndex).Container = cContainer
            Else
                Set f.VBControls(sContainer(i)).Container = cContainer
            End If
        End If
    Next
End Sub

Private Function IsHaveIndex(Code As String) As Boolean
    IsHaveIndex = (InStr(1, Code, "(") > 0)
End Function

Private Function SignBetween(iListItem As Integer, Field As String) As String
    Dim s As String
    Dim X As String
    X = Field
    s = FieldType(iListItem)
    Select Case s
        Case "adChar", "adWChar", "adVarWChar", "adVarChar", "adLongVarChar", "adGUID", _
            "adLongVarWChar", "adChar"
            X = "'nn(" & X & ")'"
        Case "adDate", "adDBDate", "adDBTime", "adDBTimeStamp"
            X = "GetDBDate(" & X & ")"
        Case Else
            X = X
    End Select
    SignBetween = X
End Function

Private Function SQLInsert() As String

    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = DBase
    Dim c As ADOX.Column

    Dim i As Integer
    Dim s As String
    Dim X As Integer
    Dim r As String
    Dim tType As Integer
    Dim TableName As String

    TableName = GetValidName(GetTabelValidName(cboTables.Text))

    X = lstSelected.ListCount - 1

    For i = 0 To lstSelected.ListCount - 1
        tType = cat.Tables(GetTabelValidName(cboTables.Text)).Columns(lstSelected.List(i)).Type
        r = r & vbTab & Replace(SignBetween(tType, "txt" & lstSelected.List(i) & ".Text") & IIf(i <> X, ", " & vbCrLf, ""), "nn(", "sq(")
        s = s & vbTab & lstSelected.List(i) & IIf(i <> X, ", " & vbCrLf, "")
    Next

    SQLInsert = ParseSQLInsert("INSERT INTO " & TableName & " (" & vbCrLf & s & ")" & vbCrLf & " VALUES(" & vbCrLf & r & ")")

End Function

Private Function SQLUpdate() As String
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = DBase
    Dim c As ADOX.Column

    Dim i As Integer
    Dim s As String
    Dim X As Integer
    Dim r As String
    Dim tType As Integer
    Dim TableName As String

    TableName = GetValidName(GetTabelValidName(cboTables.Text))

    X = lstSelected.ListCount - 1

    For i = 0 To lstSelected.ListCount - 1
        tType = cat.Tables(GetTabelValidName(cboTables.Text)).Columns(lstSelected.List(i)).Type
        r = Replace(SignBetween(tType, "txt" & lstSelected.List(i) & ".Text"), "nn(", "sq(")
        s = s & vbTab & lstSelected.List(i) & IIf(i <> X, " = " & r & ", " & vbCrLf, " = " & r)
    Next

    SQLUpdate = ParseSQLUpdate("UPDATE " & TableName & vbCrLf & "SET " & vbCrLf & s) & "    strSQL = strSQL & " & Chr(34) & GetWhereString & vbCrLf

End Function

Private Function GetWhereString() As String
    Dim TableName As String
    Dim strPrimaryKey As String
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = DBase
    TableName = GetValidName(GetTabelValidName(cboTables.Text))
    strPrimaryKey = SignBetweenAndValue(cat.Tables(TableName).Columns(cmbUniqID.Text).Type)
    Set cat = Nothing
    If strPrimaryKey = "" Then
        GetWhereString = "WHERE " & cmbUniqID.Text & "=" & Chr(34) & " & ID"
    Else
        GetWhereString = "WHERE " & cmbUniqID.Text & "=" & strPrimaryKey & Chr(34) & " & ID & " & Chr(34) & strPrimaryKey & Chr(34)
    End If
End Function

Private Function ParseSQLInsert(StrSQL As String) As String

    Debug.Print StrSQL
    
    Dim i As Integer, s As String, r As String
    Dim X As String, b As Boolean, c As Integer, a() As String

    s = StrSQL
    a = Split(s, vbCrLf)

    For i = LBound(a) To UBound(a)

        X = a(i)

        If InStr(1, a(i), "VALUES(") > 0 Then
            c = 0
            a(i) = Chr(34) & a(i)
            b = True
        End If

        If b = True Then
            If IsStringEnd(X, ")')") Then
                a(i) = Replace(a(i), "'", "")
                a(i) = Replace(a(i), "))", ")")
                a(i) = Chr(34) & "   '" & Chr(34) & " & " & a(i) & " & " & Chr(34) & "')"
            ElseIf InStr(1, X, vbTab & "'") > 0 Then
                If InStr(1, X, "')") > 0 Then
                    a(i) = Replace(a(i), "'", "")
                    a(i) = Replace(a(i), ")", "")
                    a(i) = Chr(34) & "   '" & Chr(34) & " & " & a(i) & " & " & Chr(34) & "')"
                ElseIf IsStringEnd(X, ")") Then
                    a(i) = Replace(a(i), "'", "")
                    a(i) = Chr(34) & "   '" & Chr(34) & " & " & a(i) & " & " & Chr(34) & ")"
                Else
                    a(i) = Replace(a(i), ",", "")
                    a(i) = Replace(a(i), "'", "")
                    a(i) = Chr(34) & "   '" & Chr(34) & " & " & a(i) & " & " & Chr(34) & "', "
                End If
            Else
                If c > 0 Then
                    If IsStringStart(X, vbTab & "GetDBDate") Then
                        If IsStringEnd(X, "), ") Then
                            a(i) = Chr(34) & "    " & Chr(34) & " & " & Replace(a(i), ",", "") & " & " & Chr(34) & ", "
                        ElseIf IsStringEnd(X, "))") Then
                            a(i) = Chr(34) & "    " & Chr(34) & " & " & Replace(a(i), "))", ")") & " & " & Chr(34) & ")"
                        End If
                    ElseIf InStr(1, X, ")") > 0 Then
                        a(i) = Chr(34) & "    " & Chr(34) & " & " & Replace(a(i), ")", "") & " & " & Chr(34) & ")"
                    Else
                        a(i) = Chr(34) & "    " & Chr(34) & " & " & Replace(a(i), ",", "") & " & " & Chr(34) & ", "
                    End If
                End If
            End If
        End If

    a(i) = Replace(a(i), vbTab, "    ")
    If b = True Then
        r = r & "    strSQL = strSQL & " & a(i) & Chr(34) & "'" & c & vbCrLf
    Else
        r = r & "    strSQL = strSQL & " & Chr(34) & a(i) & Chr(34) & "'" & c & vbCrLf
    End If
    c = c + 1

Next

ParseSQLInsert = r

End Function

Private Function ParseSQLUpdate(StrSQL As String) As String

    Debug.Print StrSQL

    Dim i As Integer, s As String, r As String
    Dim X As String, b As Boolean, c As Integer
    Dim a() As String, f() As String, strSpace As String
    Dim strKoma As String

    s = StrSQL
    a = Split(s, vbCrLf)

    For i = LBound(a) To UBound(a)
        X = a(i)
        f = Split(a(i), "=")

        If UBound(f) > 0 Then
            f(1) = Replace(f(1), "'", "")
            f(1) = Replace(f(1), ",", "")
        End If

        strSpace = IIf(InStr(1, a(i), "WHERE") > 0, "", "   ")

        If InStr(1, a(i), "SET") > 0 Then
            c = 0
            a(i) = Chr(34) & a(i)
            b = True
        End If
        
        If i < UBound(a) Then
            If InStr(1, X, "'") > 0 Then
                strKoma = IIf(Right(Trim$(X), 1) = ",", "', ", "' ")
            Else
                strKoma = ", "
            End If
        End If
        If i = UBound(a) Then
            If InStr(1, X, "'") > 0 Then
                strKoma = "'"
            Else
                strKoma = " "
            End If
        End If
        If b = True Then
            If InStr(1, X, "'") > 0 Then
                If InStr(1, X, "')") > 0 Then
                    a(i) = Chr(34) & "   '" & Chr(34) & " & " & f(1) & " & " & Chr(34) & "')"
                Else
                    a(i) = Chr(34) & strSpace & f(0) & "= '" & Chr(34) & " & " & f(1) & " & " & Chr(34) & strKoma
                End If
            Else
                If c > 0 Then
                    If i = UBound(a) Then
                        a(i) = Chr(34) & strSpace & f(0) & "= " & Chr(34) & " & " & f(1) & " & " & Chr(34) & strKoma
                    Else
                        a(i) = Chr(34) & strSpace & f(0) & "= " & Chr(34) & " & " & f(1) & " & " & Chr(34) & strKoma
                    End If
                End If
             End If

        End If
        
        a(i) = Replace(a(i), vbTab, "  ")
    
        If b = True Then
            r = r & "    strSQL = strSQL & " & a(i) & Chr(34) & "'" & c & vbCrLf
        Else
            r = r & "    strSQL = strSQL & " & Chr(34) & a(i) & " " & Chr(34) & "'" & c & vbCrLf
        End If
    
        c = c + 1
    Next

    ParseSQLUpdate = r

End Function

Private Function IsStringStart(s As String, X As String) As Boolean
    IsStringStart = (Left$(s, Len(X)) = X)
End Function

Private Function IsStringEnd(s As String, X As String) As Boolean
    IsStringEnd = (Right$(s, Len(X)) = X)
End Function

Private Sub AddVBComponentFromFile(Path As String)
    Call VBInstance.ActiveVBProject.VBComponents.AddFile(Path)
End Sub

Private Function GetDefinedSize(Field) As Integer
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = DBase
    Dim c As ADOX.Column
    GetDefinedSize = cat.Tables(GetTabelValidName(cboTables.Text)).Columns(Field).DefinedSize
    Set cat = Nothing
End Function
