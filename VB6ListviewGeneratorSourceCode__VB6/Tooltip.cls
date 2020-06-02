VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'************************************************************
'
' Tool Tip Class
'
' 17-NOV-2004
'
' Mark Mokoski
' C & M Telephone
' markm@cmtelephone.com
' www.rjillc.com
'
' Class for the creation of Rectangular and Balloon Multiline Tool Tips.
'
' See Code for details on Syntax, variables and constants.
'
' This Class Module works by sub classing the parent control.
' Works with most of the common controls, the ones it will not work
' with are controls that do not support tool tips themselves.
' There are some controls that support tool tips but complain with this
' sub classing module (only Microsoft knows why!).
' Do not use the native tool tip with the parent control and this together,
' you get overlapping tool tips!
'
' This class module is also know to work on the following OS's (all Win32)
' Windows Me
' Windows 2000
' Windows XP
' MSDN Docs state that IE 5 or higher is needed for Balloon tips
'
' It's usefull as it is written, so give it a try!
'
' As brought to my attention on 01-DEC-2004
' Parts of this class was writen by Eidos on PSC
' I found it in bits on other sites and did the usual inprovments
' and changes to for my needs.
' So, credit for the core of this class goes to Eidos at PSC
' Eidos's code can be founfd at:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=28419&lngWId=1
'**************************************************************
'
' Public Methods:
'
' .CreateBalloon     Create Balloon Tool Tip
' .CreateTip         Create Rectangular Tool Tip
' .Remove            Kills Tool Tip object
'
' Public Properties:
'
' .Active            Boolean    Activate (visible)/Deactivate (hide) tool tip
' .ParentControl     Long       hWnd of Control that the tool tip is subclassed (displays on)to
' .Style             Enum Type  Tool Tip style, Rectangular or Balloon
' .Centered          Boolean    Tool Tip is centered on parent control when visible
' .Icon              Enum Type  Tool Tip Icon used when tool tip has a title
' .Title             Text       Tool Tip title text
' .Fore Color        Long       Tool Tip text color and border color if Balloon tip
' .BackColor         Long       Tool Tip Background color
' .TipText           Text       Tool Tip text
' .hWnd              Long       Tool Tip Windows Handle (.hWnd) READ ONLY
' .ShowTime          Long       Tool Tip visible time in mSec (10000 = 10 Sec)
'
' Sample code:
'************************************************************
'    Option Explicit
'
'   'Make new tool tip object for this project
'   Dim cmdBrowsePrjTip              As New clsTooltips
'
'   Private Sub Form_Load()
'
'   'Make the complete Tool Tip, text, title, icon
'   cmdBrowsePrjTip.CreateBalloon cmdBrowsePrj, "I turned off all the Tool Tips " + vbCrLf + "Click to restore Tool Tips", "Tool Tips are OFF", tipIconWarning
'
'   End Sub
'
'   Put this delare in the Sub_Main module...
'       'Int Common Controls Lib
'       Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
'
'   Then call the sub in the Sub_Main
'       'Int Common Controls Lib
'       InitCommonControls
'
'************************************************************

Option Explicit
'************************************************************
' API Functions
'************************************************************
Private Declare Function CreateWindowEx Lib "USER32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ShowWindow Lib "USER32.DLL" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'Int Common Controls Lib, put in startup module and execute
'"InitCommonControls" in Sub_Main
'Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

'************************************************************
' Constants
'************************************************************
'Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const WS_POPUP As Long = &H80000000
Private Const WS_BORDER As Long = &H800000

'Tooltip Window Constants
Private Const TTF_IDISHWND                As Long = &H1
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = WM_USER + 4
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = WM_USER + 12
Private Const TTM_SETMAXTIPWIDTH = WM_USER + 24
Private Const TTM_SETTIPBKCOLOR = WM_USER + 19
Private Const TTM_SETTIPTEXTCOLOR = WM_USER + 20
Private Const TTM_SETTITLE = WM_USER + 32
Private Const TTM_SETDELAYTIME            As Long = WM_USER + 3
Private Const TTM_SETMARGIN               As Long = WM_USER + 26
Private Const TTM_UPDATE                  As Long = (WM_USER + 29)
Private Const TTS_NOPREFIX = &H2
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TTDT_AUTOPOP                As Long = 2
Private Const TTDT_AUTOMATIC              As Long = 0
Private Const TTDT_INITIAL                As Long = 3
Private Const TTDT_RESHOW                 As Long = 1
'Tool Tip Icons
Private Const TTI_ERROR                   As Long = 3
Private Const TTI_INFO                    As Long = 1
Private Const TTI_NONE                    As Long = 0
Private Const TTI_WARNING                 As Long = 2
'Tool Tip API Class
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

'************************************************************
' Types
'************************************************************

'Windows API Types

Private Type RECT
    Left                              As Long
    Top                               As Long
    Right                             As Long
    Bottom                            As Long
End Type

'Tooltip Window Types

Private Type TOOLINFO
    lSize                             As Long
    lFlags                            As Long
    TiphWnd                           As Long
    lId                               As Long
    lpRect                            As RECT
    hInstance                         As Long
    lpszText                          As String
    lParam                            As Long
End Type

'************************************************************
'Local Class variables and Data
'************************************************************

'Local variables to hold property values
Private toolBackColor                     As Long
Private toolTitle                         As String
Private toolForeColor                     As Long
Private toolParentControl                 As Long
Private toolIcon                          As toolIconType
Private toolCentered                      As Boolean
Private toolStyle                         As toolStyleEnum
Private toolTipText                       As String
Private toolActive                        As Boolean
Private toolLineLength                    As Integer
Private toolhWnd                          As Long
Private toolShowTime                      As Long

'Private Data for Class
Private TiphWnd                           As Long
Private ti                                As TOOLINFO

Public Enum toolIconType
    tipNoIcon = TTI_NONE            '= 0
    tipIconInfo = TTI_INFO          '= 1
    tipIconWarning = TTI_WARNING    '= 2
    tipIconError = TTI_ERROR        '= 3

End Enum

Public Enum toolStyleEnum
    styleStandard = 0
    styleBalloon = 1
End Enum

Public Function CreateBalloon(ByVal Parent As Object, Text As String, Optional Title As String = vbNullString, Optional Icon As Integer = TTI_NONE)

    'Used to create a Balloon Tool Tip object.
    'Pass needed parameters with call.
    'Syntax: object.CreateBalloon ParentControl, Tip Text, Title, Icon
    'Title and Icon are optional, but you cant have an Icon without a Title
    'Title can be just a space, just not an empty string

    toolStyle = styleBalloon
    toolParentControl = Parent.hWnd
    toolTipText = Text

    If Title = vbNullString Then
        toolTitle = ""
    Else
        toolTitle = Title
    End If

    If Icon = Icon > TTI_ERROR Then
        toolIcon = TTI_NONE
    Else
        toolIcon = Icon
    End If

    Call Create
    Active = True

End Function

Public Function CreateTip(ByVal Parent As Object, Text As String, Optional Title As String = vbNullString, Optional Icon As Integer = TTI_NONE)

    'Used to create a Standard (rectangle) Tool Tip object.
    'Pass needed parameters with call.
    'Syntax: object.CreateTip ParentControl, Tip Text, Title, Icon
    'Title and Icon are optional, but you cant have an Icon without a Title
    'Title can be just a space, just not an empty string

    toolStyle = styleStandard
    toolParentControl = Parent.hWnd
    toolTipText = Text

    If Title = vbNullString Then
        toolTitle = ""
    Else
        toolTitle = Title
    End If

    If Icon = vbNull Or Icon > TTI_ERROR Then
        toolIcon = TTI_NONE
    Else
        toolIcon = Icon
    End If

    Call Create
    Active = True

End Function

Private Sub Create()

    'Private sub used with Create and Update subs/functions

    Dim lpRect               As RECT
    Dim lWinStyle            As Long

    'If Tool Tip already made, destroy it and reconstruct

    If TiphWnd <> 0 Then
        DestroyWindow TiphWnd
    End If

    lWinStyle = WS_POPUP Or WS_BORDER Or TTS_ALWAYSTIP Or TTS_NOPREFIX

    'Create baloon style if desired

    If toolStyle = styleBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON

    'The parent control has to be set first

    If toolParentControl <> &H0 Then
        TiphWnd = CreateWindowEx(0&, _
        TOOLTIPS_CLASSA, _
        vbNullString, _
        lWinStyle, _
        CW_USEDEFAULT, _
        CW_USEDEFAULT, _
        CW_USEDEFAULT, _
        CW_USEDEFAULT, _
        toolParentControl, _
        0&, _
        App.hInstance, _
        0&)

        toolhWnd = TiphWnd

        'Make our tooltip window a topmost window
        SetWindowPos TiphWnd, _
        HWND_TOPMOST, _
        0&, _
        0&, _
        0&, _
        0&, _
        SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE

        'Get the rectangle of the parent control
        GetClientRect toolParentControl, lpRect

        'Now set up our tooltip info structure

        With ti
            'If we want it centered, then set that flag

            If toolCentered Then
                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
            Else
                .lFlags = TTF_SUBCLASS
            End If

            ''set the hwnd prop to our parent control's hwnd
            .TiphWnd = toolParentControl
            .lId = 0
            .hInstance = App.hInstance
            .lpszText = toolTipText
            .lpRect = lpRect
        End With

        'Add the tooltip structure
        SendMessage TiphWnd, TTM_ADDTOOLA, 0&, ti

        'Set Max Width to 32 characters, and enable Multi Line Tool Tips
        'SendMessage TiphWnd, TTM_SETMAXTIPWIDTH, 0&, &H20
        SendMessage TiphWnd, TTM_SETMAXTIPWIDTH, 0, 300

        'If we want a title or we want an icon

        'If toolTitle <> vbNullString Or toolIcon <> tipNoIcon Then
        SendMessage TiphWnd, TTM_SETTITLE, CLng(toolIcon), ByVal toolTitle
        'End If

        If toolForeColor <> Empty Then
            '0 (zero) or Null is seen by the API as the default color
            'See ForeColor property for more datails
            SendMessage TiphWnd, TTM_SETTIPTEXTCOLOR, toolForeColor, 0&
        End If

        If toolBackColor <> Empty Then
            '0 (zero) or Null is seen by the API as the default color
            'See BackColor property for more datails
            SendMessage TiphWnd, TTM_SETTIPBKCOLOR, toolBackColor, 0&
        End If

    End If

End Sub

Private Sub UpDate()

    Dim aTemp                As Boolean

    'Used to update tooltip parameters that require reconfiguration of
    'subclass to envoke

    'Get current Atcive state
    aTemp = Active
    'Refresh the object
    Call Create
    'Restore the Active state
    Active = aTemp

End Sub

Public Property Let Active(ByVal tooldata As Boolean)

    'If True, activate (show) tool tip, False deactivate (hide) tool tip
    'Syntax: object.active= true/false
    toolActive = tooldata
    SendMessage TiphWnd, TTM_ACTIVATE, CInt(toolActive), ti

End Property

Public Property Get Active() As Boolean

    'Retrieving value of a property, Boolean responce (true/false)
    'Syntax: BooleanVar = object.Active
    Active = toolActive

End Property

Public Property Set ParentControl(ByVal tooldata As Object)

    'Assigning an Object to the property, set to parent object(control)
    'that the Tool Tip will "pop" up from
    'Syntax: Set object.ParentControl = ParentObject
    toolParentControl = tooldata.hWnd
    UpDate

End Property

Public Property Get ParentControl() As Long

    'Retrieving value of a property, returns Long Windows Handle (hWnd)of Parent.
    'Syntax: ObjectVar =  object.ParentControl
    ParentControl = toolParentControl

End Property

Public Property Let Style(ByVal tooldata As toolStyleEnum)

    'Assigning a value to the property, set style param Standard or Balloon
    'Syntax: object.Style = style
    toolStyle = tooldata
    UpDate

End Property

Public Property Get Style() As toolStyleEnum

    'Retrieving value of a property, returns string.
    'Syntax: StringVar = object.Style
    Style = toolStyle

End Property

Public Property Let Centered(ByVal tooldata As Boolean)

    'Assigning a value to the property, Set Boolean true/false if ToolTip
    'is centered on the parent control
    'Syntax: object.Centered = true/false
    toolCentered = tooldata
    UpDate

End Property

Public Property Get Centered() As Boolean

    'Retrieving value of a property, returns Boolean true/false.
    'Syntax: BooleanVar =  object.Centered
    Centered = toolCentered
    UpDate

End Property

Public Property Let Icon(ByVal tooldata As toolIconType)

    'Assigning a value to the property, set icon style with type var.
    'Syntax: object.Icon = iconStyle
    'Icon Styles are: INFO, WARNING and ERROR (tipNoIcom, tipIconInfo, tipIconWarning, tipIconError)
    toolIcon = tooldata

    'If tipHwnd <> 0 And toolTitle <> Empty And toolIcon <> tipNoIcon Then

    If TiphWnd <> 0 Then
        SendMessage TiphWnd, TTM_SETTITLE, CLng(toolIcon), ByVal toolTitle
    End If

    UpDate

End Property

Public Property Get Icon() As toolIconType

    'Retrieving value of a property, returns string.
    'Syntax: StringVar =  object.Icon
    Icon = toolIcon

End Property

Public Property Let ForeColor(ByVal tooldata As Long)

    'Assigning a value to the property, set RGB value as Long.
    'Syntax: object.ForeColor = RGB (as Long)
    'Since 0 is Black (no RGB), and the API thinks 0 is
    'the default color ("off" yellow),
    'we need to "fudge" Black a bit (yes set bit "1" to "1",)
    'I couldn't resist the pun!
    'So, in module or form code, if setting to Black, make it "1"
    'if restoring the default color, make it "0"
    'Syntax: object.ForeColor = RGB(as long)

    toolForeColor = tooldata

    If TiphWnd <> 0 Then
        SendMessage TiphWnd, TTM_SETTIPTEXTCOLOR, toolForeColor, 0&
    End If

    UpDate

End Property

Public Property Get ForeColor() As Long

    'Retrieving value of a property, returns RGB value as Long.
    'Syntax: LongVar =  object.ForeColor
    ForeColor = toolForeColor

End Property

Public Property Let Title(ByVal tooldata As String)

    'Assigning a value to the property, set as string.
    'Syntax: object.Title = StringVar
    toolTitle = tooldata

    'If tipHwnd <> 0 And toolTitle <> Empty And toolIcon <> tipNoIcon Then

    If TiphWnd <> 0 Then
        SendMessage TiphWnd, TTM_SETTITLE, CLng(toolIcon), ByVal toolTitle
    End If

    UpDate

End Property

Public Property Get Title() As String

    'Retrieving value of a property, returns string.
    'Syntax: StringVar =  object.Title
    Title = toolTitle

End Property

Public Property Let BackColor(ByVal tooldata As Long)

    'Assigning a value to the property, set RGB value as Long.
    'Syntax: object.BackColor = RGB (as Long)
    'Since 0 is Black (no RGB), and the API thinks 0 is
    'the default color ("off" yellow),
    'we need to "fudge" Black a bit (yes set bit "1" to "1",)
    'I couldn't resist the pun!
    'So, in module or form code, if setting to Black, make it "1"
    'if restoring the default color, make it "0"

    toolBackColor = tooldata

    If TiphWnd <> 0 Then
        SendMessage TiphWnd, TTM_SETTIPBKCOLOR, toolBackColor, 0&
    End If

    UpDate

End Property

Public Property Get BackColor() As Long

    'Retrieving value of a property, returns RGB as Long.
    'Syntax: LongVar =  object.BackColor

    BackColor = toolBackColor

End Property

Public Property Let TipText(ByVal tooldata As String)

    'Assigning a value to the property, Set as String.
    'Syntax: object.TipText = StringVar
    'Multi line Tips are enabled in the Create sub.
    'To change lines, just add a vbCrLF between text
    'ex. object.TipText= "Line 1 text" & vbCrLF & "Line 2 text"

    toolTipText = tooldata
    ti.lpszText = toolTipText

    If TiphWnd <> 0 Then
        SendMessage TiphWnd, TTM_UPDATETIPTEXTA, 0&, ti
    End If

    UpDate

End Property

Public Property Get TipText() As String

    'Retrieving value of a property, returns string.
    'Syntax: StringVar =  object.TipText
    TipText = toolTipText

End Property

Public Property Get hWnd() As Long

    'Retrive Windows Handle of the Tool Tip
    'Syntax: LongVar = object.ToolhWnd
    hWnd = toolhWnd

End Property

Public Function Remove() As Boolean

    'Kills Tool Tip Object
    Tool_Tip_Terminate

End Function

Private Sub Tool_Tip_Terminate()

    If TiphWnd <> 0 Then
        DestroyWindow TiphWnd
    End If

End Sub

    



