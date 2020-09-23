VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SPWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Window object."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_hWnd As Long  ' hWnd to subclass
Private m_SubclassRef As Long  ' Old window procedure
Private m_Subclassed As Boolean  ' Is the window subclassed?

Public Event BeforeMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Cancel As Boolean)
Attribute BeforeMessage.VB_MemberFlags = "200"
Public Event AfterMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

' Used for Subclass method
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)

' Used to confirm that a window exists
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

' Constants for messages that can be trapped
Public Enum SPMessageConstants
  ABM_ACTIVATE = &H6               '  lParam == TRUE/FALSE means activate/deactivate
  ABM_GETAUTOHIDEBAR = &H7
  ABM_GETSTATE = &H4
  ABM_GETTASKBARPOS = &H5
  ABM_NEW = &H0
  ABM_QUERYPOS = &H2
  ABM_REMOVE = &H1
  ABM_SETAUTOHIDEBAR = &H8          '  this can fail at any time.  MUST check the result
  ABM_SETPOS = &H3
  ABM_WINDOWPOSCHANGED = &H9
  ABN_FULLSCREENAPP = &H2
  ABN_POSCHANGED = &H1
  ABN_STATECHANGE = &H0
  ABN_WINDOWARRANGE = &H3 '  lParam == TRUE means hide
  BM_GETCHECK = &HF0
  BM_GETSTATE = &HF2
  BM_SETCHECK = &HF1
  BM_SETSTATE = &HF3
  BM_SETSTYLE = &HF4
  BN_CLICKED = 0
  BN_DISABLE = 4
  BN_DOUBLECLICKED = 5
  BN_HILITE = 2
  BN_PAINT = 1
  BN_UNHILITE = 3
  CB_ADDSTRING = &H143
  CB_DELETESTRING = &H144
  CB_DIR = &H145
  CB_ERR = (-1)
  CB_ERRSPACE = (-2)
  CB_FINDSTRING = &H14C
  CB_FINDSTRINGEXACT = &H158
  CB_GETCOUNT = &H146
  CB_GETCURSEL = &H147
  CB_GETDROPPEDCONTROLRECT = &H152
  CB_GETDROPPEDSTATE = &H157
  CB_GETEDITSEL = &H140
  CB_GETEXTENDEDUI = &H156
  CB_GETITEMDATA = &H150
  CB_GETITEMHEIGHT = &H154
  CB_GETLBTEXT = &H148
  CB_GETLBTEXTLEN = &H149
  CB_GETLOCALE = &H15A
  CB_INSERTSTRING = &H14A
  CB_LIMITTEXT = &H141
  CB_MSGMAX = &H15B
  CB_OKAY = 0
  CB_RESETCONTENT = &H14B
  CB_SELECTSTRING = &H14D
  CB_SETCURSEL = &H14E
  CB_SETEDITSEL = &H142
  CB_SETEXTENDEDUI = &H155
  CB_SETITEMDATA = &H151
  CB_SETITEMHEIGHT = &H153
  CB_SETLOCALE = &H159
  CB_SHOWDROPDOWN = &H14F
  CBN_CLOSEUP = 8
  CBN_DBLCLK = 2
  CBN_DROPDOWN = 7
  CBN_EDITCHANGE = 5
  CBN_EDITUPDATE = 6
  CBN_ERRSPACE = (-1)
  CBN_KILLFOCUS = 4
  CBN_SELCHANGE = 1
  CBN_SELENDCANCEL = 10
  CBN_SELENDOK = 9
  CBN_SETFOCUS = 3
  CDM_FIRST = (&H400 + 100)
  CDM_GETFILEPATH = (CDM_FIRST + &H1)
  CDM_GETFOLDERIDLIST = (CDM_FIRST + &H3)
  CDM_GETFOLDERPATH = (CDM_FIRST + &H2)
  CDM_GETSPEC = (CDM_FIRST + &H0)
  CDM_HIDECONTROL = (CDM_FIRST + &H5)
  CDM_LAST = (&H400 + 200)
  CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)
  CDM_SETDEFEXT = (CDM_FIRST + &H6)
  CTLCOLOR_BTN = 3
  CTLCOLOR_DLG = 4
  CTLCOLOR_EDIT = 1
  CTLCOLOR_LISTBOX = 2
  CTLCOLOR_MAX = 8   '  three bits max
  CTLCOLOR_MSGBOX = 0
  CTLCOLOR_SCROLLBAR = 5
  CTLCOLOR_STATIC = 6
  EM_CANUNDO = &HC6
  EM_EMPTYUNDOBUFFER = &HCD
  EM_FMTLINES = &HC8
  EM_GETFIRSTVISIBLELINE = &HCE
  EM_GETHANDLE = &HBD
  EM_GETLINE = &HC4
  EM_GETLINECOUNT = &HBA
  EM_GETMODIFY = &HB8
  EM_GETPASSWORDCHAR = &HD2
  EM_GETRECT = &HB2
  EM_GETSEL = &HB0
  EM_GETTHUMB = &HBE
  EM_GETWORDBREAKPROC = &HD1
  EM_LIMITTEXT = &HC5
  EM_LINEFROMCHAR = &HC9
  EM_LINEINDEX = &HBB
  EM_LINELENGTH = &HC1
  EM_LINESCROLL = &HB6
  EM_REPLACESEL = &HC2
  EM_SCROLL = &HB5
  EM_SCROLLCARET = &HB7
  EM_SETHANDLE = &HBC
  EM_SETMODIFY = &HB9
  EM_SETPASSWORDCHAR = &HCC
  EM_SETREADONLY = &HCF
  EM_SETRECT = &HB3
  EM_SETRECTNP = &HB4
  EM_SETSEL = &HB1
  EM_SETTABSTOPS = &HCB
  EM_SETWORDBREAKPROC = &HD0
  EM_UNDO = &HC7
  EN_CHANGE = &H300
  EN_ERRSPACE = &H500
  EN_HSCROLL = &H601
  EN_KILLFOCUS = &H200
  EN_MAXTEXT = &H501
  EN_SETFOCUS = &H100
  EN_UPDATE = &H400
  EN_VSCROLL = &H602
  LB_ADDFILE = &H196
  LB_ADDSTRING = &H180
  LB_CTLCODE = 0&
  LB_DELETESTRING = &H182
  LB_DIR = &H18D
  LB_ERR = (-1)
  LB_ERRSPACE = (-2)
  LB_FINDSTRING = &H18F
  LB_FINDSTRINGEXACT = &H1A2
  LB_GETANCHORINDEX = &H19D
  LB_GETCARETINDEX = &H19F
  LB_GETCOUNT = &H18B
  LB_GETCURSEL = &H188
  LB_GETHORIZONTALEXTENT = &H193
  LB_GETITEMDATA = &H199
  LB_GETITEMHEIGHT = &H1A1
  LB_GETITEMRECT = &H198
  LB_GETLOCALE = &H1A6
  LB_GETSEL = &H187
  LB_GETSELCOUNT = &H190
  LB_GETSELITEMS = &H191
  LB_GETTEXT = &H189
  LB_GETTEXTLEN = &H18A
  LB_GETTOPINDEX = &H18E
  LB_INSERTSTRING = &H181
  LB_MSGMAX = &H1A8
  LB_OKAY = 0
  LB_RESETCONTENT = &H184
  LB_SELECTSTRING = &H18C
  LB_SELITEMRANGE = &H19B
  LB_SELITEMRANGEEX = &H183
  LB_SETANCHORINDEX = &H19C
  LB_SETCARETINDEX = &H19E
  LB_SETCOLUMNWIDTH = &H195
  LB_SETCOUNT = &H1A7
  LB_SETCURSEL = &H186
  LB_SETHORIZONTALEXTENT = &H194
  LB_SETITEMDATA = &H19A
  LB_SETITEMHEIGHT = &H1A0
  LB_SETLOCALE = &H1A5
  LB_SETSEL = &H185
  LB_SETTABSTOPS = &H192
  LB_SETTOPINDEX = &H197
  LBN_DBLCLK = 2
  LBN_ERRSPACE = (-2)
  LBN_KILLFOCUS = 5
  LBN_SELCANCEL = 3
  LBN_SELCHANGE = 1
  LBN_SETFOCUS = 4
  SB_BOTH = 3
  SB_BOTTOM = 7
  SB_CTL = 2
  SB_ENDSCROLL = 8
  SB_HORZ = 0
  SB_LEFT = 6
  SB_LINEDOWN = 1
  SB_LINELEFT = 0
  SB_LINERIGHT = 1
  SB_LINEUP = 0
  SB_PAGEDOWN = 3
  SB_PAGELEFT = 2
  SB_PAGERIGHT = 3
  SB_PAGEUP = 2
  SB_RIGHT = 7
  SB_THUMBPOSITION = 4
  SB_THUMBTRACK = 5
  SB_TOP = 6
  SB_VERT = 1
  SBM_ENABLE_ARROWS = &HE4 ' not in win3.1
  SBM_GETPOS = &HE1 ' not in win3.1
  SBM_GETRANGE = &HE3 ' not in win3.1
  SBM_SETPOS = &HE0 ' not in win3.1
  SBM_SETRANGE = &HE2 ' not in win3.1
  SBM_SETRANGEREDRAW = &HE6 ' not in win3.1
  WM_ACTIVATE = &H6
  WM_ACTIVATEAPP = &H1C
  WM_ASKCBFORMATNAME = &H30C
  WM_CANCELJOURNAL = &H4B
  WM_CANCELMODE = &H1F
  WM_CHANGECBCHAIN = &H30D
  WM_CHAR = &H102
  WM_CHARTOITEM = &H2F
  WM_CHILDACTIVATE = &H22
  WM_CHOOSEFONT_GETLOGFONT = (&H400 + 1)
  WM_CHOOSEFONT_SETFLAGS = (&H400 + 102)
  WM_CHOOSEFONT_SETLOGFONT = (&H400 + 101)
  WM_CLEAR = &H303
  WM_CLOSE = &H10
  WM_COMMAND = &H111
  'WM_COMMNOTIFY = &H44  ' no longer suported
  WM_COMPACTING = &H41
  WM_COMPAREITEM = &H39
  WM_CONTEXTMENU = &H7B
  WM_CONVERTREQUESTEX = &H108
  WM_COPY = &H301
  WM_COPYDATA = &H4A
  WM_CREATE = &H1
  WM_CTLCOLORBTN = &H135
  WM_CTLCOLORDLG = &H136
  WM_CTLCOLOREDIT = &H133
  WM_CTLCOLORLISTBOX = &H134
  WM_CTLCOLORMSGBOX = &H132
  WM_CTLCOLORSCROLLBAR = &H137
  WM_CTLCOLORSTATIC = &H138
  WM_CUT = &H300
  WM_DEADCHAR = &H103
  WM_DELETEITEM = &H2D
  WM_DESTROY = &H2
  WM_DESTROYCLIPBOARD = &H307
  WM_DEVMODECHANGE = &H1B
  WM_DRAWCLIPBOARD = &H308
  WM_DRAWITEM = &H2B
  WM_DROPFILES = &H233
  WM_ENABLE = &HA
  WM_ENDSESSION = &H16
  WM_ENTERIDLE = &H121
  WM_ENTERMENULOOP = &H211
  WM_ERASEBKGND = &H14
  WM_EXITMENULOOP = &H212
  WM_FONTCHANGE = &H1D
  WM_GETDLGCODE = &H87
  WM_GETFONT = &H31
  WM_GETHOTKEY = &H33
  WM_GETMINMAXINFO = &H24
  WM_GETTEXT = &HD
  WM_GETTEXTLENGTH = &HE
  WM_HOTKEY = &H312
  WM_HSCROLL = &H114
  WM_HSCROLLCLIPBOARD = &H30E
  WM_ICONERASEBKGND = &H27
  WM_IME_CHAR = &H286
  WM_IME_COMPOSITION = &H10F
  WM_IME_COMPOSITIONFULL = &H284
  WM_IME_CONTROL = &H283
  WM_IME_ENDCOMPOSITION = &H10E
  WM_IME_KEYDOWN = &H290
  WM_IME_KEYLAST = &H10F
  WM_IME_KEYUP = &H291
  WM_IME_NOTIFY = &H282
  WM_IME_SELECT = &H285
  WM_IME_SETCONTEXT = &H281
  WM_IME_STARTCOMPOSITION = &H10D
  WM_INITDIALOG = &H110
  WM_INITMENU = &H116
  WM_INITMENUPOPUP = &H117
  WM_KEYDOWN = &H100
  WM_KEYFIRST = &H100
  WM_KEYLAST = &H108
  WM_KEYUP = &H101
  WM_KILLFOCUS = &H8
  WM_LBUTTONDBLCLK = &H203
  WM_LBUTTONDOWN = &H201
  WM_LBUTTONUP = &H202
  WM_MBUTTONDBLCLK = &H209
  WM_MBUTTONDOWN = &H207
  WM_MBUTTONUP = &H208
  WM_MDIACTIVATE = &H222
  WM_MDICASCADE = &H227
  WM_MDICREATE = &H220
  WM_MDIDESTROY = &H221
  WM_MDIGETACTIVE = &H229
  WM_MDIICONARRANGE = &H228
  WM_MDIMAXIMIZE = &H225
  WM_MDINEXT = &H224
  WM_MDIREFRESHMENU = &H234
  WM_MDIRESTORE = &H223
  WM_MDISETMENU = &H230
  WM_MDITILE = &H226
  WM_MEASUREITEM = &H2C
  WM_MENUCHAR = &H120
  WM_MENUSELECT = &H11F
  WM_MOUSEACTIVATE = &H21
  WM_MOUSEFIRST = &H200
  WM_MOUSELAST = &H209
  WM_MOUSEMOVE = &H200
  WM_MOVE = &H3
  WM_NCACTIVATE = &H86
  WM_NCCALCSIZE = &H83
  WM_NCCREATE = &H81
  WM_NCDESTROY = &H82
  WM_NCHITTEST = &H84
  WM_NCLBUTTONDBLCLK = &HA3
  WM_NCLBUTTONDOWN = &HA1
  WM_NCLBUTTONUP = &HA2
  WM_NCMBUTTONDBLCLK = &HA9
  WM_NCMBUTTONDOWN = &HA7
  WM_NCMBUTTONUP = &HA8
  WM_NCMOUSEMOVE = &HA0
  WM_NCPAINT = &H85
  WM_NCRBUTTONDBLCLK = &HA6
  WM_NCRBUTTONDOWN = &HA4
  WM_NCRBUTTONUP = &HA5
  WM_NEXTDLGCTL = &H28
  'WM_OTHERWINDOWCREATED = &H42    '  no longer suported
  'WM_OTHERWINDOWDESTROYED = &H43  '  no longer suported
  WM_PAINT = &HF
  WM_PAINTCLIPBOARD = &H309
  WM_PAINTICON = &H26
  WM_PALETTECHANGED = &H311
  WM_PALETTEISCHANGING = &H310
  WM_PARENTNOTIFY = &H210
  WM_PASTE = &H302
  WM_PENWINFIRST = &H380
  WM_PENWINLAST = &H38F
  WM_POWER = &H48
  WM_PSD_ENVSTAMPRECT = (&H400 + 5)
  WM_PSD_FULLPAGERECT = (&H400 + 1)
  WM_PSD_GREEKTEXTRECT = (&H400 + 4)
  WM_PSD_MARGINRECT = (&H400 + 3)
  WM_PSD_MINMARGINRECT = (&H400 + 2)
  WM_PSD_PAGESETUPDLG = (&H400)
  WM_PSD_YAFULLPAGERECT = (&H400 + 6)
  WM_QUERYDRAGICON = &H37
  WM_QUERYENDSESSION = &H11
  WM_QUERYNEWPALETTE = &H30F
  WM_QUERYOPEN = &H13
  WM_QUEUESYNC = &H23
  WM_QUIT = &H12
  WM_RBUTTONDBLCLK = &H206
  WM_RBUTTONDOWN = &H204
  WM_RBUTTONUP = &H205
  WM_RENDERALLFORMATS = &H306
  WM_RENDERFORMAT = &H305
  WM_SETCURSOR = &H20
  WM_SETFOCUS = &H7
  WM_SETFONT = &H30
  WM_SETHOTKEY = &H32
  WM_SETREDRAW = &HB
  WM_SETTEXT = &HC
  WM_SHOWWINDOW = &H18
  WM_SIZE = &H5
  WM_SIZECLIPBOARD = &H30B
  WM_SPOOLERSTATUS = &H2A
  WM_SYSCHAR = &H106
  WM_SYSCOLORCHANGE = &H15
  WM_SYSCOMMAND = &H112
  WM_SYSDEADCHAR = &H107
  WM_SYSKEYDOWN = &H104
  WM_SYSKEYUP = &H105
  WM_TIMECHANGE = &H1E
  WM_TIMER = &H113
  WM_UNDO = &H304
  WM_USER = &H400
  WM_VKEYTOITEM = &H2E
  WM_VSCROLL = &H115
  WM_VSCROLLCLIPBOARD = &H30A
  WM_WINDOWPOSCHANGED = &H47
  WM_WINDOWPOSCHANGING = &H46
  WM_WININICHANGE = &H1A
End Enum
Public Property Let hWnd(NewhWnd As Long)
If m_hWnd = 0& Then
  If IsWindow(NewhWnd) Then
    m_hWnd = NewhWnd
  Else
    Err.Raise 10003, "Let hWnd", "Invalid hWnd"
  End If
Else
  Err.Raise 10004, "Let hWnd", "Can only set hWnd once"
End If
End Property

Public Sub RaiseAfterMsg(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Attribute RaiseAfterMsg.VB_MemberFlags = "40"
RaiseEvent AfterMessage(uMsg, wParam, lParam)
End Sub

Public Sub RaiseBeforeMsg(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Cancel As Boolean)
Attribute RaiseBeforeMsg.VB_MemberFlags = "40"
RaiseEvent BeforeMessage(uMsg, wParam, lParam, Cancel)
End Sub
Public Property Get Subclassed() As Boolean
Subclassed = m_Subclassed
End Property


Public Property Get SubclassRef() As Long
Attribute SubclassRef.VB_Description = "Reference to the old subclass procedure."
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubclassRef
    SubclassRef = m_SubclassRef
End Property



Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "The hWnd of the subclassed window."
Attribute hWnd.VB_UserMemId = 0
Attribute hWnd.VB_MemberFlags = "400"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hWnd
    hWnd = m_hWnd
End Property

Public Sub UnSubclass()
Attribute UnSubclass.VB_Description = "Ends subclassing."
Attribute UnSubclass.VB_MemberFlags = "40"
If Subclassed Then
  SetWindowLong m_hWnd, GWL_WNDPROC, m_SubclassRef
  m_Subclassed = False
Else
  Err.Raise 10002, "UnSubclass", "Cannot unsubclass an already unsubclassed window"
End If
End Sub


Public Sub Subclass()
Attribute Subclass.VB_Description = "Starts subclassing."
Attribute Subclass.VB_MemberFlags = "40"
If m_hWnd <> 0 Then
  If Not (Subclassed) Then
    m_SubclassRef = SetWindowLong(m_hWnd, GWL_WNDPROC, AddressOf WindowProc)
    m_Subclassed = True
  Else
    Err.Raise 10001, "Subclass", "Cannot subclass an already subclassed window"
  End If
Else
  Err.Raise 10000, "Subclass", "Invalid Window handle (must not be 0)"
End If
End Sub