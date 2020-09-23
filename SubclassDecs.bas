Attribute VB_Name = "modSubclassDecs"
Option Explicit

' Subclassing declarations
' Copies memory locations
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' For WM_GETMINMAXINFO message
Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type MINMAXINFO
  ptReserved As POINTAPI  ' Reserved (not used)
  ptMaxSize As POINTAPI  ' Max Width/Heigth of window
  ptMaxPosition As POINTAPI  ' Left/Top of Maximized window
  ptMinTrackSize As POINTAPI  ' Minimum sizable Width/Height of window
  ptMaxTrackSize As POINTAPI  ' Maximum sizable Width/Height of window
End Type

' For WM_DRAWITEM message
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type DRAWITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemAction As Long
  itemState As Long
  hWndItem As Long
  hDC As Long
  rcItem As RECT
  itemData As Long
End Type

' for WM_MEASUREITEM message
Public Type MEASUREITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemWidth As Long
  itemHeight As Long
  itemData As Long
End Type

' DrawItemStruct.itemAction
Public Const ODA_DRAWENTIRE = &H1, ODA_FOCUS = &H4, ODA_SELECT = &H2

' DrawItemStruct.itemState
Public Const ODS_CHECKED = &H8, ODS_DISABLED = &H4, ODS_FOCUS = &H10
Public Const ODS_GRAYED = &H2, ODS_SELECTED = &H1

' DrawItemStruct.ctlType
Public Const ODT_BUTTON = 4, ODT_COMBOBOX = 3
Public Const ODT_LISTBOX = 2, ODT_MENU = 1

' Controls to set as owner-draw

' For setting styles (to owner-draw)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_STYLE = (-16)


Public Const BS_OWNERDRAW = &HB&  ' Command button

' Combo box
Public Const CBS_OWNERDRAWFIXED = &H10&  ' Fixed height items
Public Const CBS_OWNERDRAWVARIABLE = &H20&  ' Variable height items

Public Const LBS_OWNERDRAWFIXED = &H10&  ' Fixed height items
Public Const LBS_OWNERDRAWVARIABLE = &H20&  ' Variable height items

' Menu items
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long

Public Const MF_OWNERDRAW = &H100&

' Other declarations used for subclassing

' Used for owner-draw command button
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

' DrawText constants
Public Const DT_BOTTOM = &H8, DT_CENTER = &H1, DT_LEFT = &H0
Public Const DT_RIGHT = &H2, DT_SINGLELINE = &H20, DT_VCENTER = &H4

Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long

Public Const MF_STRING = &H0&
Public Const MF_SEPARATOR = &H800&

' For WM_MENUSELECT message
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
