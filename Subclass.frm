VERSION 5.00
Begin VB.Form frmSubclass 
   Caption         =   "A Demonstration of Subclassing"
   ClientHeight    =   4095
   ClientLeft      =   1665
   ClientTop       =   1845
   ClientWidth     =   5910
   Icon            =   "Subclass.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5910
   Begin VB.CommandButton cmdSubclass 
      Caption         =   "Owner-Draw"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtSubclass 
      Height          =   3015
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Right click on me--No context menu!"
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3840
      Width           =   5895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "&Help Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Subclassing Demonstration
' by Steve Weller
' 7/5/2000

' This program demonstrates the benefits of subclassing including:
' ** Displaying text when a menu item is selected
' ** Suppressing a text box's right-click context menu (perhaps replacing
'    it with your own [Just use PopupMenu])
' ** Creating your own system menu commands
' ** Creating owner-draw buttons,list boxes, combo boxes, and menus
' Note: read the text file accompanying this project for warnings, etc.

Private WithEvents FormHook As SPWindow
Attribute FormHook.VB_VarHelpID = -1
Private WithEvents TextHook As SPWindow
Attribute TextHook.VB_VarHelpID = -1

Private Const MyMenuID& = 5000&

Private Sub cmdSubclass_Click()

End Sub


Private Sub Form_Load()
' Append menu to form's system menu
AppendMenu GetSystemMenu(Me.hWnd, False), MF_SEPARATOR, 0&, vbNullString
AppendMenu GetSystemMenu(Me.hWnd, False), MF_STRING, MyMenuID&, "About..."

' Set command button's style to owner-draw
' Note: WM_MEASUREITEM and WM_DRAWITEM sent
' to control's parent (the form in this case)
Dim Style&
Style& = GetWindowLong(cmdSubclass.hWnd, GWL_STYLE)
Style& = Style& Or BS_OWNERDRAW
SetWindowLong cmdSubclass.hWnd, GWL_STYLE, Style&

Set FormHook = AddWindow(Me.hWnd)
FormHook.Subclass
Set TextHook = AddWindow(txtSubclass.hWnd)
TextHook.Subclass
End Sub


Private Sub Form_Resize()
With txtSubclass
  .Width = Me.Width - 120
  If Me.Height > 2000 Then
    .Height = Me.Height - .Top - lblStatus.Height - 700
  End If
End With
With lblStatus
  .Width = Me.Width - 120
  .Top = Me.Height - .Height - 690
End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
FormHook.UnSubclass
TextHook.UnSubclass
End Sub


Private Sub FormHook_BeforeMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Cancel As Boolean)
Select Case uMsg
  Case WM_SYSCOMMAND  ' System command clicked
    If wParam = MyMenuID& Then
      ' Show about dialog
      Call mnuHelpAbout_Click
    End If
  Case WM_GETMINMAXINFO
    ' Copy structure for Min/Max Info
    Dim MinMax As MINMAXINFO
    CopyMemory MinMax, ByVal lParam, Len(MinMax)

    ' Set Info
    With MinMax
      .ptMaxSize.x = 600
      .ptMaxSize.y = 400
      .ptMaxTrackSize.x = 600
      .ptMaxTrackSize.y = 400
      .ptMinTrackSize.x = 200
      .ptMinTrackSize.y = 200
      ' Centers the form on the screen when maximized
      .ptMaxPosition.x = ((Screen.Width / Screen.TwipsPerPixelX) - 600) / 2
      .ptMaxPosition.y = ((Screen.Height / Screen.TwipsPerPixelY) - 400) / 2
    End With
    CopyMemory ByVal lParam, MinMax, Len(MinMax)
    ' Cancel message (required for WM_GETMINMAXINFO)
    Cancel = True
  Case WM_MENUSELECT
    ' Get Menu Caption
    Dim MenuCaption$, MenuCaptionLen&
    MenuCaption$ = Space$(256)
    
    ' lParam is hMenu
    ' MenuID is Low-order word of wParam
    MenuCaptionLen& = GetMenuString(lParam, wParam And &HFFFF&, MenuCaption$, Len(MenuCaption$), 0&)
    MenuCaption$ = Left$(MenuCaption$, MenuCaptionLen&)
    
    ' Extract Shortcut chars, if any (VB uses chr$(9)[Tab]
    ' to separate menu caption and shortcut keys)
    If InStr(MenuCaption$, Chr$(9)) > 0 Then
      MenuCaption$ = Left$(MenuCaption$, InStr(MenuCaption$, Chr$(9)) - 1)
    End If
    
    Select Case MenuCaption
      Case mnuFileNew.Caption: lblStatus = "Opens a new document"
      Case mnuFileOpen.Caption: lblStatus = "Opens an existing document"
      Case mnuFileSaveAs.Caption: lblStatus = "Saves document in a new file"
      Case mnuFileSave.Caption: lblStatus = "Saves document"
      Case mnuFileExit.Caption: lblStatus = "Exits the program"
      Case mnuHelpTopics.Caption: lblStatus = "Displays Help"
      Case mnuHelpAbout.Caption: lblStatus = "Shows about box"
      ' Case Else needed to clear status bar, etc.
      Case Else: lblStatus = ""
    End Select
  Case WM_MEASUREITEM
    ' Get dimensions of command button
    Dim MeasItem As MEASUREITEMSTRUCT
    CopyMemory MeasItem, ByVal lParam, Len(MeasItem)
    With MeasItem
      .itemHeight = ScaleY(cmdSubclass.Height, , vbPixels)
      .itemWidth = ScaleX(cmdSubclass.Width, , vbPixels)
    End With
    ' Copy back
    CopyMemory ByVal lParam, MeasItem, Len(MeasItem)
  Case WM_DRAWITEM
    ' Get draw item info
    Dim DrawCmd As DRAWITEMSTRUCT
    CopyMemory DrawCmd, ByVal lParam, Len(DrawCmd)
    With DrawCmd
      If .itemAction = ODA_FOCUS Or .itemAction = ODA_DRAWENTIRE Then
        Dim CmdBrush&  ' Brush for drawing on Command button
        If .itemState And ODS_FOCUS Then
          ' Has the focus
          CmdBrush& = CreateSolidBrush(vbRed)
          FillRect .hDC, .rcItem, CmdBrush&
          ' Draw the focus rectangle
          DrawFocusRect .hDC, .rcItem
        Else
          ' Lost the focus
          '* If someone can make this show as
          '* a blue command button (and not gray)
          '* I would appreciate knowing how
          '* Steve Weller
          CmdBrush& = CreateSolidBrush(vbBlue)
          FillRect .hDC, .rcItem, CmdBrush&
        End If
        ' Set background mode to Transparent
        SetBkMode .hDC, 1  ' Transparent
        
        ' Draw text on the command button
        DrawText .hDC, cmdSubclass.Caption, Len(cmdSubclass.Caption), .rcItem, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
        
        ' Set background mode to Opaque
        SetBkMode .hDC, 0  ' Opaque
        
        ' Destroy new brush
        DeleteObject CmdBrush&
      End If
    End With
End Select
End Sub


Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileNew_Click()

End Sub


Private Sub mnuHelpAbout_Click()
MsgBox "Demonstration of Subclassing" & vbCrLf & "Created by Steve Weller" & vbCrLf & "Copyright 2000", vbInformation, "About..."
End Sub


Private Sub mnuHelpTopics_Click()
' Show message box on this program
MsgBox "This program shows many uses of subclassing.  Click on the About... command on the system menu.  Right-click on the text box.  Select a menu item.  Click on and off of the command button (this is still a test run because of a bug).", vbInformation, "Subclassing"
End Sub


Private Sub TextHook_BeforeMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Cancel As Boolean)
If uMsg = WM_CONTEXTMENU Then
  ' Cancel default Text
  ' box context menu
  Cancel = True
End If
End Sub


Private Sub txtSubclass_Change()

End Sub


