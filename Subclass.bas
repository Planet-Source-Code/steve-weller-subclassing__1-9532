Attribute VB_Name = "modSubclass"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private WinCollection As New Collection

Public Function AddWindow(hWnd As Long) As SPWindow
' Add window object to private collection
Dim Win As SPWindow
Set Win = New SPWindow
Win.hWnd = hWnd

' CStr(hWnd) ensures that no duplicates exist
' (Collections don't allow duplicate keys)
WinCollection.Add Win, CStr(hWnd)
Set AddWindow = Win
End Function

Public Sub Destroy()
' If you are forced to end the program
' by the End button or resetting the project,
' type Destroy in the Immediate window to
' safely get back into VB
Call UnSubclass
Do Until WinCollection.Count = 0
  WinCollection.Remove 1
Loop
End Sub

Public Sub Subclass(Optional ByVal hWnd As Long = -1&)
Dim Win As SPWindow
If hWnd = -1& Then
  ' Subclass all Windows
  For Each Win In WinCollection
    If Not (Win.Subclassed) Then
      Win.Subclass
    End If
  Next Win
Else
  ' Subclass specified window
  Set Win = WinCollection(CStr(hWnd))
  If Not (Win.Subclassed) Then
    Win.Subclass
  End If
End If
End Sub

Public Sub UnSubclass(Optional ByVal hWnd As Long = -1&)
Dim Win As SPWindow
If hWnd = -1& Then
  ' Unsubclass all Windows
  For Each Win In WinCollection
    If Win.Subclassed Then
      Win.UnSubclass
    End If
  Next Win
Else
  ' Unsubclass specified window
  Set Win = WinCollection(CStr(hWnd))
  If Win.Subclassed Then
    Win.UnSubclass
  End If
End If
End Sub


Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Ignore errors (prevents crashing)
On Error Resume Next

' Get Window object for hWnd
Dim Win As SPWindow, Cancel As Boolean
Cancel = False
For Each Win In WinCollection
  If Win.hWnd = hWnd Then
    Win.RaiseBeforeMsg uMsg, wParam, lParam, Cancel
    
    If Cancel = False Then
      ' Send procedure to its
      ' old message handler
      WindowProc = CallWindowProc(Win.SubclassRef, hWnd, uMsg, wParam, lParam)
      
      Win.RaiseAfterMsg uMsg, wParam, lParam
    Else
      If uMsg = WM_GETMINMAXINFO Then
        ' This prevents an odd bug
        ' with this message
        WindowProc = 0
      Else
        ' Call default processing
        WindowProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
      End If
    End If
    Exit For
  End If
Next Win
End Function


