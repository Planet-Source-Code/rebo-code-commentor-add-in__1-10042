Attribute VB_Name = "modMain"
Option Explicit

Private Declare Sub PostMessage Lib "user32" Alias "PostMessageA" (ByVal Hwnd&, ByVal msg&, ByVal wp&, ByVal lp&)
Private Declare Sub SetFocus Lib "user32" (ByVal Hwnd&)
Private Declare Function GetParent Lib "user32" (ByVal Hwnd&) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal Hwnd As Long) As Long

Const WM_SYSKEYDOWN = &H104
Const WM_SYSKEYUP = &H105
Const WM_SYSCHAR = &H106
Const VK_F = 70                         ' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'

Dim hwndMenu       As Long              'needed to pass the menu keystrokes to VB

Global gVBInstance  As VBIDE.VBE        'instance of VB IDE
Global gwinWindow   As VBIDE.Window     'used to make sure we only run one instance
Global gdocProcAndFunc As Object        'user doc object

Global Const APP_CATEGORY = "Microsoft Visual Basic AddIns"

Public Type CommentOptions
    useSorted As Boolean
    useOrganisation As Boolean
    useDateTime As Boolean
    useParameters As Boolean
    useStarOrMin As Boolean
    Initials As String
    UserName As String
    Organisation As String
    ComPrefix As String
End Type

Global ComOpt As CommentOptions 'Comment options type

Function InRunMode(VBInst As VBIDE.VBE) As Boolean
    InRunMode = (VBInst.CommandBars("File").Controls(1).Enabled = False)
End Function

Sub HandleKeyDown(ud As Object, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrExit

    If Shift <> 4 Then Exit Sub
    If KeyCode < 65 Or KeyCode > 90 Then Exit Sub
    If gVBInstance.DisplayModel = vbext_dm_SDI Then Exit Sub

    If hwndMenu = 0 Then hwndMenu = FindHwndMenu(ud.Hwnd)
    PostMessage hwndMenu, WM_SYSKEYDOWN, KeyCode, &H20000000
    KeyCode = 0
    SetFocus hwndMenu
    Exit Sub

ErrExit:

End Sub

Function FindHwndMenu&(ByVal Hwnd&)
    Dim h As Long
  
Loop2:
    h = GetParent(Hwnd)
    If h = 0 Then FindHwndMenu = Hwnd: Exit Function
    Hwnd = h
    GoTo Loop2
End Function
