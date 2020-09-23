VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10245
   ClientLeft      =   8730
   ClientTop       =   4425
   ClientWidth     =   14580
   _ExtentX        =   25718
   _ExtentY        =   18071
   _Version        =   393216
   Description     =   $"RBVBCommentor.dsx":0000
   DisplayName     =   "VB 6 Procedure & Fuction Commentor"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   SatName         =   "RBVBCommentor.dll"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const guidMYTOOL$ = "_P_r_o_c__A_n_d__F_u_n_c_"

Public WithEvents PrjHandler  As VBProjectsEvents          'projects event handler
Attribute PrjHandler.VB_VarHelpID = -1
Public WithEvents CmpHandler  As VBComponentsEvents        'components event handler
Attribute CmpHandler.VB_VarHelpID = -1
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Dim mcbMenuCommandBar         As Office.CommandBarControl  'command bar object


Sub Show()
    On Error GoTo ShowErr
  
    gwinWindow.Visible = True
 
    Exit Sub
    
ShowErr:
    MsgBox "Connect.Show: " & Err.Description
End Sub

Public Property Get NonModalApp() As Boolean
    NonModalApp = True  'used by addin toolbar
End Property

'------------------------------------------------------
'this method adds the Add-In to the VB Tools menu
'it is called by the VB addin manager
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
On Error GoTo AddinInstance_OnConnectionErr
  
    Dim aiTmp As AddIn
  
    'save the vb instance
    Set gVBInstance = Application

    If Not gwinWindow Is Nothing Then
        'already running so just show it
        Show
        If ConnectMode = ext_cm_AfterStartup Then
            'started from the addin manager
            AddToCommandBar
        End If
        Exit Sub
    End If
  
    'create the tool window
    If ConnectMode = ext_cm_External Then
    
        'need to see if it is already running
        On Error Resume Next
        Set aiTmp = gVBInstance.Addins("ProcAndFunc.Connect")
        On Error GoTo AddinInstance_OnConnectionErr
    
        If aiTmp Is Nothing Then
            'app is not in the VBADDIN.INI file so it is not in the collection
            'so lets attempt to use the 1st addin in the collection just
            'to get this app running and if there are none, an error
            'will occur and this app will not run
            Set gwinWindow = gVBInstance.Windows.CreateToolWindow(gVBInstance.Addins(1), "ProcAndFunc.docProcAndFunc", "Procedure & Function Commentor Window", guidMYTOOL$, gdocProcAndFunc)
        Else
            If aiTmp.Connect = False Then
                Set gwinWindow = gVBInstance.Windows.CreateToolWindow(aiTmp, "ProcAndFunc.docProcAndFunc", "Procedure & Function Commentor Window", guidMYTOOL$, gdocProcAndFunc)
            End If
        End If
    Else
        'must've been called from addin mgr
        Set gwinWindow = gVBInstance.Windows.CreateToolWindow(AddInInst, "ProcAndFunc.docProcAndFunc", "Procedure & Function Commentor Window", guidMYTOOL$, gdocProcAndFunc)
    End If

    'sink the project, components and controls event handler
    Set Me.PrjHandler = gVBInstance.Events.VBProjectsEvents
    Set Me.CmpHandler = gVBInstance.Events.VBComponentsEvents(Nothing)
    
'    Set Me.CtlHandler = gVBInstance.Events.VBControlsEvents(Nothing, Nothing)
  
    If ConnectMode = vbext_cm_External Then
        'started from the addin toolbar
        Show
    ElseIf ConnectMode = vbext_cm_AfterStartup Then
        'started from the addin manager
        AddToCommandBar
    End If

    Exit Sub
  
AddinInstance_OnConnectionErr:
    MsgBox "Connect.AddinInstance_OnConnection: " & Err.Description
End Sub

'------------------------------------------------------
'this event removes the commandbar menu
'it is called by the VB addin manager
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
On Error GoTo IDTExtensibility_OnDisconnectionErr
  
    'delete the command bar entry
    mcbMenuCommandBar.Delete
  
    'save the form state for next time VB is loaded
    If gwinWindow.Visible Then
        SaveSetting APP_CATEGORY, App.Title, "DisplayOnConnect", "1"
    Else
        SaveSetting APP_CATEGORY, App.Title, "DisplayOnConnect", "0"
    End If
  
    Set gwinWindow = Nothing
  
IDTExtensibility_OnDisconnectionErr:
  
End Sub

'this event fires when the IDE is fully loaded
Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    AddToCommandBar
End Sub

'this event fires when the command bar control is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Show
End Sub

'this event fires when a form becomes activated in the IDE
Private Sub CmpHandler_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
On Error GoTo CmpHandler_ItemActivatedErr

    If gwinWindow.Visible Then
        gdocProcAndFunc.RefreshList
    End If
    
CmpHandler_ItemActivatedErr:

End Sub

'this event fires when a form is selected in the project window
Private Sub CmpHandler_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    CmpHandler_ItemActivated VBComponent
End Sub

'this event fires when a form is removed in the project window
Private Sub CmpHandler_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
    CmpHandler_ItemActivated VBComponent
End Sub

'this event fires when a form is reloaded in the project window
Private Sub CmpHandler_ItemReloaded(ByVal VBComponent As VBIDE.VBComponent)
    CmpHandler_ItemActivated VBComponent
End Sub

Sub AddToCommandBar()
    On Error GoTo AddToCommandBarErr
  
    'make sure the standard toolbar is visible
    gVBInstance.CommandBars(2).Visible = True
  
    'add it to the command bar
    'the following line will add the ProcAndFunc manager to the
    'Standard toolbar to the right of the ToolBox button
    Set mcbMenuCommandBar = gVBInstance.CommandBars(2).Controls.Add(1, , , gVBInstance.CommandBars(2).Controls.Count)
  
    'set the caption
    mcbMenuCommandBar.Caption = "Proc & Func Commentor Window"
  
    'copy the icon to the clipboard
    Clipboard.SetData LoadResPicture(1000, 0)
  
    'set the icon for the button
    mcbMenuCommandBar.PasteFace
  
    'sink the event
    Set Me.MenuHandler = gVBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
  
    'restore the last state
    If GetSetting(APP_CATEGORY, App.Title, "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
  
    Exit Sub
    
AddToCommandBarErr:
    MsgBox "Connect.AddToCommandBarErr: " & Err.Description
End Sub

Private Sub PrjHandler_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    'this takes care of the user removing the only project
    If gwinWindow.Visible Then
        gdocProcAndFunc.RefreshList
    End If
End Sub
