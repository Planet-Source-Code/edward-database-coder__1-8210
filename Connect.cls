VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "qbd Database Coder"
' ==============================================================
' Module:       Connect
' Purpose:      Connection/disconnection with VB
' ==============================================================

Option Explicit

Implements IDTExtensibility

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmMain                 As New frmMain
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Public NonModalApp As Boolean                              'used by AddIn Toolbar



Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    Unload mfrmMain
   
End Sub

Sub Show()
  
    On Error Resume Next
    
    If mfrmMain Is Nothing Then
        Set mfrmMain = New frmMain
    End If
    
    Set mfrmMain.VBInstance = VBInstance
    Set mfrmMain.Connect = Me
    FormDisplayed = True
    mfrmMain.Show vbModal
   
End Sub

Private Sub Class_Initialize()
  NonModalApp = False   'used by addin toolbar

End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = VBInst
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    Debug.Print VBInst.FullName

    If ConnectMode = vbext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("VB Database Coder")
        'sink the event
        Set Me.MenuHandler = VBInst.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
 
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
    
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    
    Unload mfrmMain
    Set mfrmMain = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
            
    End If
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
'
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
        
    ' Add icon to add-in toolbar
    
    
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function




