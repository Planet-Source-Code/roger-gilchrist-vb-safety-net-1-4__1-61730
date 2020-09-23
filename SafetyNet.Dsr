VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   6795
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   13425
   _ExtentX        =   23680
   _ExtentY        =   11986
   _Version        =   393216
   Description     =   $"SafetyNet.dsx":0000
   DisplayName     =   "VB Safety Net"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private FormDisplayed               As Boolean
'Private mcbMenuCommandBar          As Office.CommandBarButton
Private mcbMenuCommandBar           As Office.CommandBarControl
Private mcbMenuCommandBar1          As Office.CommandBarControl
Private mcbMenuCommandBar2          As Office.CommandBarControl
Public WithEvents MenuHandler       As CommandBarEvents    'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents MenuHandler1      As CommandBarEvents    'command bar event handler
Attribute MenuHandler1.VB_VarHelpID = -1
Public WithEvents MenuHandler2      As CommandBarEvents    'command bar event handler
Attribute MenuHandler2.VB_VarHelpID = -1
Public WithEvents FileLoad          As FileControlEvents
Attribute FileLoad.VB_VarHelpID = -1


'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, _
                                       ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, _
                                       ByVal AddInInst As Object, _
                                       custom() As Variant)

  On Error GoTo error_handler
  Set VBInstance = Application
  DangerFound = False ' needed to rest on new loadings
  ArrFuncPropSub = Array("Function", "Sub", "Property")
  arrMaliciousStart = Array("Print", "Write", "Open", "Shell", "Kill", "Enabled", "\Windows\", "\System32\", ".shell", ".regwrite", _
                            ".regdelete", ".regread", "WinExec", "TerminateProcess", "ExitProcess", "RemoveDirectory", "DeleteFile", _
                            "SetVolumeLabel", "ExitWindows", "ExitWindowsEx", "SetComputerName", "DeletePrinter", "DeletePort", _
                            "DeletePrintDrive", "DeleteMoniter", "DeletePrintProcessor", "DeletePrintProvider", "RegDeleteKey", _
                            "RegDeleteKeyEx", "RegDeleteValue", "RegReplaceKey", "RegSetValue", "RegSetValueEx", "RegCreateKey", _
                            "RegCreateKeyEx", "MoveFile", "MoveFileEx", "OpenFile", "WriteFile", "LoadKeyBoardLayout", "SetLocaleInfo", _
                            "SetLocalTime", "SetSysTime", "SetSystemTimeAdjustInformation", "SetSysColors", "SetEnvironmentVariable", _
                            "SystemParametersInfo", "SystemParametersInfoByval", "UnloadKeyboardLayout", "RtlAdjustPrivilege", "NtShutdownSystem", _
                            "RegisterServiceProcess", "Shell32.Shell", "ShellExecute")

  StrGuard = "'SAFETY NET Ignore Threat OK by " & GetRegisteredDetails("RegisteredOwner")
'activate the FileControlEvents
  Set FileLoad = VBInstance.Events.FileControlEvents(Nothing)
  strMode = GetSetting(App.EXEName, "Options", "Mode", "Hard")
  Set mcbMenuCommandBar = AddToAddInCommandBar("VB Safety Net", 1)
  Set mcbMenuCommandBar1 = AddToAddInPopup(mcbMenuCommandBar, "Ignore Threat")
  Set MenuHandler1 = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar1)
  Set mcbMenuCommandBar2 = AddToAddInPopup(mcbMenuCommandBar, "Mode Hard")
  Set MenuHandler2 = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar2)

Exit Sub

error_handler:
  MsgBox Err.Description, , "SAFETY NET"

End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, _
                                          custom() As Variant)

  On Error Resume Next
'delete the command bar entry
  mcbMenuCommandBar.Delete
'shut down the Add-In
  If FormDisplayed Then
    SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
    FormDisplayed = False
   Else
    SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
  End If
  On Error GoTo 0

End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)

'if any dangerous code was found jump the IDE to it

  If DangerFound Then
    JumpTo WARNING_MSG & "DANGEROUS CODE DISABLED FOR SAFETY PURPOSES"
  End If

End Sub

Private Function AddToAddInCommandBar(ByVal sCaption As String, _
                                      ByVal mode As Long) As Office.CommandBarControl

  Dim cbMenuCommandBar As Office.CommandBarControl   'command bar object
  Dim cbMenu           As Object

  On Error GoTo AddToAddInCommandBarErr
'see if we can find the Add-Ins menu
  Set cbMenu = VBInstance.CommandBars("Add-Ins")
  If Not cbMenu Is Nothing Then
'add it to the command bar
    Select Case mode
     Case 0
      Set cbMenuCommandBar = cbMenu.Controls.Add(1)
     Case 1
      Set cbMenuCommandBar = cbMenu.Controls.Add(Type:=msoControlPopup)
    End Select
'set the caption
    cbMenuCommandBar.Caption = sCaption
    Set AddToAddInCommandBar = cbMenuCommandBar
  End If
AddToAddInCommandBarErr:

End Function

Private Function AddToAddInPopup(RootMenu As CommandBarControl, _
                                 ByVal sCaption As String) As Office.CommandBarControl

  Dim cbMenuCommandBar As Office.CommandBarControl   'command bar object

  On Error GoTo AddToAddInCommandBarErr
  If Not RootMenu Is Nothing Then
'add it to the command bar
    Set cbMenuCommandBar = RootMenu.Controls.Add(1)
'set the caption
    cbMenuCommandBar.Caption = sCaption
    Set AddToAddInPopup = cbMenuCommandBar
  End If
AddToAddInCommandBarErr:

End Function

Private Sub FileLoad_BeforeLoadFile(ByVal VBProject As VBIDE.VBProject, _
                                    FileNames() As String)

'get to files before they are loaded
'note you cannot stop the file loading but you can manipulate it.

  VBPHandler VBProject, FileNames()

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)

'<STUB> Reason:'


End Sub

Private Sub MenuHandler1_Click(ByVal CommandBarControl As Object, _
                               handled As Boolean, _
                               CancelDefault As Boolean)

  Dim sline As Long

'stamp a line into the code
  With VBInstance.ActiveCodePane
    If .Window.Visible Then
      .GetSelection sline, 0, 0, 0
      .CodeModule.InsertLines sline, StrGuard & " " & Now
    End If
  End With

End Sub

Private Sub MenuHandler2_Click(ByVal CommandBarControl As Object, _
                               handled As Boolean, _
                               CancelDefault As Boolean)

  Select Case CommandBarControl.Caption
   Case "Mode Hard"
    strMode = "Soft"
   Case "Mode Soft"
    strMode = "Hard"
  End Select
  CommandBarControl.Caption = "Mode " & strMode
  SaveSetting App.EXEName, "Options", "Mode", strMode

End Sub


':)Code Fixer V4.0.0 (Wednesday, 17 August 2005 07:53:35) 10 + 144 = 154 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|33332222222222222222222222222222|1112222|2221222|222222222233|111111111111|1222222222220|333333|

