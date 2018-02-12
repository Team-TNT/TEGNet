Attribute VB_Name = "mdlSysTray"
Option Explicit

'//////////////////////////////////////////////////////////////////
'//Project:         Sample Code Library
'//Date:            May 2nd, 1999  8:30AM EST
'//Programmer:      Robert J. Reich  (cypher@tir.com)
'//Company:         CypherSolutions
'//
'//Name:            System Tray Application
'//Description:     Demonstrates how to create an application that
'//                 resides in the systray instead of on the taskbar.
'//                 This is done with the WIN32 API.
'//
'//Note:            Usually the two APIs, the constants, and UDT will
'//                 all be declared in a stadard module and not in a
'//                 form.  This is only done here to make the example
'//                 as simple as possible.
'//////////////////////////////////////////////////////////////////
'//
'//WARNING:         If you run this in the IDE, do NOT use the VCR-STOP
'//                 button to end this program.  That will bypass the
'//                 Form_Unload event which is neccisary to restore
'//                 the system tray back to it's original state.
'//////////////////////////////////////////////////////////////////


'//These are the two API functions we'll need to use here.  The first
'//is the one that really does the work.  The second function is used
'//to take action when the user clicks on the mouse icon (restores the
'//program and brings it to front of all other windows.)
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
          (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetForegroundWindow Lib "user32" _
          (ByVal hwnd As Long) As Long


'//UDT required by Shell_NotifyIcon API call
Private Type NOTIFYICONDATA
 cbSize As Long             '//size of this UDT
 hwnd As Long               '//handle of the app
 uId As Long                '//unused (set to vbNull)
 uFlags As Long             '//Flags needed for actions
 uCallBackMessage As Long   '//WM we are going to subclass
 hIcon As Long              '//Icon we're going to use for the systray
 szTip As String * 64       '//ToolTip for the mouse_over of the icon.
End Type


'//Constants required by Shell_NotifyIcon API call:
Private Const NIM_ADD = &H0             '//Flag : "ALL NEW nid"
Private Const NIM_MODIFY = &H1          '//Flag : "ONLY MODIFYING nid"
Private Const NIM_DELETE = &H2          '//Flag : "DELETE THE CURRENT nid"
Private Const NIF_MESSAGE = &H1         '//Flag : "Message in nid is valid"
Private Const NIF_ICON = &H2            '//Flag : "Icon in nid is valid"
Private Const NIF_TIP = &H4             '//Flag : "Tip in nid is valid"
Private Const WM_MOUSEMOVE = &H200      '//This is our CallBack Message
Private Const WM_LBUTTONDOWN = &H201    '//LButton down
Private Const WM_LBUTTONUP = &H202      '//LButton up
Private Const WM_LBUTTONDBLCLK = &H203  '//LDouble-click
Private Const WM_RBUTTONDOWN = &H204    '//RButton down
Private Const WM_RBUTTONUP = &H205      '//RButton up
Private Const WM_RBUTTONDBLCLK = &H206  '//RDouble-click

Private nid As NOTIFYICONDATA       '//global UDT for the systray function


Public Sub SysTrayInicializar(idForm As Long, strTip As String, imagen As Image)
'//////////////////////////////////////////////////////////////////
'//Purpose:         Load up the UDT for the Systray Function.  This
'//                 must be done after the form is fully visable.
'//                 The Form_Activate is a perfect place for that.
'//////////////////////////////////////////////////////////////////
 
  With nid
    .cbSize = Len(nid)
    .hwnd = idForm
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = imagen.Picture
    .szTip = strTip & vbNullChar
  End With
 
  Shell_NotifyIcon NIM_ADD, nid
End Sub

Public Sub SysTrayMouseMove(frmform As Form, Button As Integer, Shift As Integer, _
                                X As Single, Y As Single)
'//////////////////////////////////////////////////////////////////
'//Purpose:         This is the callback function of icon in the
'//                 system tray.  This is where will will process
'//                 what the application will do when Mouse Input
'//                 is given to the icon.
'//
'//Inputs:          What Button was clicked (this is button & shift),
'//                 also, the X & Y coordinates of the mouse.
'//////////////////////////////////////////////////////////////////

  Dim msg As Long     '//The callback value
  
  '//The value of X will vary depending
  '//upon the ScaleMode setting.  Here
  '//we are using that fact to determine
  '//what the value of 'msg' should really be
  If (frmform.ScaleMode = vbPixels) Then
    msg = X
  Else
    msg = X / Screen.TwipsPerPixelX
  End If

  Select Case msg
    Case WM_LBUTTONDBLCLK    '515 restore form window
      SystrayRestaurar frmform
    Case WM_RBUTTONUP        '517 display popup menu
      Call SetForegroundWindow(frmform.hwnd)
      frmform.PopupMenu frmform.mnuSysTray
    
    Case WM_LBUTTONUP        '514 restore form window
      '//commonly an application on the
      '//systray will do nothing on a
      '//single mouse_click, so nothing
  End Select

  '//small note:  I just learned that when using a Select Case
  '//structure you always want to place the most commonly anticipated
  '//action highest. Saves CPU cycles becuase of less evaluations.
End Sub

Public Sub SysTrayResize(frmform As Form)
'//////////////////////////////////////////////////////////////////
'//Purpose:         This is just to check to make sure that, if
'//                 indeed the application is minimized (hence on
'//                 the systray) to also hide the form.
'//////////////////////////////////////////////////////////////////
  If (frmform.WindowState = vbMinimized) Then frmform.Hide
End Sub

Public Sub SystrayRestaurar(frmform As Form)
'//////////////////////////////////////////////////////////////////
'//Purpose:         When the application is minimized on the systray
'//                 this will restore it.
'//////////////////////////////////////////////////////////////////
  frmform.WindowState = vbNormal
  Call SetForegroundWindow(frmform.hwnd)
  frmform.Show
  
End Sub

Public Sub SysTrayExit(frmform As Form)
'//////////////////////////////////////////////////////////////////
'//Purpose:         When the application is minimized on the systray
'//                 this will close the application.
'//////////////////////////////////////////////////////////////////
  Unload frmform
End Sub

Public Sub SysTrayUnload(frmform As Form)
'//////////////////////////////////////////////////////////////////
'//Purpose:         Deletes the systray icon, and makes the application
'//                 "safe" to unload.
'//////////////////////////////////////////////////////////////////
   Shell_NotifyIcon NIM_DELETE, nid
   Set frmform = Nothing
End Sub

Public Sub SysTrayChangeTip(idForm As Long, strNuevoTip As String)
'//////////////////////////////////////////////////////////////////
'//Purpose:         Change the ToolTip of the System Tray Icon
'//////////////////////////////////////////////////////////////////
  Dim nidNewTip As NOTIFYICONDATA     '//New ToolTip nid
    
  With nidNewTip
    .cbSize = Len(nidNewTip)
    .hwnd = idForm
    .uId = vbNull
    .uFlags = NIF_TIP       '//Here the Tip is the only valid "new data"
    .szTip = strNuevoTip & vbNullChar
  End With
    
  Shell_NotifyIcon NIM_MODIFY, nidNewTip
End Sub

Public Sub SysTrayChangeIcon(idForm As Long, imagen As Image)
'//////////////////////////////////////////////////////////////////
'//Purpose:         Load up the UDT for the Systray Function.  This
'//                 must be done after the form is fully visable.
'//                 The Form_Activate is a perfect place for that.
'//////////////////////////////////////////////////////////////////
 Dim nidNewTip As NOTIFYICONDATA     '//New Image nid
  
  With nidNewTip
    .cbSize = Len(nidNewTip)
    .hwnd = idForm
    .uId = vbNull
    .uFlags = NIF_ICON       '//Here the Tip is the only valid "new data"
    .hIcon = imagen.Picture
  End With
 
  Shell_NotifyIcon NIM_MODIFY, nidNewTip
End Sub


