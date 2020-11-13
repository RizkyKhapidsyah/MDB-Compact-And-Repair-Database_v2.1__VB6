VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMiscForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------
' Adapted from an example in VBPJ Tech Tips, 10th Edition.
'
' Declares needed for centering a form on the screen
' ---------------------------------------------------------------------------
  Private Const SM_CXFULLSCREEN As Long = 16
  Private Const SM_CYFULLSCREEN As Long = 17

  Private Declare Function GetSystemMetrics Lib "user32" _
          (ByVal nIndex As Long) As Long

  Private Declare Function GetWindowLong Lib "user32" _
          Alias "GetWindowLongA" (ByVal hwnd As Long, _
          ByVal nIndex As Long) As Long

' ---------------------------------------------------------------------------
' Needed for the testing of a previous instance of this application
' ---------------------------------------------------------------------------
  Private Const SW_SHOWMINIMIZED    As Long = 2
  Private Const SW_SHOWNORMAL       As Long = 1
  Private Const SW_SHOWNOACTIVATE   As Long = 4
  Private Const SW_RESTORE          As Long = 9

  Private Type RECT
       Left    As Long
       Top     As Long
       Right   As Long
       Bottom  As Long
  End Type

  Private Type POINTAPI
       x       As Long
       Y       As Long
  End Type

  Private Type WINDOWPLACEMENT
       Length            As Long
       flags             As Long
       showCmd           As Long
       ptMinPosition     As POINTAPI
       ptMaxPosition     As POINTAPI
       rcNormalPosition  As RECT
  End Type

' ---------------------------------------------------------------------------
' Declares required for verifying a previous instance of program executiion
' ---------------------------------------------------------------------------
  Private Declare Function FindWindow Lib "user32" _
          Alias "FindWindowA" (ByVal lpClassName As String, _
          ByVal lpWindowName As String) As Long
          
  Private Declare Function SetWindowPlacement Lib "user32" _
          (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
  
  Private Declare Function SetForegroundWindow Lib "user32" _
          (ByVal hwnd As Long) As Long
  
  Private Declare Function GetWindowPlacement Lib "user32" _
          (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
  
  Private Declare Function SendMessageArray Lib "user32" _
          Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
          ByVal wParam As Long, lParam As Any) As Long

' ---------------------------------------------------------------------------
' Remove the "X" from the window and menu
'
' A VB developer may find themselves developing an application who's
' integrity is crucial, and therefore must prevent the user from
' accidentally terminating the application during its life, while still
' displaying the system menu.  And while Visual Basic does provide two
' places to cancel an impending close - the QueryUnload and Unload subs -
' such a sensitive application may need to totally prevent even activation of
' the shutdown.
'
' Although it is not possible to simply disable the Close button while the
' Close system menu option is present, just a few lines of API code will
' remove the system menu Close option and in doing so permanently disable
' the titlebar close button.
' ---------------------------------------------------------------------------
  Private Const MF_BYPOSITION   As Long = &H400
  Private Const MF_REMOVE       As Long = &H1000

  Private Declare Function DrawMenuBar Lib "user32" _
          (ByVal hwnd As Long) As Long
  
  Private Declare Function GetMenuItemCount Lib "user32" _
          (ByVal hMenu As Long) As Long
  
  Private Declare Function GetSystemMenu Lib "user32" _
          (ByVal hwnd As Long, ByVal bRevert As Long) As Long
  
  Private Declare Function RemoveMenu Lib "user32" _
          (ByVal hMenu As Long, ByVal nPosition As Long, _
          ByVal wFlags As Long) As Long

' ---------------------------------------------------------------------------
' This is a rough translation of the GetTickCount API. The tick count of a
' PC is only valid for the first 49.7 days since it was last rebooted.  When
' you capture the tick count, you are capturing the total number of
' milliseconds since the PC was last rebooted.
' ---------------------------------------------------------------------------
  Private Declare Function GetTickCount Lib "kernel32" () As Long
  
Public Sub Delay(intAmtOfDelay As Integer, _
                 Optional strTypeOfDelay As String = "s")

' -----------------------------------------------------------
' This routine will cause a delay for the time requested,
' yet will not interfere with the program progress like the
' Sleep API.  This routine does not rely on the Timer event,
' a timer control, or the tickcount of the last reboot.  All
' of these have their drawbacks when dealing with time
' comparisons.
'
' Timer control is only valid from midnight of that particular
' day.  If you want to delay 5 minutes and the time is 23:59:00.
' You will never reach the finish time because at midnight,
' the timer control is reset.
'
' The timer event is based on a single precision caluclation
' of the date and time.  to the left of the decimal is the
' date and to the right is the time.  Somewhere in the 24 hour
' cycle, this is reset.  I suspect midnight.  this is good for
' immediate testing but not for comparisons.
'
' I use the system date and time by calling the VB function
' Now().  As long as the machine is running, it will have a
' system date and time stamp that is being updated.
'
' Parameters:
'       intAmtOfDelay - amount of time to delay
' -----------------------------------------------------------

' -----------------------------------------------------------
' Define local variables
' -----------------------------------------------------------
  Dim varDelayTime  As Variant
  Dim varCurrTime   As Variant
  
' -----------------------------------------------------------
' Determine the length of time to delay using the
' VB DateAdd function.
'
'    "s" - seconds      "d" - days
'    "n" - minutes      "m" - months
'    "h" - hours        "yyyy" - years
'
' We are adding the amount of delay to the current time
' -----------------------------------------------------------
  varDelayTime = Format(DateAdd(strTypeOfDelay, intAmtOfDelay, Now), "hh:mm:ss")
  
' -----------------------------------------------------------
' Loop thru and continualy check the curent time with the
' calculated time so we know when to leave
' -----------------------------------------------------------
  Do
      varCurrTime = Format(Now, "hh:mm:ss")
      DoEvents
      DoEvents
      
      ' if the string1 is greater than string2,
      ' a one will be returned
      If StrComp(varCurrTime, varDelayTime) = 1 Then
          Exit Do
      End If
  Loop

End Sub

Public Sub Mini_Delay(lngDelayAmt As Long)

' -----------------------------------------------------------
' This routine will cause a delay for the time requested,
' yet will not interfere with the program progress like the
' Sleep API.  I use this for other than whole seconds.
'
' Remenber 1000 milliseconds = 1 second.
'
' Parameters:
'       lngDelayAmt - amount of time to delay
'                   Tickcount + x number of milliseconds
'
' syntax:     Mini_Delay 750     ' 3/4 of a second
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngNewTime   As Long
  Dim lngCurrent   As Long
  
' ---------------------------------------------------------------------------
' Calculate the new waiting time
' ---------------------------------------------------------------------------
  lngNewTime = GetTickCount + lngDelayAmt
  
' ---------------------------------------------------------------------------
' Loop thru and continualy check the curent time with the
' calculated time so we know when to leave
' ---------------------------------------------------------------------------
  Do
      lngCurrent = GetTickCount      ' get the current millisecond count
      DoEvents
      
      ' if the current millisecond count has not
      ' caught up with the delay amount then
      ' we will try again.
      If lngCurrent >= lngNewTime Then
          Exit Do
      End If
  Loop

End Sub

Public Sub RemoveX(frm As Form)

' ---------------------------------------------------------------------------
' For completeness, you may want to confirm that the menuItemCount
' matches the value you expect before performing the removal. For
' example, on a normal form with a full system menu, menuItemCount will
' return seven.
'
' But what if your application is an MDI app and you want to disable the
' close button on the parent?  Just pass MDIForm1.hwnd as the form hwnd
' parameter in the GetSystemMenu() call.
' ---------------------------------------------------------------------------
   
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim hMenu           As Long
  Dim menuItemCount   As Long

' ---------------------------------------------------------------------------
' Obtain the handle to the form's system menu
' ---------------------------------------------------------------------------
  hMenu = GetSystemMenu(frm.hwnd, 0)
  
  If hMenu Then
      
     ' Obtain the number of items in the menu
      menuItemCount = GetMenuItemCount(hMenu)
    
     ' Remove the system menu Close menu item.
     ' The menu item is 0-based, so the last
     ' item on the menu is menuItemCount - 1
      Call RemoveMenu(hMenu, menuItemCount - 1, MF_REMOVE Or MF_BYPOSITION)
   
     ' Remove the system menu separator line
      Call RemoveMenu(hMenu, menuItemCount - 2, MF_REMOVE Or MF_BYPOSITION)
    
     ' Force a redraw of the menu. This
     ' refreshes the titlebar, dimming the X
      Call DrawMenuBar(frm.hwnd)

   End If
   
End Sub

Public Function IsAnotherInstance(strApplName As String) As Boolean

' ---------------------------------------------------------------------------
' Call this module from the Sub Main()
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variable
' ---------------------------------------------------------------------------
  Dim strSaveTitle As String
  
' ---------------------------------------------------------------------------
' Check for a previous instance of a program running
' ---------------------------------------------------------------------------
  If App.PrevInstance Then

      ' change the new instance title to prevent it
      ' from being located instead of the original
      ' instance.  Note however that as this is in
      ' a BAS module and not the form load sub,
      ' change "pgm_name" to the name of the application
      ' that you do not a dupliate instance of.
      strSaveTitle = App.Title
      App.Title = strApplName     ' name of executable w/o .exe extention
      RestorePreviousInstance strSaveTitle
      IsAnotherInstance = True
      Exit Function
  End If

  IsAnotherInstance = False
  
End Function

Private Sub RestorePreviousInstance(strPreviousTitle As String)

' ---------------------------------------------------------------------------
' Define local variable
' ---------------------------------------------------------------------------
  Dim lngRetCode   As Long
  Dim lngHandle    As Long
  Dim WinLocation  As WINDOWPLACEMENT
     
' ---------------------------------------------------------------------------
' VB3 & VB4 use class name "ThunderRTForm"
' VB5 uses class name "ThunderRT5Form"
' VB6 uses class name "ThunderRT6FormDC"
'
' Including the class name for the compiled EXE class prevents the routine
' from finding and attempting to activate the project form of the same name.
' ---------------------------------------------------------------------------
  lngHandle = FindWindow("ThunderRT6FormDC", strPreviousTitle)
  DoEvents
   
' ---------------------------------------------------------------------------
' If application is already executing
' ---------------------------------------------------------------------------
  If lngHandle > 0 Then
      ' get the current window state of the previous instance
      WinLocation.Length = Len(WinLocation)
      lngRetCode = GetWindowPlacement(lngHandle, WinLocation)
               
      ' if the WinLocation.showCmd member indicates that
      ' the window is currently minimized, it needs
      ' to be restored.
      If WinLocation.showCmd = SW_SHOWMINIMIZED Then
          With WinLocation
               .Length = Len(WinLocation)
               .flags = 0&
               .showCmd = SW_SHOWNORMAL
          End With
          lngRetCode = SetWindowPlacement(lngHandle, WinLocation)
      End If
       
      ' bring the window to the front and make
      ' the active window.  Without this, it
      ' may remain behind other windows.
      lngRetCode = SetForegroundWindow(lngHandle)
      DoEvents
  End If
   
End Sub

Public Function CenterForm(objForm As Form) As Boolean

' ---------------------------------------------------------------------------
' Adapted from an example in VBPJ Tech Tips, 10th Edition.
'
' Centers form on screen, accounting for taskbars such as the Windows
' taskbar, MS Office taskbar, etc.  Pass the form you want to center as
' the objForm parameter.  Returns TRUE if successful, FALSE if not.
'
' Scale mode of form must be vbTwips, otherwise this won't work.
'
' if you think that form's scale mode might be something else, incorporate
' the perpixelx and perpixely functions for metrics other than twips.
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngLeft As Long
  Dim lngTop As Long
        
  On Error GoTo Normal_Exit
' ---------------------------------------------------------------------------
' If the scale mode is not in TWIPS then leave
' ---------------------------------------------------------------------------
  If objForm.ScaleMode <> vbTwips Then
      CenterForm = False
      Exit Function
  End If

' ---------------------------------------------------------------------------
' Determine the full width of the screen.  Then calculate the width
' of the actual usable surface. Then calculate the centering of the form
' from left to right.
' ---------------------------------------------------------------------------
  lngLeft = (Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXFULLSCREEN)) / 2
  lngLeft = lngLeft - (objForm.Width / 2)
        
' ---------------------------------------------------------------------------
' Determine the full height of the screen.  Then calculate the height
' of the actual usable surface.  Then calculate the centering of the form
' from top to bottom.
' ---------------------------------------------------------------------------
  lngTop = (Screen.TwipsPerPixelY * GetSystemMetrics(SM_CYFULLSCREEN)) / 2
  lngTop = lngTop - (objForm.Height / 2)
        
' ---------------------------------------------------------------------------
' Now move the form into position
' ---------------------------------------------------------------------------
  objForm.Move lngLeft, lngTop
        
Normal_Exit:
' ---------------------------------------------------------------------------
' If not error occured then return a TRUE else FALSE
' ---------------------------------------------------------------------------
  CenterForm = (Err.Number And Err.LastDllError)

End Function


