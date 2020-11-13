Attribute VB_Name = "basDBCompact"
Option Explicit

' ---------------------------------------------------------------------------
' Compact MDB files
'
' Written by Kenneth Ives              kenaso@home.com
'
' This program will allow the user to select a MDB file
' to compact.  The size of the file is captured and a
' calculation of twice that size is made to determine
' the amount of free space required to compact the
' database.  Half that amount is used for a backup copy
' of the original database and the other half is for
' the compacted database.  if there is not enough space,
' the user is prompted to select another path in which
' to perform this operation or leave the application.
' After the database is compacted, the original is deleted
' and the new version is moved back into the place of the
' original.
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define global variables
' ---------------------------------------------------------------------------
  Public g_strDatabase    As String
  Public g_strWorkDrive   As String
  Public g_strResFolder   As String
  Public g_strResFile     As String
  Public g_clsDFI         As clsDFInfo
  
' ---------------------------------------------------------------------------
' Define module level variables
' ---------------------------------------------------------------------------
  Private m_strPartialTitle   As String
  Private m_bUsedTempDir      As Boolean
  Private m_lngRetCode        As Long
  Private m_strNewFile        As String
  Private m_strBakFile        As String
  Private m_strDBPath         As String
  Private m_strDBName         As String
  Private m_strTempDir        As String
  Private m_lngAppHandle      As Long
  Private m_lngStartSize      As Long
  Private m_lngEndSize        As Long

  Private Declare Function CopyFile Lib "kernel32" _
          Alias "CopyFileA" (ByVal lpExistingFileName As String, _
          ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

  Private Declare Function GetWindowText Lib "user32" _
          Alias "GetWindowTextA" _
          (ByVal hwnd As Long, ByVal lpString As String, _
          ByVal cch As Long) As Long

  Private Declare Function FindWindow Lib "user32" _
          Alias "FindWindowA" (ByVal lpClassName As String, _
          ByVal lpWindowName As String) As Long
          
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hObject As Long) As Long
  
  Private Declare Function EnumWindows Lib "user32" _
          (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

' ---------------------------------------------------------------------------
' Type layout for disk information
' ---------------------------------------------------------------------------
Public Function FindApplication(ByVal app_hWnd As Long, ByVal param As Long) As Long

' ---------------------------------------------------------------------------
' Check the title line of all active application windows
' while looking for a match on all or part of the title.
'
' Called from IsTaskActive
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lLength As Long    ' Length of the title string
  Dim sBuffer As String  ' buffer area to hold the title string
  Dim sTitle As String   ' application title string after formatting

' ---------------------------------------------------------------------------
' Initialize buffer string with blank values
' ---------------------------------------------------------------------------
  sBuffer = Space(256)
  
' ---------------------------------------------------------------------------
' Get the window's title.  (API call)
' ---------------------------------------------------------------------------
  lLength = GetWindowText(app_hWnd, sBuffer, Len(sBuffer))
  sTitle = StrConv(Left(sBuffer, lLength), vbLowerCase)

' ---------------------------------------------------------------------------
' See if this is the target window.
' ---------------------------------------------------------------------------
  If InStr(1, sTitle, m_strPartialTitle) > 0 Then
      
      ' capture the handle of the application window
      m_lngAppHandle = FindWindow(vbNullString, sTitle)
      Exit Function
  End If
    
' ---------------------------------------------------------------------------
' Continue searching the application windows
' ---------------------------------------------------------------------------
  FindApplication = 1

End Function

Public Function IsTaskActive(SApplName As String) As Long

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim m_lngRetCode As Long
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  m_strPartialTitle = StrConv(SApplName, vbLowerCase)
  m_lngAppHandle = 0
  
' ---------------------------------------------------------------------------
' Ask Windows for the list of tasks.
' ---------------------------------------------------------------------------
  m_lngRetCode = EnumWindows(AddressOf FindApplication, 0&)
  
' ---------------------------------------------------------------------------
' If successful then return the application handle
' ---------------------------------------------------------------------------
  IsTaskActive = m_lngAppHandle
  
End Function

Private Function Remove_Old_Databases() As Boolean

  On Error GoTo Remove_Old_Databases_Errors
  
' ---------------------------------------------------------------------------
' Get rid of any of the old databases, if they exist
' ---------------------------------------------------------------------------
  If g_clsDFI.File_Exist(m_strDBPath & m_strNewFile) Then
      Kill m_strDBPath & m_strNewFile
  End If
  
  If g_clsDFI.File_Exist(m_strDBPath & m_strBakFile) Then
      Kill m_strDBPath & m_strBakFile
  End If
  
  If g_clsDFI.File_Exist(g_strWorkDrive & m_strNewFile) Then
      Kill g_strWorkDrive & m_strNewFile
  End If
  
  If g_clsDFI.File_Exist(g_strWorkDrive & m_strBakFile) Then
      Kill g_strWorkDrive & m_strBakFile
  End If
  
  Remove_Old_Databases = True
  Exit Function

Remove_Old_Databases_Errors:

  MsgBox "An error occurred while attempting to remove some temporary " & _
         "databases.  These are their names:" & vbCrLf & vbCrLf & _
         vbTab & m_strDBPath & m_strNewFile & vbCrLf & _
         vbTab & m_strDBPath & m_strBakFile & vbCrLf & _
         vbTab & g_strWorkDrive & m_strNewFile & vbCrLf & _
         vbTab & g_strWorkDrive & m_strNewFile, vbOKOnly, "I/O Error"
  Remove_Old_Databases = False

End Function

Public Sub CompactMDB()

' ----------------------------------------------------------------------------
' CompactMDB routine is the heart of this module.
'
' Written by Kenneth Ives          kenaso@home.com
'
' This program will allow the user to select a MDB file to compact.  The size
' of the file is captured and a calculation of twice that size is made to
' determine the amount of free space required to compact the database.  Half
' that amount is used for a backup copy of the original database and the
' other half is for the compacted database.  if there is not enough space,
' the user is prompted to select another path in which to perform this
' operation or leave the application.  After the database is compacted, the
' original is deleted and the new version is moved back into the place of the
' original.
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strMsg            As String
  Dim strStartSize      As String
  Dim strEndSize        As String
  Dim strDiff           As String
  Dim strRatio          As String
  Dim strMsgBoxText     As String
  Dim strMsgBoxTitle    As String
  Dim n                 As Integer
  Dim intMsgBoxResp     As Integer
  Dim intResponse       As Integer
  Dim intAttributes     As Integer
  Dim lngDifference     As Long
  Dim sngRatio          As Single
  Dim dblBufferSize     As Double
  Dim dblTotalSpace     As Double
  Dim dblFreeSpace      As Double
  Dim dblUsedSpace      As Double

' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
StartOver:
  intAttributes = GetAttr(g_strDatabase)          ' Save the DB attributes
  n = InStrRev(g_strDatabase, "\")                ' parse backwards for first backslash
  m_strDBPath = Mid(g_strDatabase, 1, n)          ' save path to database
  m_strDBName = Mid(g_strDatabase, n + 1)         ' Save name of the database
  m_strDBName = StrConv(m_strDBName, vbUpperCase) ' convert name to uppercase
  m_lngStartSize = FileLen(g_strDatabase)           ' get database size before compacting
  dblBufferSize = CDbl(m_lngStartSize * 2)          ' Determine the buffer size for compacting
  
' ---------------------------------------------------------------------------
' Get the amount of free space on the drive where work directory is located.
' ---------------------------------------------------------------------------
  If Not g_clsDFI.Get_Disk_Space(g_strWorkDrive, dblTotalSpace, _
                                 dblFreeSpace, dblUsedSpace) Then
      GoTo Normal_Exit
  End If
  
' ---------------------------------------------------------------------------
' Display a message if not enough free space. If there is not at least two
' times the size of the database.
'     1.  Make a bakup copy of the database
'     2.  Creation of the new database
' if all is sucessful, these will be deleted.
' ---------------------------------------------------------------------------
  If dblBufferSize > dblFreeSpace Then
      '
      strMsgBoxTitle = "Need More Free Space"
      
      strMsgBoxText = "There is not enough free space on drive "
      strMsgBoxText = strMsgBoxText & g_strWorkDrive & ".    " & vbCrLf & vbCrLf
      strMsgBoxText = strMsgBoxText & "Need at least " & Format(dblBufferSize, "#,0") & " bytes of free "
      strMsgBoxText = strMsgBoxText & vbCrLf & "space to compact this database." & vbCrLf
      strMsgBoxText = strMsgBoxText & "    " & vbCrLf & "If you elect to free up more "
      strMsgBoxText = strMsgBoxText & "space, then this    " & vbCrLf & "application will "
      strMsgBoxText = strMsgBoxText & "terminate while you perform    " & vbCrLf
      strMsgBoxText = strMsgBoxText & "this action.    " & vbCrLf & "    " & vbCrLf
      strMsgBoxText = strMsgBoxText & "If you want to select another drive,    " & vbCrLf
      strMsgBoxText = strMsgBoxText & "click RETRY.    " & vbCrLf & "    "
      
      intMsgBoxResp = vbRetryCancel + vbQuestion + vbApplicationModal + vbDefaultButton1
      
      intResponse = MsgBox(strMsgBoxText, intMsgBoxResp, strMsgBoxTitle)
 
      Select Case intResponse
             Case vbRetry:                         ' Open browse for folder dialog box
                  Find_Work_Drive                  ' Select another drive/path
                  If Len(g_strWorkDrive) = 0 Then  ' if user did not select anything
                      GoTo Normal_Exit             '      then leave
                  Else
                      GoTo StartOver               ' Go back and test this new path
                  End If
                  
             Case vbCancel: GoTo Normal_Exit       ' time to leave
     End Select
  End If
  
' ---------------------------------------------------------------------------
' Make sure we can open this database in exclusive mode
' ---------------------------------------------------------------------------
  If Not Open_Exclusive Then
      GoTo Normal_Exit
  End If
  
' ---------------------------------------------------------------------------
' Initialize database variables
' ---------------------------------------------------------------------------
  m_strNewFile = UCase(Left(m_strDBName, Len(m_strDBName) - 3) & "NEW")
  m_strBakFile = UCase(Left(m_strDBName, Len(m_strDBName) - 3) & "BAK")
  DoEvents

' ---------------------------------------------------------------------------
' Get rid of any of the old databases, if they exist
' ---------------------------------------------------------------------------
  If Not Remove_Old_Databases Then
      GoTo Normal_Exit
  End If
  
' ---------------------------------------------------------------------------
' copy the existing database to the same name with a "BAK" extention and
' copy it to the temp directory.  Verify it got there and then delete it
' from this directory.
' ---------------------------------------------------------------------------
  DoEvents
  Screen.MousePointer = vbHourglass
  m_lngRetCode = CopyFile(g_strDatabase, g_strWorkDrive & m_strBakFile, False)
  Screen.MousePointer = vbNormal
  DoEvents
  
  If Not g_clsDFI.File_Exist(g_strWorkDrive & m_strBakFile) Then
      strMsgBoxText = "Failed to make a backup copy "
      strMsgBoxText = strMsgBoxText & "of the database." & vbLf & "Try again."
      MsgBox strMsgBoxText, vbOKOnly, "Bad Database Backup"
      GoTo Normal_Exit
  End If
  
On Error GoTo ErrorHandler
' ---------------------------------------------------------------------------
' Change to the database directory
' ---------------------------------------------------------------------------
  ChDrive m_strDBPath
  ChDir m_strDBPath
  
' ---------------------------------------------------------------------------
' Compact the database into the current name with an extention of "NEW"
' ---------------------------------------------------------------------------
  Screen.MousePointer = vbHourglass
  DoEvents      ' Now repair & compact the database
  DBEngine.CompactDatabase m_strDBName, g_strWorkDrive & m_strNewFile
  Screen.MousePointer = vbNormal
  DoEvents
 
' ---------------------------------------------------------------------------
' If the new database does not exist, then the compression was a failure.
' See if the user wants to attempt a repair.  If so, Do a repair on the
' database.  Go back to the top of this routine and start over.
' ---------------------------------------------------------------------------
  If Not g_clsDFI.File_Exist(g_strWorkDrive & m_strNewFile) Then
      strMsgBoxTitle = "Error Compacting Database"
      strMsgBoxText = "ERR: " & CStr(Err) & vbLf & Err.Description & vbLf & vbLf
      strMsgBoxText = strMsgBoxText & "==> " & g_strDatabase & vbLf & vbLf
      strMsgBoxText = strMsgBoxText & "Do you want to try this process again?  "
      intMsgBoxResp = vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton1
      
      intResponse = MsgBox(strMsgBoxText, intMsgBoxResp, strMsgBoxTitle)
 
      Select Case intResponse
             Case vbYes:
                  ' copy the original database back again and start over
                  If Replace_Original_DB(intAttributes) Then
                      ' remove all the temp databases
                      If Remove_Old_Databases Then
                          GoTo StartOver
                      Else
                          GoTo Normal_Exit
                      End If
                  Else
                      GoTo Normal_Exit
                  End If
                  
             Case vbNo:
                  ' restore everything back to square one
                  Replace_Original_DB intAttributes
                  GoTo Normal_Exit  ' Time to leave
      End Select
  End If
   
' ---------------------------------------------------------------------------
' Delete the original database because we successfully completed the
' previous steps.
' ---------------------------------------------------------------------------
  If g_clsDFI.File_Exist(g_strDatabase) Then
      Kill g_strDatabase
  End If

' ---------------------------------------------------------------------------
' move the new database to the original
' ---------------------------------------------------------------------------
  ' first copy the compacted database to the original location
  DoEvents
  Screen.MousePointer = vbHourglass
  m_lngRetCode = CopyFile(g_strWorkDrive & m_strNewFile, g_strDatabase, False)
  Screen.MousePointer = vbNormal
  DoEvents
  
  ' verify the copy was successful
  If Not g_clsDFI.File_Exist(g_strDatabase) Then
      strMsgBoxText = "Failed to replace database.  Contact support.   "
      MsgBox strMsgBoxText, vbOKOnly, "Bad Database Replace"
      GoTo Normal_Exit
  End If
  
' ---------------------------------------------------------------------------
' Get rid of any of the old databases
' ---------------------------------------------------------------------------
  DoEvents
  If Not Remove_Old_Databases Then
      GoTo Normal_Exit
  End If

' ---------------------------------------------------------------------------
' If we had to create the temp directory, then remove it.
' ---------------------------------------------------------------------------
  If m_bUsedTempDir Then
      Screen.MousePointer = vbHourglass
      If g_clsDFI.File_Exist(g_strWorkDrive) Then
          g_clsDFI.DelTree32 g_strWorkDrive
      End If
  End If
  Screen.MousePointer = vbNormal
  
' ---------------------------------------------------------------------------
' get the database data after compacting
' ---------------------------------------------------------------------------
  m_lngEndSize = FileLen(g_strDatabase)    ' get database size
  SetAttr g_strDatabase, intAttributes   ' Reset the DB attributes
  
' ----------------------------------------------------
' calculate the difference in size and compute the
' compression ratio
' ----------------------------------------------------
  lngDifference = m_lngStartSize - m_lngEndSize
  If lngDifference = 0 Then
      sngRatio = 0
  Else
      sngRatio = CSng(((m_lngStartSize - m_lngEndSize) / m_lngStartSize))
  End If
  
' ----------------------------------------------------
' Format the database results
' ----------------------------------------------------
  strStartSize = Format(Format(m_lngStartSize, "#,0"), "!@@@@@@@@@@@")
  strEndSize = Format(Format(m_lngEndSize, "#,0"), "!@@@@@@@@@@@")
  strDiff = Format(Format(lngDifference, "#,0"), "!@@@@@@@@@@@")
  strRatio = Format(Format(sngRatio, "Percent"), "!@@@@@@@@@@@@")
  
' ---------------------------------------------------
' Display a message showing how the database
' being compressed saved x amount of space.
' ---------------------------------------------------
  DoEvents
  strMsgBoxText = vbLf & g_strDatabase & vbTab & vbTab & vbLf & vbLf
  strMsgBoxText = strMsgBoxText & "Original File Size" & vbTab & vbTab & strStartSize & " Bytes" & Space(5) & vbLf
  strMsgBoxText = strMsgBoxText & "Compressed File Size" & vbTab & strEndSize & " Bytes" & Space(5) & vbLf
  strMsgBoxText = strMsgBoxText & "File Space Freed Up" & vbTab & strDiff & " Bytes" & Space(5) & vbLf
  strMsgBoxText = strMsgBoxText & "Compression Percentage" & vbTab & strRatio
  
  intMsgBoxResp = vbOKOnly + vbInformation + vbApplicationModal + vbDefaultButton1
      
  MsgBox strMsgBoxText, intMsgBoxResp, "Database Maintenance Info"
  
  
Normal_Exit:
  DoEvents
  Screen.MousePointer = vbNormal
  StopTheProgram
  Exit Sub


' =========================================================
'      E R R O R   P R O C E S S I N G   S E C T I O N
' =========================================================
ErrorHandler:
' ---------------------------------------------------------------------------
' Display a message the operation was a failure
' ---------------------------------------------------------------------------
  DoEvents
  Screen.MousePointer = vbNormal
  
  strMsgBoxText = "ERR: " & CStr(Err) & vbLf & Err.Description & vbLf & vbLf
  strMsgBoxText = strMsgBoxText & "DB:  " & g_strDatabase & vbLf & vbLf & "Contact support personnel."
  intMsgBoxResp = vbOKOnly + vbCritical + vbApplicationModal + vbDefaultButton1
  MsgBox strMsgBoxText, intMsgBoxResp, "Error Compacting Database"
  Err.Clear
  Resume Normal_Exit
  
End Sub

Private Function Replace_Original_DB(intAttributes As Integer) As Boolean

  On Error GoTo Replace_Original_DB_Errors
' ---------------------------------------------------------------------------
' Change the cursor to an hourglass, and delete the DB being replaced
' ---------------------------------------------------------------------------
  DoEvents
  Screen.MousePointer = vbHourglass    ' change to hourglass
  Kill g_strDatabase                   ' get rid of this bad database

' ---------------------------------------------------------------------------
' get the original backup copy
' ---------------------------------------------------------------------------
  m_lngRetCode = CopyFile(g_strWorkDrive & m_strBakFile, g_strDatabase, False)
  Screen.MousePointer = vbNormal        ' change the cursor back to normal
  DoEvents
    
' ---------------------------------------------------------------------------
' verify the copy was successful
' ---------------------------------------------------------------------------
  If g_clsDFI.File_Exist(g_strDatabase) Then
      ' Set the attributes back to their original state
      SetAttr g_strDatabase, intAttributes  ' Reset the DB attributes
      Replace_Original_DB = True
  Else
      
      MsgBox "Failed to replace database.  Contact support.   ", _
             vbOKOnly, "Bad Database Replace"
      Replace_Original_DB = False
  End If

Normal_Exit:
  Exit Function
  
Replace_Original_DB_Errors:
  Replace_Original_DB = False
  Resume Normal_Exit
  
End Function
Private Function Find_Work_Drive() As Boolean

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intResponse     As Integer
  Dim strMsgBoxText   As String
  Dim strTmp          As String
  
StartOver:
' ---------------------------------------------------------------------------
' Display a list of all available drives
' ---------------------------------------------------------------------------
  m_bUsedTempDir = False
  strTmp = ""
  g_strWorkDrive = ""
  g_strWorkDrive = g_clsDFI.BrowseForFolder("Pick a temporary work area")        ' Select another drive/path
  
' ---------------------------------------------------------------------------
' If CANCEL was pressed, find the Windows temp directory and create a temp
' work folder under it.  The strTmp variable is not referenced.
' ---------------------------------------------------------------------------
  If Len(Trim(g_strWorkDrive)) = 0 Then
      g_strWorkDrive = g_clsDFI.Create_A_Temp_File(g_strWorkDrive, strTmp)
  End If
  g_strWorkDrive = g_clsDFI.Add_Trailing_Slash(g_strWorkDrive)
  
' ---------------------------------------------------------------------------
' If this is the root directory of a drive then create a temporary
' subdirectory to work in.  Never do your work in the root level.  Too
' easy to corrupt the directory structure.
' ---------------------------------------------------------------------------
  If Len(Trim(g_strWorkDrive)) = 3 Then
      m_strTempDir = g_clsDFI.Create_Temp_Name(8)
      g_strWorkDrive = g_strWorkDrive & m_strTempDir
      MkDir g_strWorkDrive
      g_strWorkDrive = g_clsDFI.Add_Trailing_Slash(g_strWorkDrive)
      m_bUsedTempDir = True
  End If
  
' ---------------------------------------------------------------------------
' See if the workspace is a restricted area.  Do we have update authority?
' ---------------------------------------------------------------------------
  If g_clsDFI.IsThisRestricted(g_strWorkDrive) Then
      strMsgBoxText = "This path is a restricted area." & vbCrLf & _
                      "Do you want to select another work area?   "
      intResponse = MsgBox(strMsgBoxText, vbQuestion + vbOKCancel + vbDefaultButton1, _
                           "Restricted Area")
      If intResponse = vbOK Then
          GoTo StartOver
      Else
          Find_Work_Drive = False
          Exit Function
      End If
  End If
    
  Find_Work_Drive = True
  
End Function

Public Function Open_Exclusive() As Boolean

' ------------------------------------------------------------------------------
' Written by Kenneth Ives     kenaso@home.com
'
' See if we can get an exclusive hold of this database.  It is
' mandatory if we are going to compact it.
' ------------------------------------------------------------------------------
  
  On Error GoTo Open_Exclusive_Errors

' ---------------------------------------------------------------------------
' Definee local variables
' ---------------------------------------------------------------------------
  Dim WS As Workspace
  Dim DB As Database
  
' ---------------------------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------------------------
  Set WS = Nothing
  Set DB = Nothing
  
' ---------------------------------------------------------------------------
' Open the MS Access Database in exclusive mode
' ---------------------------------------------------------------------------
  Set WS = CreateWorkspace("", "admin", "", dbUseJet)
  Set DB = WS.OpenDatabase(g_strDatabase, True, False)
  DB.Close
  WS.Close
  Open_Exclusive = True
  

Normal_Exit:
' ---------------------------------------------------------------------------
' close everything and leave
' ---------------------------------------------------------------------------
  Set DB = Nothing
  Set WS = Nothing
  Exit Function
  
  
Open_Exclusive_Errors:
' ---------------------------------------------------------------------------
' Could not open in exclusive mode
' ---------------------------------------------------------------------------
  MsgBox "Database is currently being accessed by another application.", _
         vbOKOnly, "Cannot continue"
  
  Open_Exclusive = False
  Resume Normal_Exit
  
End Function

Public Sub Main()

' ---------------------------------------------------------------------------
' Set up the path where all of the mail processing
' will take place.
' ---------------------------------------------------------------------------
  ChDrive App.Path
  ChDir App.Path
      
' ---------------------------------------------------------------------------
' See if there is another instance of this program running
' ---------------------------------------------------------------------------
  Dim clsPrev As clsMiscForm
  Set clsPrev = New clsMiscForm
  If clsPrev.IsAnotherInstance("compmdb") Then
      Set clsPrev = Nothing
      Exit Sub
  Else
      Set clsPrev = Nothing
  End If
  
' ---------------------------------------------------------------------------
' Set up the disk processing class
' ---------------------------------------------------------------------------
  Set g_clsDFI = New clsDFInfo
  
' ---------------------------------------------------------------------------
' Get the work area drive
' ---------------------------------------------------------------------------
  Load frmDBMaint
  If Not Find_Work_Drive Then
      StopTheProgram
      Exit Sub
  End If
    
' ---------------------------------------------------------------------------
' If we have a work drive selected then start the
' maintenance
' ---------------------------------------------------------------------------
  If Len(Trim(g_strWorkDrive)) = 0 Then
      MsgBox "Application terminated because no work area was selected.", _
             vbOKOnly + vbExclamation, "No Work Area"
      StopTheProgram
  Else
      frmDBMaint.Reset_frmDBMaint
  End If
  
End Sub
Public Sub StopTheProgram()

' ---------------------------------------------------------------------------
' Upload all forms from memory and termiante this
' application
' ---------------------------------------------------------------------------
  Set g_clsDFI = Nothing
  Unload_All_Forms
  
' ---------------------------------------------------------------------------
' Delete the temporary work folder and all of its contents
' ---------------------------------------------------------------------------
  Shell App.Path & "\_DelTemp.exe " & g_strWorkDrive
  End
  
End Sub
Public Sub Unload_All_Forms()

' ---------------------------------------------------------------------------
' Written by Kenneth Ives          kenaso@home.com
'
' Unload all forms before terminating an application
' The calling module will call this routine and usually
' executes END when it returns.
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim frm As Form
  
' ---------------------------------------------------------------------------
' If the form.name property is not the same as the form
' calling this routine, then unload it and free up memory.
' ---------------------------------------------------------------------------
  For Each frm In Forms
      frm.Hide
      Unload frm
      Set frm = Nothing
  Next
  
End Sub


