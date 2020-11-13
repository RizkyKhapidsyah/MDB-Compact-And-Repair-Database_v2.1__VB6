Attribute VB_Name = "basDelTemp"
Option Explicit

Sub Main()

' ---------------------------------------------------------
' _DelTemp
' Written by Kenneth Ives  kenaso@home.com
'
' Wrote this application to remove any temporary folder
' that my applications may have created.  The normal way
' this is used is a call from a BAS module after all the
' forms and classes have been freed from memory.  This
' small program is normally distributed with an application
' EXE and stored in the application directory.
'
' SYNTAX:
'   Shell app.path & "\_DelTemp.exe" full_path_of_temp_folder, vbHide
'
' Examples of usage:
'   CompMDB.exe  BAS module - ShutDownProgram
'   ResDemo.exe  Form - Form_QueryUnload event
' ---------------------------------------------------------

' ---------------------------------------------------------
' Define local variables
' ---------------------------------------------------------
  Dim clsDFI        As clsDFInfo
  Dim strTmpPath    As String
  Dim strTmp        As String
  Dim n             As Integer
  
' ---------------------------------------------------------
' Initialize variables and read the parameters passed to
' this program (Name of folder to delete).  Folder must
' be prefixed with an underscore ("_") to prevent deleting
' another application folder.
' ---------------------------------------------------------
  Set clsDFI = New clsDFInfo
  strTmpPath = Command
  strTmpPath = Trim(strTmpPath)
  
' ---------------------------------------------------------
' Was anything passed to this application
' ---------------------------------------------------------
  If Len(strTmpPath) = 0 Then
      GoTo Normal_Exit
  End If
  
' ---------------------------------------------------------
' Strip the trailing "\" if it exist
' ---------------------------------------------------------
  If Right(strTmpPath, 1) = "\" Then
      strTmpPath = Left(strTmpPath, Len(strTmpPath) - 1)
  End If
  
' ---------------------------------------------------------
' Strip the folder name.
' ---------------------------------------------------------
  n = InStrRev(strTmpPath, "\")
  strTmp = Mid(strTmpPath, n + 1)
  
' ---------------------------------------------------------
' Test the prefix
' ---------------------------------------------------------
  If Left(strTmp, 1) = "_" Then
      GoTo Start_Work
  Else
      GoTo Normal_Exit
  End If
  
Start_Work:
' ---------------------------------------------------------
' Delete the temporary folder and its contents
' ---------------------------------------------------------
  If clsDFI.Folder_Exist(strTmpPath) Then
      ' It does exist.  Now delete it and everything in it.
      clsDFI.DelTree32 strTmpPath
  End If
  
Normal_Exit:
' ---------------------------------------------------------
' cleanup and terminate application
' ---------------------------------------------------------
  Set clsDFI = Nothing
  End
  
End Sub


