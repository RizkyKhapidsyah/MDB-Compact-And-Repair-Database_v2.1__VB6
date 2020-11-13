VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDFInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Class module used to determine if a drive, folder, file, or storage device is present."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ---------------------------------------------------------
' MISCELLANEOUS CONSTANTS
' ---------------------------------------------------------
  Private Const BIF_RETURNONLYFSDIRS      As Long = 1
  Private Const INVALID_HANDLE_VALUE      As Long = -1
  Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10
  Private Const SEE_MASK_INVOKEIDLIST     As Long = &HC
  Private Const SEE_MASK_NOCLOSEPROCESS   As Long = &H40
  Private Const SEE_MASK_FLAG_NO_UI       As Long = &H400
  Private Const OF_READWRITE              As Long = &H2
  Private Const MAX_SIZE                  As Long = 260
  Private Const MAX_FILE_COUNT            As Long = 32768
  
' ------------------------------------------------------------------------
' Required for SHBrowseForFolder API call
' ------------------------------------------------------------------------
  Private Type BrowseInfo
       hWndOwner        As Long
       pIDLRoot         As Long
       pszDisplayName   As Long
       lpszTitle        As Long
       ulFlags          As Long
       lpfnCallback     As Long
       lParam           As Long
       iImage           As Long
  End Type

  Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
  
  Private Declare Function SHBrowseForFolder Lib "shell32" _
          (lpbi As BrowseInfo) As Long
  
  Private Declare Function SHGetPathFromIDList Lib "shell32" _
          (ByVal pidList As Long, ByVal lpBuffer As String) As Long

  Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
          (ByVal lpString1 As String, ByVal lpString2 As String) As Long

' ---------------------------------------------------------------------------
' Needed to build a temporary file name
' ---------------------------------------------------------------------------
  Private Declare Function GetTempPath Lib "kernel32.dll" _
          Alias "GetTempPathA" (ByVal nBufferLength As Long, _
          ByVal lpBuffer As String) As Long
            
  Private Declare Function GetTempFileName Lib "kernel32.dll" _
          Alias "GetTempFileNameA" (ByVal lpszPath As String, _
          ByVal lpPrefixString As String, ByVal wUnique As Long, _
          ByVal lpTempFileName As String) As Long
            
' ---------------------------------------------------------------------------
' Drive types
' ---------------------------------------------------------------------------
  Private Const DRV_FIXED1    As Long = 1  ' Local hard drive
  Private Const DRV_REMOVABLE As Long = 2  ' Floppy or Jaz drive
  Private Const DRV_FIXED2    As Long = 3  ' Local hard drive
  Private Const DRV_NETWORK   As Long = 4  ' Shared Network drive
  Private Const DRV_CDROM     As Long = 5  ' CD-Rom drive
  Private Const DRV_RAMDISK   As Long = 6  ' Virtual memory disk
  
  Private Declare Function GetDriveType Lib "kernel32" _
          Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
  
' ---------------------------------------------------------
' Structures defined for Date/Time API calls
' ---------------------------------------------------------
  Private Type OFSTRUCT
       cBytes               As Byte
       fFixedDisk           As Byte
       nErrCode             As Integer
       Reserved1            As Integer
       Reserved2            As Integer
       szPathName(MAX_SIZE) As Byte
  End Type

  Private Type FILETIME
       dwLowDateTime        As Long
       dwHighDateTime       As Long
  End Type

  Private Type SYSTEMTIME
       wYear                As Integer
       wMonth               As Integer
       wDayOfWeek           As Integer
       wDay                 As Integer
       wHour                As Integer
       wMinute              As Integer
       wSecond              As Integer
       wMilliseconds        As Long
  End Type

  Private Type TIME_ZONE_INFORMATION
       Bias                 As Long
       StandardName(32)     As Integer
       StandardDate         As SYSTEMTIME
       StandardBias         As Long
       DaylightName(32)     As Integer
       DaylightDate         As SYSTEMTIME
       DaylightBias         As Long
  End Type

' ---------------------------------------------------------------------------
' Structures needed for file/directory information
' ---------------------------------------------------------------------------
  Private Type WIN32_FIND_DATA
       dwFileAttributes     As Long     ' file attributes
       ftCreationTime       As FILETIME ' File creation date and time
       ftLastAccessTime     As FILETIME ' File last accessed date and time
       ftLastWriteTime      As FILETIME ' File last modified date and time
       nFileSizeHigh        As Long     ' file sizes over 2GB (2,147,483,647)
       nFileSizeLow         As Long     ' file sizes under 2GB (2,147,483,647)
       dwReserved0          As Long
       dwReserved1          As Long
       cFilename            As String * MAX_SIZE ' full file name w/o path
       cAlternate           As String * 14       ' short file name w/o path
  End Type
    
  Private Type FILE_DATA
       cName                As String  ' Name of file w/o path
       cFolder              As String  ' path of file w/o drive
       cSize                As Double  ' size of file
       cAttributes          As String  ' RHA <- Read-only, hidden, archive
       cCreationDate        As String  ' This format:  17-Feb-2000 3:15 pm
       cLastAccessDate      As String  ' This format:  17-Feb-2000 3:15 pm
       cLastModifiedDate    As String  ' This format:  17-Feb-2000 3:15 pm
  End Type
  
' ---------------------------------------------------------------------------
' Needed for file information
' ---------------------------------------------------------------------------
  Private Declare Function FindFirstFile Lib "kernel32" _
          Alias "FindFirstFileA" (ByVal lpFileName As String, _
          lpFindFileData As WIN32_FIND_DATA) As Long
            
  Private Declare Function FindNextFile Lib "kernel32" _
          Alias "FindNextFileA" (ByVal hFindFile As Long, _
          lpFindFileData As WIN32_FIND_DATA) As Long

  ' Capture the file handle
  Private Declare Function GetModuleHandle Lib "kernel32" _
          Alias "GetModuleHandleA" _
          (ByVal lpModuleName As String) As Long
  
  Private Declare Function OpenFile Lib "kernel32" _
          (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, _
          ByVal wStyle As Long) As Long
  
  ' Use the file handle to always close the last file
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hFile As Long) As Long

  Private Declare Function FindClose Lib "kernel32" _
          (ByVal hFindFile As Long) As Long
            
' ---------------------------------------------------------------------------
' Needed for a list of all available drive letters
' ---------------------------------------------------------------------------
  Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
          Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
          ByVal lpBuffer As String) As Long

' ---------------------------------------------------------
' Needed for the volume information
' ---------------------------------------------------------
  Private Declare Function GetVolumeInformation Lib "kernel32" _
          Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
          ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
          lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
          lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
          ByVal nFileSystemNameSize As Long) As Long

  Private Declare Function SetVolumeLabel Lib "kernel32" _
          Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, _
          ByVal lpVolumeName As String) As Long

' ---------------------------------------------------------
' Needed for Windows directory paths
' ---------------------------------------------------------
  Private Declare Function GetWindowsDirectory Lib "kernel32" _
          Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
          ByVal nSize As Long) As Long
  
  Private Declare Function GetSystemDirectory Lib "kernel32" _
          Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
          ByVal nSize As Long) As Long
  
' ---------------------------------------------------------
' Needed for Date/Time API calls
' ---------------------------------------------------------
  Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
  
  Private Declare Function GetFileTime Lib "kernel32" _
          (ByVal lngFileHandle As Long, lpCreationTime As FILETIME, _
          lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
  
  Private Declare Function SetFileTime Lib "kernel32" _
          (ByVal lngFileHandle As Long, lpCreationTime As FILETIME, _
          lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
  
  Private Declare Function FileTimeToSystemTime Lib "kernel32" _
          (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
  
  Private Declare Function SystemTimeToFileTime Lib "kernel32" _
          (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
  
' ---------------------------------------------------------------------------
' Needed to determine space requirements of a drive
' ---------------------------------------------------------------------------
  ' Do we have the ability to use GetDiskFreeSpaceEx
  Private Declare Function LoadLibrary Lib "kernel32" _
          Alias "LoadLibraryA" _
          (ByVal lpLibFileName As String) As Long
          
  ' Verify the proceedure address within the kernel32 dll
  Private Declare Function GetProcAddress Lib "kernel32" _
          (ByVal hModule As Long, ByVal lpProcName As String) As Long
          
  ' Decrement the DLL counter when we are finished.  This is
  ' our safety net.
  Private Declare Function FreeLibrary Lib "kernel32" _
          (ByVal hLibModule As Long) As Long
          
  ' Drives over 2GB (2,147,483,647 bytes)
  Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
          Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
          FreeBytesAvailableToCaller As Currency, _
          TotalNumberOfBytes As Currency, _
          TotalNumberOfFreeBytes As Currency) As Long
          
  ' Drives under 2GB (2,147,483,647 bytes)
  Private Declare Function GetDiskFreeSpace Lib "kernel32" _
          Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
          lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
          lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
          
' ---------------------------------------------------------------------------
' Module level variables
' ---------------------------------------------------------------------------
  Private m_lngDIndex     As Long
  Private m_lngFIndex     As Long
  Private m_lngFileCount  As Long
  Private m_lngRetCode    As Long
  Private m_strDirPath()  As String
  Private FD()            As FILE_DATA
  Private OFS             As OFSTRUCT
  Private SYS_TIME        As SYSTEMTIME
  Private WFD             As WIN32_FIND_DATA
  
Public Function Create_Temp_Name(intLength As Integer) As String

' ***************************************************************************
' Routine:       Create_Temp_Name
'
' Description:   System generated temporary name.  User determines the length
'                of the name, not to exceed eight characters.
'
' Parameters:    intLength - length of character string desired
'
' Returns:       New name
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 09-APR-2000  Kenneth Ives     Modified and documented by kenaso@home.com
' ***************************************************************************
   
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strTmp As String
  Dim intChar As Integer
  Dim n As Integer
  
' ---------------------------------------------------------------------------
' initialize varibles and randomize generator
' ---------------------------------------------------------------------------
  Randomize
  strTmp = ""
  
' ---------------------------------------------------------------------------
' Test length of name desired
' ---------------------------------------------------------------------------
  Select Case intLength
         Case Is < 1:        ' is less than or equal zero
              Create_Temp_Name = vbNullString
              Exit Function
         Case 1:        ' Nothing to generate.  Return just an underscore.
              Create_Temp_Name = "_"
              Exit Function
         Case Is > 8:   ' is greater than eight, reset back to eight
              intLength = 8
  End Select
  
' ---------------------------------------------------------------------------
' Adjust new length. One character is the underscore.
' ---------------------------------------------------------------------------
  intLength = intLength - 1
  
' ---------------------------------------------------------------------------
' Randomly generate number of characters needed
' ---------------------------------------------------------------------------
  For n = 1 To intLength
      Do
         intChar = Int(Rnd * 64) + 26    ' number between 65-90
         
         ' Test this twice just to limit the number of passes
         If intChar < 64 Then intChar = intChar + 26
         If intChar < 64 Then intChar = intChar + 26
      Loop Until (intChar > 64 And intChar < 91)
      
      ' Append new character to new name
      strTmp = strTmp & Chr(intChar)  ' convert to letter of alphabet
  Next
  
' ---------------------------------------------------------------------------
' Return the new name
' ---------------------------------------------------------------------------
  Create_Temp_Name = "_" & strTmp
  
End Function

Public Function Get_Disk_Space(ByVal strPath As String, _
                               dblTotalSpace As Double, _
                               dblFreeSpace As Double, _
                               dblUsedSpace As Double) As Boolean
  
' ***************************************************************************
' Routine:       Get_Disk_Space
'
' Description:   Get the total number of bytes, total number of free bytes,
'                and total number of used bytes on a hard drive.  This  will
'                also determine which API call to make based on the size of
'                the drive.
'
' Parameters:    strDriveLtr   - drive to be queried (ex:  C:)
'                dblTotalSpace - Value to be returned
'                dblFreeSpace  - Value to be returned
'                dblUsedSpace  - Value to be returned
'
' Return Values: Total number of bytes, total number of free bytes, and total
'                number of used bytes on a drive.
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 03-MAR-2000  Kenneth Ives     Wrote routine
' 14-MAR-2000  Kenneth Ives     Modified and documented to handle under and
'                               over 2gb drives
' ***************************************************************************

  On Error GoTo Get_Disk_Space_Errors
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngHandle                 As Long
  Dim lngProcAddress            As Long
  Dim lngSectorsPerCluster      As Long
  Dim lngBytesPerSector         As Long
  Dim lngNumberOfFreeClusters   As Long
  Dim lngTotalNumberOfClusters  As Long
  Dim curFreeBytesToCaller      As Currency
  Dim curTotalFreeBytes         As Currency
  Dim curTotalBytes             As Currency
  Dim strDriveLtr               As String
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  dblFreeSpace = 0
  dblTotalSpace = 0
  dblUsedSpace = 0
  curFreeBytesToCaller = 0
  curTotalFreeBytes = 0
  curTotalBytes = 0
  lngSectorsPerCluster = 0
  lngBytesPerSector = 0
  lngNumberOfFreeClusters = 0
  lngTotalNumberOfClusters = 0
  strDriveLtr = Trim(strPath)
  
' ---------------------------------------------------------------------------
' Verify data was passed
' ---------------------------------------------------------------------------
  If Len(Trim(strDriveLtr)) = 0 Then
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Remove any trailing backslashes
' ---------------------------------------------------------------------------
  strDriveLtr = Left(strDriveLtr, 1) & ":"

' ---------------------------------------------------------------------------
' First, determine if we can call the GetDiskFreeSpaceEx function.
' ---------------------------------------------------------------------------
  lngHandle = LoadLibrary("kernel32.Dll")
  
' ---------------------------------------------------------------------------
' if the return code is not zero, we were successful
' ---------------------------------------------------------------------------
  If lngHandle Then
      lngProcAddress = GetProcAddress(lngHandle, "GetDiskFreeSpaceExA")
      If lngProcAddress Then
          ' Make API call to query the hard drive for space
          m_lngRetCode = GetDiskFreeSpaceEx(strDriveLtr, _
                                            curFreeBytesToCaller, _
                                            curTotalBytes, _
                                            curTotalFreeBytes)
            
          ' Just to be safe, decrement Dll's usage counter.
          FreeLibrary lngHandle
            
          ' If we were successful, then multiply the results
          ' by 10000 to move the decimal to the right 4 positions
          If m_lngRetCode Then
              ' Calculate size of the hard drive
              dblTotalSpace = CDbl(curTotalBytes * 10000)
              ' Calculate the amount of free space
              dblFreeSpace = CDbl(curTotalFreeBytes * 10000)
              ' Calculate the amount of used space
              dblUsedSpace = dblTotalSpace - dblFreeSpace
          End If
      End If

      ' Just to be safe, decrement Dll's usage counter.
      FreeLibrary lngHandle
  End If

' ---------------------------------------------------------------------------
' If this drive is less than 2GB in total size then reset values and make
' the correct API call.
' ---------------------------------------------------------------------------
  If dblTotalSpace < 2147483647# Then
      dblTotalSpace = 0
      dblFreeSpace = 0
      dblUsedSpace = 0
      m_lngRetCode = 0
  Else
      Get_Disk_Space = True
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' if we get here, GetDiskFreeSpaceEx isn't available or raised an error or
' the drive is less than 2GB in size.
' ---------------------------------------------------------------------------
  m_lngRetCode = GetDiskFreeSpace(strDriveLtr, lngSectorsPerCluster, _
                                  lngBytesPerSector, lngNumberOfFreeClusters, _
                                  lngTotalNumberOfClusters)
                                  
' ---------------------------------------------------------------------------
' Test the return code to see if we had a successful call
' ---------------------------------------------------------------------------
  If m_lngRetCode Then
      ' Calculate size of the hard drive
      dblFreeSpace = CDbl(lngNumberOfFreeClusters * lngSectorsPerCluster * _
                          lngBytesPerSector)
      ' Calculate the amount of free space
      dblTotalSpace = CDbl(lngTotalNumberOfClusters * lngSectorsPerCluster * _
                          lngBytesPerSector)
      ' Calculate the amount of used space
      dblUsedSpace = dblTotalSpace - dblFreeSpace
      Get_Disk_Space = True
  Else
      Get_Disk_Space = False
  End If

Normal_Exit:
  Exit Function
  
Get_Disk_Space_Errors:
  Get_Disk_Space = False
  Resume Normal_Exit
  
End Function

Public Function Type_Of_Drive(strDriveLtr As String) As Long
Attribute Type_Of_Drive.VB_Description = "Determine the type of drive we are querying.  Returns 1-Local hard drive (Primary partition), 2-Floppy or other removeable drive, 3-Local hard drive (Secondary partition), 4-Shared Network drive, 5-CD-Rom device, 6-Virtual memory disk"

' ***************************************************************************
' Routine:       Type_Of_Drive
'
' Description:   Determine the type of drive we are querying
'
' Parameters:    strDriveLtr - drive to be queried (ex:  C)
'
' Return Values: Numeric designator for the type of drive
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 02-DEC-1999  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngType As Long
  
' ---------------------------------------------------------------------------
' Verify data was passed
' ---------------------------------------------------------------------------
  If Len(Trim(strDriveLtr)) = 0 Then
      Type_Of_Drive = 0
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Verify drive letter that was passed is in the right format
' ---------------------------------------------------------------------------
  strDriveLtr = StrConv(Left(strDriveLtr, 1), vbUpperCase)
  
' ---------------------------------------------------------------------------
' Append a colon to the drive letter
' ---------------------------------------------------------------------------
  strDriveLtr = strDriveLtr & ":"
    
' ---------------------------------------------------------------------------
' Drive type:
'    DRV_FIXED1    = 1  Local hard drive
'    DRV_REMOVABLE = 2  Floppy or other removable drive
'    DRV_FIXED2    = 3  Local hard drive
'    DRV_NETWORK   = 4  Shared Network drive
'    DRV_CDROM     = 5  CD-Rom device
'    DRV_RAMDISK   = 6  Virtual memory disk
'
' Get the drive number constant
' ---------------------------------------------------------------------------
  lngType = GetDriveType(strDriveLtr)
  Select Case lngType
         Case 1: Type_Of_Drive = 1
         Case 2: Type_Of_Drive = 2
         Case 3: Type_Of_Drive = 3
         Case 4: Type_Of_Drive = 4
         Case 5: Type_Of_Drive = 5
         Case 6: Type_Of_Drive = 6
         Case Else: Type_Of_Drive = 0  ' unknown type
  End Select
                       
End Function

Public Function File_Exist(strSearchItem As String) As Boolean
Attribute File_Exist.VB_Description = "Test to see if a drive, path, or file exists.  One drawback is this routine assumes that all filenames have an extension."

' ***************************************************************************
' Routine:       File_Exist
'
' Description:   Test to see if a file exists.  Need the fully qualified path.
'
' Syntax:        File_Exist("C:\Program Files\Desktop.ini")
'
' Parameters:    strSearchItem - Path\filename to be queried.
'
' Return Values: True or False
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 02-DEC-1999  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim WFD            As WIN32_FIND_DATA
  Dim lngType        As Long
  Dim lngHandle      As Long
  Dim lngPosition    As Long
  Dim strTmpSource   As String
  Dim strDriveLtr    As String
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strTmpSource = Trim(strSearchItem)
  If Len(Trim(strTmpSource)) = 0 Then
      File_Exist = False
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Test #1
' see if there is a drive letter, then see if it is valid.
' ---------------------------------------------------------------------------
  If Not Drive_Exist(strTmpSource) Then
      File_Exist = False
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Test #2
' Determine the drive type.
'    DRV_FIXED1    = 1  Local hard drive
'    DRV_REMOVABLE = 2  Floppy or other removeable drive
'    DRV_FIXED2    = 3  Local hard drive
'    DRV_NETWORK   = 4  Shared Network drive
'    DRV_CDROM     = 5  CD-Rom device
'    DRV_RAMDISK   = 6  Virtual memory disk
' ---------------------------------------------------------------------------
  If Mid(strTmpSource, 2, 1) = ":" Then
      ' save just the drive letter
      strDriveLtr = StrConv(Left(strTmpSource, 1), vbLowerCase)
  Else
      ' something is wrong.  Leave!
      File_Exist = False
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Get the numeric value representing the type of drive
' ---------------------------------------------------------------------------
  lngType = Type_Of_Drive(strDriveLtr)
  
' ---------------------------------------------------------------------------
' if an unknown device code of zero is returned then leave
' ---------------------------------------------------------------------------
  If lngType < 1 Or lngType > 6 Then
      File_Exist = False
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Test #3
' Check to see if this is a valid folder or are we in the root directory?
' Do this by parsing backwards thru the path looking for the first backslash.
' ---------------------------------------------------------------------------
  lngPosition = InStrRev(strSearchItem, "\")
  
' ---------------------------------------------------------------------------
' If first backslash found is in a position greater than 3 then verify the
' folders existance; otherwise, we must be in the root directory.
' ---------------------------------------------------------------------------
  If lngPosition > 3 Then
  
      ' Save just the path
      strTmpSource = Left(strSearchItem, lngPosition)
      
      ' if the folder does not exist, then leave
      If Not Folder_Exist(strTmpSource) Then
          File_Exist = False
          Exit Function
      End If
  End If

' ---------------------------------------------------------------------------
' Test #4
' Make the API call to see if the file exist
' ---------------------------------------------------------------------------
  lngHandle = FindFirstFile(strSearchItem, WFD)
  
' ---------------------------------------------------------------------------
' Return either TRUE or FALSE
' ---------------------------------------------------------------------------
  File_Exist = lngHandle <> INVALID_HANDLE_VALUE
  
' ---------------------------------------------------------------------------
' Always close the file handle
' ---------------------------------------------------------------------------
  m_lngRetCode = FindClose(lngHandle)
   
End Function

Public Function Available_Drives() As String
Attribute Available_Drives.VB_Description = "Return a string of available drive letters, each separated by a colon, backslash and null character.  (i.e.  a: c: d:)"

' ***************************************************************************
' Routine:       Available_Drives
'
' Description:   Get the list of available drive letters, each separated by
'                a null character.  (i.e.  a:\ c:\ d:\)
'
' Return Values: String of available drive letters
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 02-DEC-1999  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strDriveLtrs As String
  
' ---------------------------------------------------------------------------
' Use spaces and not null values.  Reduces your errors.
' ---------------------------------------------------------------------------
  strDriveLtrs = Space(MAX_SIZE)
  
' ---------------------------------------------------------------------------
' capture all the available drives in one long string
' ---------------------------------------------------------------------------
  m_lngRetCode = GetLogicalDriveStrings(Len(strDriveLtrs), strDriveLtrs)
  
' ---------------------------------------------------------------------------
' Remove all excess spaces before returning the string of drive letters.
' ---------------------------------------------------------------------------
  strDriveLtrs = Trim(strDriveLtrs)
  Available_Drives = strDriveLtrs

End Function
Public Function Create_A_Temp_File(strTEMP_Folder As String, strTEMP_File As String, _
                                 Optional strPrefix As String) As String

' ***************************************************************************
' Routine:       Create_A_Temp_File
'
' Description:   System generated temporary folder and file.  The folder
'                will be located in the Windows default temp directory and
'                is system generated.
'
'                Found this example at http://www.vbapi.com/
'
' Parameters:    strTEMP_Folder - Path to the temporary folder.  Created if this
'                             parameter is empty
'                strTEMP_File - This value is returned as the new temp filename
'                strPrefix - If missing, then system generated
'
' Returns:       strTEMP_Folder and strTEMP_File
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 10-MAR-2000  Kenneth Ives     Modified and documented by kenaso@home.com
' ***************************************************************************
   
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strDefaultFolder As String
  
' ---------------------------------------------------------------------------
' See if this temp folder already exist
' ---------------------------------------------------------------------------
  If Len(Trim(strTEMP_Folder)) > 0 Then
  
      ' Add a trailing backslash if it is missing
      strTEMP_Folder = Add_Trailing_Slash(strTEMP_Folder)
      
      ' see if this folder already exist
      If Folder_Exist(strTEMP_Folder) Then
          ' It does exist.  Bypass the folder
          ' creation step.
          GoTo Create_Temp_File
      Else
          ' Folder does not exist.
          ' Must be garbage in the name.
          strTEMP_Folder = ""
      End If
  End If
  
' ---------------------------------------------------------------------------
' Get Windows temporary directory path.  Use spaces and not null values.
' Reduces your errors.
' ---------------------------------------------------------------------------
  strTEMP_Folder = Space(MAX_SIZE)                        ' initialize buffer
  m_lngRetCode = GetTempPath(MAX_SIZE, strTEMP_Folder)    ' read the path name
  strDefaultFolder = Create_Temp_Name(8)
  
' ---------------------------------------------------------------------------
' extract data from the variable (ex:  "C:\Temp\")
' ---------------------------------------------------------------------------
  If m_lngRetCode Then
      ' We found the Windows default Temp folder.  Remove
      ' the trialing nulls and append the name of our
      ' temporary folder.
      strTEMP_Folder = Remove_Nulls(strTEMP_Folder)
      strTEMP_Folder = strTEMP_Folder & strDefaultFolder
  Else
      ' We did not find the Windows default temp folder
      strTEMP_Folder = "C:\" & strDefaultFolder
  End If
  
' ---------------------------------------------------------------------------
' Now we will create our TEMP folder which we will delete when this
' application terminates.
' ---------------------------------------------------------------------------
  On Error Resume Next    ' if the folder already exist, go to next line
  MkDir strTEMP_Folder    ' Create temp folder
  Err.Clear               ' reset error code
  On Error GoTo 0         ' nullify the previous "On Error" statement
  strTEMP_Folder = Add_Trailing_Slash(strTEMP_Folder)
  
Create_Temp_File:
' ---------------------------------------------------------------------------
' Use spaces and not null values.  Reduces your errors.
' ---------------------------------------------------------------------------
  strTEMP_File = Space(MAX_SIZE)
      
' ---------------------------------------------------------------------------
' Create a unique temporary file name.  I define the prefix.  A four
' digit hex number is returned.  (Ex: 49489 becomes C151)
' ---------------------------------------------------------------------------
  If Len(Trim(strPrefix)) = 0 Then
      strPrefix = Create_Temp_Name(3)
  End If
  m_lngRetCode = GetTempFileName(strTEMP_Folder, strPrefix, 0&, strTEMP_File)

' ---------------------------------------------------------------------------
' Note that the file is also created for you. Extract data from the variable
' (Ex:  "C:\Temp\_MyTmp_\resC151.TMP")
' ---------------------------------------------------------------------------
  strTEMP_File = Remove_Nulls(strTEMP_File)
  
' ---------------------------------------------------------------------------
' Return the path and name of the temp file
' ---------------------------------------------------------------------------
  Create_A_Temp_File = strTEMP_File
  
End Function

Public Function Folder_Exist(strFolder As String) As Boolean
Attribute Folder_Exist.VB_Description = "Test for the existance of a folder.  Returns True or False"
  
' ***************************************************************************
' Routine:       Folder_Exist
'
' Description:   Test for the existance of a folder.
'
' Parameters:    strFolder - full path/folder name
'
' Return Values: True or False
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 02-DEC-1999  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim WFD         As WIN32_FIND_DATA
  Dim lngHandle   As Long
   
' ---------------------------------------------------------------------------
' remove trailing slash before verifying
' ---------------------------------------------------------------------------
  strFolder = Remove_Trailing_Slash(strFolder)

' ---------------------------------------------------------------------------
' Make the API call
' ---------------------------------------------------------------------------
  lngHandle = FindFirstFile(strFolder, WFD)
   
' ---------------------------------------------------------------------------
' if a valid file handle was returned and the directory attribute is set,
' then the folder exists.
' ---------------------------------------------------------------------------
  Folder_Exist = (lngHandle <> INVALID_HANDLE_VALUE) And _
                 (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
   
' ---------------------------------------------------------------------------
' Always close the file handle
' ---------------------------------------------------------------------------
  m_lngRetCode = FindClose(lngHandle)
   
End Function

Public Function Remove_Trailing_Slash(ByVal strPath As String) As String

' ***************************************************************************
' Routine:       Remove_Trailing_Slash
'
' Description:   Removes the trailing backslash from the path if it exist
'
' Parameters:    strPath - Full path to be queried
'
' Return Values: Reformatted path
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 02-DEC-1999  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strPath = Trim(strPath)
   
' ---------------------------------------------------------------------------
' Remove the trailing backslash if it exist
' ---------------------------------------------------------------------------
  If Right(strPath, 1) = "\" Then
      Remove_Trailing_Slash = Left(strPath, Len(strPath) - 1)
  Else
      Remove_Trailing_Slash = strPath
  End If
   
End Function

Public Sub Do_Pattern_Search(strPath As String, _
                             Optional strPattern As String = "*.*")
  
' ***************************************************************************
' Routine:       Do_Pattern_Search
'
' Description:   This is the routine that is called from the main form to
'                start searching the target directory.
'
' Parameters:    strPath - The initial path to start searching in.
'                strPattern - Filename or pattern to search for.
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-FEB-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define loacl variables
' ---------------------------------------------------------------------------
  Dim intReturn As Integer
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  ReDim m_strDirPath(0)  ' empties and sizes to smallest possible
  ReDim FD(0)
  m_lngFileCount = 0
  m_lngFIndex = -1       ' After first increment, we start with index of zero
  m_lngDIndex = -1
  
' ---------------------------------------------------------------------------
' See if anything was passed
' ---------------------------------------------------------------------------
  If Len(Trim(strPath)) = 0 Then
      intReturn = MsgBox("Need fully qualified path else the root of drive C:\ " & _
                         "will become the starting point.  Do you want to continue?", _
                         vbCritical + vbYesNo + vbDefaultButton2, _
                         "Path Missing")
      If intReturn = vbNo Then
          Exit Sub
      Else
          strPath = "C:\"
      End If
  End If

' ---------------------------------------------------------------------------
' Read the directory paths and capture the file information
' ---------------------------------------------------------------------------
  Start_Processing strPath, strPattern
  
End Sub
 
Private Sub Start_Processing(strPath As String, strPattern As String)

' ***************************************************************************
' Routine:       Start_Processing
'
' Description:   Starts the processes of the initial read of the base path.
'                then we will check the directory array to process each of
'                them on a one by one basis.
'
' Parameters:    strPath - The initial path to start searching in.
'                strPattern - Filename or pattern to search for.
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-FEB-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim n As Integer
  
' ---------------------------------------------------------------------------
' Perform initial read of the directory and load the file into the file array
' ---------------------------------------------------------------------------
  If Read_Dir_Path(strPath, strPattern) Then
  
      ' See if there are any subdirectories
      ' in the directory array
       m_lngDIndex = UBound(m_strDirPath)

      If m_lngDIndex > 0 Then
          ' loop thru the directory array
          ' and process the files
          For n = 0 To m_lngDIndex
              strPath = m_strDirPath(n)            ' capture the path
              Read_Dir_Path strPath, strPattern    ' Read the directory
          Next
      End If
  End If
  
End Sub

Private Function Read_Dir_Path(strPath As String, strPattern As String) As Boolean
    
' ***************************************************************************
' Routine:       Read_Dir_Path
'
' Description:   Reads a specified path for a certain file or group of files
'                based on the pattern desired.  this will search the current
'                directory and then any subsequent subdirectories.
'
' Parameters:    strPath - The initial path to start searching in.
'                strPattern - Filename or pattern to search for.
'
' Return Values: True or False based on completion
'
' Special Logic: NONE
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-FEB-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngHandle   As Long
  Dim strFile     As String
  Dim strUnknown  As String
  Dim strSubDir   As String
    
' ---------------------------------------------------------------------------
' Add a trailing backslash if needed
' ---------------------------------------------------------------------------
  strPath = Add_Trailing_Slash(strPath)
  
' ---------------------------------------------------------------------------
' If the search for text doesn't contain any * or ?, then add asteriks before
' and after for a complete wildcard search
' ---------------------------------------------------------------------------
  If InStr(1, strPattern, "?") = 0 And InStr(1, strPattern, "*") = 0 Then
      strPattern = "*" & strPattern & "*"
  End If
  
' ---------------------------------------------------------------------------
' Read the first entry in the directory path.  Always use a trailing null to
' signal the end of the parameter.
' ---------------------------------------------------------------------------
  lngHandle = FindFirstFile(strPath & strPattern & Chr(0), WFD)
    
' ---------------------------------------------------------------------------
' If we found something, start processing
' ---------------------------------------------------------------------------
  If lngHandle > 0 Then
      Do
          ' strip trailing nulls from what we found
          strUnknown = Remove_Nulls(WFD.cFilename)
          
          ' if this is a file, then add it to the collection
          If (GetAttr(strPath & strUnknown) And vbDirectory) <> vbDirectory Then
              
              ' start loading the collection
              strFile = strUnknown   ' this is used only to clarify the process
              
              ' increment the index for the collection
              m_lngFileCount = m_lngFileCount + 1
              
              ' Test to make sure we do not exceed the max limit.
              ' If so, display an error message and then leave.
              If m_lngFileCount > MAX_FILE_COUNT Then
                  MsgBox "File count has exceeded " & Format(MAX_FILE_COUNT, "#,0") & _
                         vbCrLf & "Cannot continue.", vbExclamation + vbOKOnly, _
                         "Too many files"
                         
                  m_lngFileCount = m_lngFileCount - 1 ' decrement the count before leaving
                  Read_Dir_Path = False               ' set the flag for a bad finish
                  GoTo Normal_Exit                    ' close last handle and leave
              End If
                            
              ' store the file information to an array
              Store_File_Data strPath & strFile
          End If
        
          ' Find the next file
          m_lngRetCode = FindNextFile(lngHandle, WFD)
          
      Loop While m_lngRetCode
    
      ' After all the files are processed in a directory, we
      ' will look for any subdirectories and store their full
      ' path names in an array.
      lngHandle = FindFirstFile(strPath & "*", WFD)
    
      ' Did we find anything in this directory?
      If lngHandle > 0 Then
          Do
              ' strip trailing nulls from what we found
              strUnknown = Remove_Nulls(WFD.cFilename)
              
              ' See if this is a subdirectory that needs to be
              ' added to the directory array
              If (GetAttr(strPath & strUnknown) And vbDirectory) = vbDirectory And _
                 (strUnknown <> "." And strUnknown <> "..") Then
                 
                     ' If this is a directory and is not a  "." or ".."
                     ' then add it to the directory array
                     strSubDir = strUnknown  ' this is used only to clarify the process
                          
                     ' increment the array index by 1.
                     ' first time incremented, this will be zero
                     m_lngDIndex = m_lngDIndex + 1
                     
                     ' resize the array to hold the next element
                     ReDim Preserve m_strDirPath(m_lngDIndex)
                     
                     ' store path in the array
                     m_strDirPath(m_lngDIndex) = strPath & strSubDir
              End If
            
              ' Finds next file
              m_lngRetCode = FindNextFile(lngHandle, WFD)
          
          Loop While m_lngRetCode  ' loop while not = 0
      End If
  End If
    
' ---------------------------------------------------------------------------
' Successful finish
' ---------------------------------------------------------------------------
  Read_Dir_Path = True
  
Normal_Exit:
' ---------------------------------------------------------------------------
' Close any open file handle
' ---------------------------------------------------------------------------
  m_lngRetCode = FindClose(lngHandle)
    
End Function

Private Sub Store_File_Data(strFullPath As String)

' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim n As Integer
  Dim strTmp As String
  
' ---------------------------------------------------------------------------
' Increment the index.  First time incremented, value = 0
' ---------------------------------------------------------------------------
  m_lngFIndex = m_lngFIndex + 1

' ---------------------------------------------------------------------------
' Set aside the space for another element in the array and do not remove any
' previous stored elements.
' ---------------------------------------------------------------------------
  ReDim Preserve FD(m_lngFIndex)
  
' ---------------------------------------------------------------------------
' Convert path\filename to lowercase.  Then parse out the path from filename.
' ---------------------------------------------------------------------------
  ' convert string to lowercase
  strFullPath = StrConv(Trim(strFullPath), vbLowerCase)
  
  ' Parse backwards to determine just the file name
  n = InStrRev(strFullPath, "\")
  
  ' save the file name w/o backslashes
  FD(m_lngFIndex).cName = Mid(strFullPath, n + 1)
  strTmp = Mid(strFullPath, 1, n - 1)
  
  ' Parse forward to determine folder only w/o backslashes.
  ' Start in the third position in case the dirve is UNC.
  n = InStr(3, strTmp, "\")
  FD(m_lngFIndex).cFolder = Mid(strTmp, n + 1)
  
' ---------------------------------------------------------------------------
' Determine the size of the file
' ---------------------------------------------------------------------------
  If WFD.nFileSizeHigh > 0 Then
      ' if this file is greater than 2gb (2,147,483,647 bytes)
      FD(m_lngFIndex).cSize = WFD.nFileSizeHigh
  Else
      ' less than 2gb in size (2,147,483,647 bytes)
      FD(m_lngFIndex).cSize = WFD.nFileSizeLow
  End If
 
' ---------------------------------------------------------------------------
' Format the file dates into an acceptable format
' ---------------------------------------------------------------------------
  FD(m_lngFIndex).cCreationDate = Get_File_Date_String(WFD.ftCreationTime)
  FD(m_lngFIndex).cLastAccessDate = Get_File_Date_String(WFD.ftLastAccessTime)
  FD(m_lngFIndex).cLastModifiedDate = Get_File_Date_String(WFD.ftLastWriteTime)
 
' ---------------------------------------------------------------------------
' Get the file attributes and convert to a string format
' ---------------------------------------------------------------------------
  FD(m_lngFIndex).cAttributes = Attribute_String(WFD.dwFileAttributes)
  
End Sub

Public Function Drive_Exist(ByVal strTmpSource As String) As Boolean
Attribute Drive_Exist.VB_Description = "Checks to see if the drive we are to query is available."

' ***************************************************************************
' Routine:       Drive_Exist
'
' Description:   Checks to see if the drive we are to query is available
'
' Parameters:    strTmpSource - Full path to be queried
'
' Return Values: True or False
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 02-DEC-1999  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strDriveLtr As String
  Dim strTmpDrvLtrs  As String
  
' ---------------------------------------------------------------------------
' Capture just the drive letter adn compare it against the list of drives
' available to us.
' ---------------------------------------------------------------------------
  If Mid(strTmpSource, 2, 1) = ":" Then
  
      ' save just our drive letter
      strDriveLtr = StrConv(Left(strTmpSource, 1), vbLowerCase)
      
      ' get a list of all available drive letters
      strTmpDrvLtrs = StrConv(Available_Drives, vbLowerCase)
    
      ' is our drive letter in this list
      If InStr(1, strTmpDrvLtrs, strDriveLtr) = 0 Then
          Drive_Exist = False  ' Could not find a match
      Else
          Drive_Exist = True   ' Found a match
      End If
  Else
      Drive_Exist = False      ' something wrong happened
  End If

End Function

Public Function Add_Trailing_Slash(ByVal strPath As String) As String

' ***************************************************************************
' Routine:       Add_Trailing_Slash
'
' Description:   Add the trailing backslash from the path if it does not exist
'
' Parameters:    strPath - Full path to be queried
'
' Return Values: Reformatted path
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 02-DEC-1999  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strPath = Trim(strPath)
   
' ---------------------------------------------------------------------------
' Remove the trailing backslash if it exist
' ---------------------------------------------------------------------------
  If Right(strPath, 1) = "\" Then
      Add_Trailing_Slash = strPath
  Else
      Add_Trailing_Slash = strPath & "\"
  End If
   
End Function

Public Function Attribute_String(ByVal lngValue As Long) As String

' ***************************************************************************
' Routine:       Attribute_String
'
' Description:   The numeric value representing the attributes of a
'                directory or file are passed and the alphabetic
'                representation is returned.
'
'                Syntax::    Attribute_String 35
'                            Returns   "RHA"
'
' Parameters:    lngValue - numeric attribute representation passed into this
'                           routine
'
' Return Values: Alphabetic attributes are returned
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-FEB-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strTmp    As String
  
' ---------------------------------------------------------------------------
' Intialize variables
' ---------------------------------------------------------------------------
  strTmp = ""
  
' ---------------------------------------------------------------------------
' See if anything was passed
' ---------------------------------------------------------------------------
  If lngValue < 1 Then
      Attribute_String = strTmp
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Determine the attribute combination
' ---------------------------------------------------------------------------
  Select Case lngValue
         Case 1:  strTmp = "R"
         Case 2:  strTmp = "H"
         Case 3:  strTmp = "RH"
         Case 4:  strTmp = "S"
         Case 5:  strTmp = "RS"
         Case 6:  strTmp = "HS"
         Case 7:  strTmp = "RHS"
         Case 16: strTmp = "D"
         Case 17: strTmp = "DR"
         Case 18: strTmp = "DH"
         Case 19: strTmp = "DRH"
         Case 20: strTmp = "DS"
         Case 23: strTmp = "DRHS"
         Case 32: strTmp = "A"
         Case 33: strTmp = "RA"
         Case 34: strTmp = "HA"
         Case 35: strTmp = "RHA"
         Case 36: strTmp = "RHSA"
         Case 48: strTmp = "DA"
         Case 49: strTmp = "DRA"
         Case 50: strTmp = "DHA"
         Case 51: strTmp = "DRHA"
         Case 52: strTmp = "DSA"
         Case 53: strTmp = "DRSA"
         Case 54: strTmp = "DHSA"
         Case 55: strTmp = "DRHSA"
         Case Else: strTmp = ""
  End Select

' ---------------------------------------------------------------------------
' Return the attribute string
' ---------------------------------------------------------------------------
  Attribute_String = strTmp
  
End Function

Public Function Get_Volume_Info(strDriveLtr As String, strVolName As String, _
                                strFileSysType As String, strVolSN As String)

' ***************************************************************************
' Routine:       Get_Volume_Info
'
' Description:   Make an API call to a specific drive and capture the volume
'                information
'
' Parameters:    strDriveLtr - Drive letter to be queried
'                strVolName  - returns the volume label
'                strFileSysType - Returns the file system type (FAT, NTFS)
'                strVolSN    - Returns the volume serial number
'
' Return Values: Volume label, file system type, volume serial number
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 25-MAR-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngLowWord   As Long
  Dim lngHiWord    As Long
  Dim lngVolSn     As Long
  Dim strHiWord    As String
  Dim strLowWord   As String
  
' ---------------------------------------------------------------------------
' Initialize variables.  ' Use spaces and not null values.  Reduces your
' errors.
' ---------------------------------------------------------------------------
  strVolName = Space(MAX_SIZE)
  strFileSysType = Space(MAX_SIZE)
  strLowWord = ""
  strHiWord = ""
  lngLowWord = 0
  lngHiWord = 0
  lngVolSn = 0
  
' ---------------------------------------------------------------------------
' Save the drive letter
' ---------------------------------------------------------------------------
  strDriveLtr = Left(strDriveLtr, 1) & ":\"
  
' ---------------------------------------------------------------------------
' Make the API call to get the volume information. The next-to-last parameter
' of this call returns a string that gives the type of file system on that
' volume.
'
'    FAT, NTFS, HPFS, CDFS (CD-Rom File System)
' ---------------------------------------------------------------------------
  m_lngRetCode = GetVolumeInformation(strDriveLtr, strVolName, _
                                      MAX_SIZE, lngVolSn, 0, 0, _
                                      strFileSysType, MAX_SIZE)
  If m_lngRetCode <> 0 Then
      ' save the volume label
      strVolName = Remove_Nulls(strVolName)
      If Len(Trim(strVolName)) = 0 Then
          strVolName = "No_Label"
      End If
      
      ' Save the file system type
      strFileSysType = Remove_Nulls(strFileSysType)
      
      ' Process both sides of the volume serial number
      lngHiWord = Get_Hi_Word(lngVolSn) And &HFFFF&
      strHiWord = Pad_With_Zeros(lngHiWord, 4, 4, True)
  
      lngLowWord = Get_Low_Word(lngVolSn) And &HFFFF&
      strLowWord = Pad_With_Zeros(lngLowWord, 4, 4, True)
  
      ' reformat the volume serial number (0311-B12C)
      ' and convert to uppercase
      strVolSN = StrConv(strHiWord & "-" & strLowWord, vbUpperCase)
  End If
  
End Function

Private Function Get_Hi_Word(lngValue As Long) As Long

' ---------------------------------------------------------------------------
' Called by Get_Volume_Info routine. Get the LEFT side of the volume
' serial number
' ---------------------------------------------------------------------------
  If lngValue And &H80000000 Then
      Get_Hi_Word = (lngValue \ 65535) - 1
  Else
      Get_Hi_Word = lngValue \ 65535
  End If
  
End Function

Private Function Get_Low_Word(lngValue As Long) As Long

' ---------------------------------------------------------------------------
' Called by Get_Volume_Info routine. Get the RIGHT side of the volume
' serial number
' ---------------------------------------------------------------------------
  If lngValue And &H8000& Then
      Get_Low_Word = &H8000 Or (lngValue And &H7FFF&)
  Else
      Get_Low_Word = lngValue And &HFFFF&
  End If
  
End Function

Private Function Pad_With_Zeros(lngValue As Long, _
                                intQty2Pad As Integer, _
                                intRetLength As Integer, _
                                bPrefix As Boolean) As String

' ***************************************************************************
' Routine:       Pad_With_Zeros
'
' Description:   This routine will verify that the value returned is a full
'                four byte length. Receives a numeric value and converts to
'                hex.  Then pads with zeros either leading or trialing.
'
' Parameters:    lngValue - number to be converted to hex
'                intQty2Pad - number of zeros to pad
'                intRetLength - length of the return string
'                bPrefix - TRUE  if leading zeros
'                          FALSE if trailing zeros
'
' Return Values: hex value (full four characters)
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-FEB-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strTmp  As String
  
' ---------------------------------------------------------------------------
' Determine how many zeros are to be added to the string and are they to
' be prefixed or suffixed.
' ---------------------------------------------------------------------------
  If bPrefix Then
      strTmp = String(intQty2Pad, "0") & Trim(CStr(Hex(lngValue)))
      strTmp = Right(strTmp, intRetLength)
  Else
      strTmp = Trim(CStr(Hex(lngValue))) & String(intQty2Pad, "0")
      strTmp = Left(strTmp, intRetLength)
  End If
  
' ---------------------------------------------------------------------------
' Return the new string value
' ---------------------------------------------------------------------------
  Pad_With_Zeros = strTmp
  
End Function

Public Function Remove_Nulls(strInput As String) As String
    
' ***************************************************************************
' Routine:       Remove_Nulls
'
' Description:   Receives a data string and looks for the first null.  If
'                found, the string is truncated at that point and returned.
'
' Parameters:    strInput - Input string to be inspected
'
' Return Values: Data string with the trailing nulls
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-FEB-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intPosition As Integer
    
' ---------------------------------------------------------------------------
' Look for first null in the data string
' ---------------------------------------------------------------------------
  intPosition = InStr(1, strInput, Chr(0))
  
' ---------------------------------------------------------------------------
' if we found one, then save all the data up to that position
' ---------------------------------------------------------------------------
  If intPosition Then
      Remove_Nulls = Left(strInput, intPosition - 1)
  Else
      Remove_Nulls = strInput     ' no nulls found
  End If
      
End Function

Private Function Get_File_Date_String(varFileTime As FILETIME) As String

' ***************************************************************************
' Routine:       Get_File_Date_String
'
' Description:   Receives the timestamp for a file as to when it was last
'                updated.  the date and time are then reformatted into an
'                acceptable display format.
'
' Parameters:    varFileTime - file timestamp of when it was last updated.
'
' Return Values: Formatted timestamp (i.e., "23-Feb-2000 8:35 am")
'
' Special Logic: NONE
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-FEB-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim sngDateSerial   As Single
  Dim sngTimeSerial   As Single
  
' ---------------------------------------------------------------------------
' Convert the file timestamp to the same format as the destop system
' timestamp.
' ---------------------------------------------------------------------------
  m_lngRetCode = FileTimeToSystemTime(varFileTime, SYS_TIME)
  
' ---------------------------------------------------------------------------
' if the return value is non-zero then obtain the system time in serial
' format.  Return the completed formatted string this type of format:
'
'            17-Feb-2000 3:15 pm
' ---------------------------------------------------------------------------
  If m_lngRetCode Then
      ' If we have an invalid date assigned
      If SYS_TIME.wYear < 1980 Or _
         SYS_TIME.wMonth = 0 Or _
         SYS_TIME.wDay = 0 Then
              Get_File_Date_String = "<Unknown>"
      Else
          ' reformat the date
          sngDateSerial = DateSerial(SYS_TIME.wYear, SYS_TIME.wMonth, SYS_TIME.wDay)
          sngTimeSerial = TimeSerial(SYS_TIME.wHour, SYS_TIME.wMinute, SYS_TIME.wSecond)
          Get_File_Date_String = Format(sngDateSerial, "dd-mmm-yyyy")
          Get_File_Date_String = Get_File_Date_String & " " & _
                                 StrConv(Format(sngTimeSerial, "h:mm ampm"), vbLowerCase)
      End If
  Else
      Get_File_Date_String = "<Unknown>"
  End If

End Function
 
Public Function DelTree32(strPathToDel As String) As Boolean

' ---------------------------------------------------------------------------
'                  DelTree32 routine
'
' This code was originally written by Rod Stephens at
' http://www.vb-helper.com/HowTo/
'
' Modified and documented by Kenneth Ives  kenaso@home.com
'
' Deletes an entire directory tree regardless of the attributes of the
' directory or it's contents.  It will handle long and short filenames.
'
' Syntax:
'
'      Delete everything on drive A:
'
'           DelTree32 "A:"
'
'  Delete everything starting at "c:\Dir_1\subDir_1" and below
'  to include "\subDir_1"
'
'           DelTree32 "c:\Dir_1\subDir_1"
' ---------------------------------------------------------------------------

  On Error GoTo DelTree32_Errors
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strFilename   As String
  Dim colFilenames  As Collection
  Dim i             As Integer
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  Set colFilenames = New Collection
  
' ---------------------------------------------------------------------------
' if there is a trailing backslash then remove it.
' ---------------------------------------------------------------------------
  If Right(strPathToDel, 1) = "\" Then
      strPathToDel = Left(strPathToDel, Len(strPathToDel) - 1)
  End If
  
' ---------------------------------------------------------------------------
' Get a list of path & filenames from this folder on down.
' ---------------------------------------------------------------------------
  strFilename = Dir(strPathToDel & "\*.*", vbNormal Or vbReadOnly Or _
                                           vbHidden Or vbSystem Or _
                                           vbArchive Or vbDirectory)
    
' ---------------------------------------------------------------------------
' Loop thru the directory structure and add the path & filename to the
' collection.
' ---------------------------------------------------------------------------
  Do While Len(strFilename) > 0
      
      If (strFilename <> "..") And (strFilename <> ".") Then
          ' add to the collection
          colFilenames.Add strPathToDel & "\" & strFilename
      End If
      
      strFilename = Dir()   ' is there anything left?
      DoEvents              ' allow other processes to happen
  Loop

' ---------------------------------------------------------------------------
' Loop thru the collection and delete the colFilenames and directories
' ---------------------------------------------------------------------------
  For i = 1 To colFilenames.Count
      
      ' move the path & filename to a variable
      strFilename = colFilenames(i)
      
      ' See if it is a directory.
      If GetAttr(strFilename) And vbDirectory Then
          ' This is a directory.
          ' Delete everything in it.
          DelTree32 strFilename
      
      ' This is a file.  Delete it.
      Else
          SetAttr strFilename, vbNormal  ' reset the attributes to normal
          Kill strFilename               ' delete the file
      End If
  Next

' ---------------------------------------------------------------------------
' If this is the root directory, then leave.  Cannot delete root directory.
' ---------------------------------------------------------------------------
  If Len(strPathToDel) > 2 Then
      RmDir strPathToDel
  End If
  
  DelTree32 = True
  
Normal_Exit:
  ' always empty your objects when not in use.
  Set colFilenames = Nothing
  Exit Function
  
DelTree32_Errors:
'  Debug.Print strFilename
'  Debug.Print Err.Number & ": " & Err.Description
  DelTree32 = False
  Resume Normal_Exit
  
End Function
Public Function BrowseForFolder(Optional strPrompt As String) As String

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngFolderList  As Long
  Dim lngHandle      As Long
  Dim strTmpPath     As String
  Dim BI             As BrowseInfo

' ---------------------------------------------------------------------------
' Set up Browse Dialog Box title prompt
' ---------------------------------------------------------------------------
  If IsMissing(strPrompt) Then
      strPrompt = "Browse for work area"
  End If

' ---------------------------------------------------------------------------
' Set up Browse Dialog Box parameters
' ---------------------------------------------------------------------------
  With BI
       ' use zero as the default window handle
       .hWndOwner = 0&
       ' strPrompt will be the title on the Browse dialog box
       .lpszTitle = lstrcat(strPrompt, "")
       ' Display only directories
       .ulFlags = BIF_RETURNONLYFSDIRS
  End With

' ---------------------------------------------------------------------------
' display the browse dialog box and  create an index of the folders in the
' dialog box
' ---------------------------------------------------------------------------
  lngFolderList = SHBrowseForFolder(BI)
  
' ---------------------------------------------------------------------------
' If a folder was highlighted then format the folder name so it can be
' returned
' ---------------------------------------------------------------------------
  If lngFolderList Then
      
      ' set up a pre-padded buffer area for the folder name.
      ' Use spaces and not null values.  Reduces your errors.
      strTmpPath = Space(MAX_SIZE)
      
      ' Get the name of the folder from the list
      m_lngRetCode = SHGetPathFromIDList(lngFolderList, strTmpPath)
      Call CoTaskMemFree(lngFolderList)
      
      ' Strip any null characters from the folder name
      strTmpPath = Remove_Nulls(strTmpPath)
  End If

' ---------------------------------------------------------------------------
' return the formatted folder name
' ---------------------------------------------------------------------------
  BrowseForFolder = strTmpPath

End Function

Public Function IsThisRestricted(strPath As String) As Boolean

  On Error GoTo IsThisRestricted_Errors
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intFile      As Integer
  Dim strTmpPath   As String
  
' ---------------------------------------------------------------------------
' initialize local variables
' ---------------------------------------------------------------------------
  intFile = FreeFile
  strTmpPath = Trim(strPath)
  
' ---------------------------------------------------------------------------
' See if there is a file parameter on the end of these path strings.  If so,
' remove it so we can see if we have access to their areas.
' ---------------------------------------------------------------------------
  If Len(Trim(strTmpPath)) > 0 Then
  
      ' add trailing "\" if missing and the dummy filename
      strTmpPath = Add_Trailing_Slash(strTmpPath)
      strTmpPath = strTmpPath & "_Test.$$$"
      
      ' open the new file on the drive to make sure
      ' we can write/delete
      Open strTmpPath For Output As #intFile  ' Open the file
      Print #intFile, "x"                     ' Write to it
      Close #intFile                          ' Save it to disk
      Kill strTmpPath                         ' Delete the file
      '
      IsThisRestricted = False
  Else
      IsThisRestricted = True
  End If
  
  
Normal_Exit:
  Exit Function
  
  
IsThisRestricted_Errors:
  IsThisRestricted = True
  Resume Normal_Exit
  
End Function

Public Function Shrink_2_Fit(strInText As String, intMaxLen As Integer) As String

' ---------------------------------------------------------------------------
' This function will shorten a directory name to the length passed to the
' Max parameter.
'
' Syntax:
' strTmp = Shrink_2_Fit("C:\Program Files\Netscape\Navigator\Programs\Bookmark.htm", 30)
'
' Returns:   strTmp = "C:\...\Programs\Bookmark.htm"
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim n            As Integer
  Dim intStart     As Integer
  Dim intNeeded    As Integer
  Dim strTmp       As String
  Dim strPrefix    As String
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strTmp = Trim(strInText)

' ---------------------------------------------------------------------------
' if the string is equal to or less than the desired length then leave
' this routine
' ---------------------------------------------------------------------------
  If Len(strTmp) <= intMaxLen Then
      Shrink_2_Fit = strTmp
      Exit Function
  End If

' ---------------------------------------------------------------------------
' Prepare the drive and the elipces.  Start in the third position in case
' this is a UNC path
' ---------------------------------------------------------------------------
  n = InStr(3, strTmp, "\")
  intStart = n + 1
  intNeeded = intMaxLen - (n + 3)      ' 3 elipces
  strPrefix = Left(strTmp, n) & "..."
  
' ---------------------------------------------------------------------------
' prepare the suffix part of the path to the desired length
' ---------------------------------------------------------------------------
  Do
      n = InStr(intStart, strTmp, "\")  ' look for a backslash
      If n > 0 Then
          strTmp = Mid(strTmp, n)  ' Reformat string length
      Else
          Exit Do                  ' could not find "\", so leave
      End If
      
  Loop Until intNeeded >= Len(strTmp)
  
' ---------------------------------------------------------------------------
' Return the shortened string
' ---------------------------------------------------------------------------
  Shrink_2_Fit = strPrefix & strTmp
  
End Function

