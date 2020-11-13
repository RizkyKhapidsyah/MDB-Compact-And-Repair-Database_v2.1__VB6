VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Function Rename_File(strFilename As String, strExtension As String)

' ***************************************************************************
' Routine:       Rename_File
'
' Description:   Returns an object from a resource file.  This object is
'                stored when you select the "Add a CUSTOM" icon inside the
'                resource editor window.  If you do not want to use the
'                reference "CUSTOM" then right mouse click the resource ID
'                and select properties.  Change the name.
'
'                This is a generic process to load WAV, MIDI, and AVI
'                type files from a resource file.  You just have to
'                keep track of them.  Once the file is identified, it can
'                be accessed by the appropriate control.
'
'                Animation control - AVI files
'                MultiMedia Control - WAV and MIDI files
'
' Parameters:    strFilename - full path and filename that will be modified
'                strExtension - name of the new extension for the file
'
' Return Values: Filename with new extension.
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 03-APR-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strTmp   As String
  Dim n        As Integer
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strTmp = ""
  
' ---------------------------------------------------------------------------
' parse the filenaem backwards to find the first period
' ---------------------------------------------------------------------------
  n = InStrRev(strFilename, ".")
  
' ---------------------------------------------------------------------------
' Reformat the new name.
' ---------------------------------------------------------------------------
  strTmp = Left(strFilename, n) & strExtension
    
' ---------------------------------------------------------------------------
' Rename the temp file to the new filename
' ---------------------------------------------------------------------------
  Name strFilename As strTmp
  
' ---------------------------------------------------------------------------
' Update the variable name to match the new filename
' ---------------------------------------------------------------------------
  strFilename = strTmp
  
End Function
  
Public Function Load_Res_Data(ByVal lngResID As Long, _
                              ByVal strResCustomName As String, _
                              strFileExt As String, _
                              strTEMP_Folder As String, _
                              strTEMP_File As String) As Boolean
                                  
' ***************************************************************************
' Routine:       Load_Res_Data
'
' Description:   Returns an object from a resource file.  This object is
'                stored when you select the "Add a CUSTOM" icon inside the
'                resource editor window.  If you do not want to use the
'                reference "CUSTOM" then right mouse click the resource ID
'                and select properties.  Change the name.
'
'                This is a generic process to load WAV, MIDI, and AVI
'                type files from a resource file.  You just have to
'                keep track of them.  Once the file is identified, it can
'                be accessed by the appropriate control.
'
'                Animation control - AVI files
'                MultiMedia Control - WAV and MIDI files
'
' Parameters:    lngResID - long integer index assigned to this data source
'                           in the resource file
'                strResCustomName - User assigns this name if they do not want
'                           to use the name "CUSTOM" in the resource file
'                strFileExt - name of the temp file extension when finished.
'                strTEMP_Folder - name of temp folder.  If we had to create it,
'                           we must track the name to be able to delete it.
'                strTmpResFile - full path and filename of the temp data file
'
' Return Values: TRUE or FALSE based on completion
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 03-APR-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim clsFile As clsDFInfo
  
' ---------------------------------------------------------------------------
' Initialize return variables
' ---------------------------------------------------------------------------
  Set clsFile = New clsDFInfo
  strTEMP_File = ""
  
' ---------------------------------------------------------------------------
' Create a temp file
' ---------------------------------------------------------------------------
  clsFile.Create_A_Temp_File strTEMP_Folder, strTEMP_File
  Set clsFile = Nothing
  
' ---------------------------------------------------------------------------
' Save the resource item to disk
' ---------------------------------------------------------------------------
  If Save_Data_As_File(lngResID, strResCustomName, strTEMP_File, strFileExt) Then
      Load_Res_Data = True
  Else
      Load_Res_Data = False
  End If
    
End Function
Public Function Load_Pic_Data(ByVal lngResID As Long, _
                              ByVal strResCustomName As String, _
                              strFileExt As String, _
                              strTEMP_Folder As String) As Picture
                                  
' ***************************************************************************
' Routine:       Load_Pic_Data
'
' Description:   Returns an object from a resource file.  This object is
'                stored when you select the "Add a CUSTOM" icon inside the
'                resource editor window.  If you do not want to use the
'                reference "CUSTOM" then right mouse click the resource ID
'                and select properties.  Change the name.
'
'                This is a generic process to load GIF and JPG type files
'                from a resource file into a picture control.  You just have
'                to keep track of them.  Once the file is identified, it can
'                be accessed by the appropriate control.
'
'                Picture and Image controls
'
' Parameters:    lngResID - long integer index assigned to this data source
'                           in the resource file
'                strResCustomName - User assigns this name if they do not want
'                           to use the name "CUSTOM" in the resource file
'                strFileExt - name of the temp file extension when finished.
'                strTEMP_Folder - name of temp folder.  If we had to create it,
'                           we must track the name to be able to delete it.
'
' Return Values: Requested image
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 03-APR-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim clsFile         As clsDFInfo
  Dim strTEMP_File    As String
  
' ---------------------------------------------------------------------------
' Initialize return variables
' ---------------------------------------------------------------------------
  strTEMP_File = ""
  Set clsFile = New clsDFInfo
  
' ---------------------------------------------------------------------------
' Create a temp file
' ---------------------------------------------------------------------------
  clsFile.Create_A_Temp_File strTEMP_Folder, strTEMP_File
  Set clsFile = Nothing
  
' ---------------------------------------------------------------------------
' Save the resource item to disk
' ---------------------------------------------------------------------------
  If Save_Data_As_File(lngResID, strResCustomName, strTEMP_File, strFileExt) Then
      Set Load_Pic_Data = LoadPicture(strTEMP_File)
      Kill strTEMP_File
  End If
    
End Function

Private Function Save_Data_As_File(ByVal lngResID As Long, _
                                   ByVal strResCustomName As String, _
                                   strTEMP_File As String, _
                                   strFileExt As String) As Boolean
                                   
' ***************************************************************************
' Routine:       Save_Data_As_File
'
' Description:   Receive the resource data into a byte array and then
'                transfers that byte array to a temporary file.  After we
'                are finished with the temp resource file, it will be
'                deleted.  Once I have this as a file, it is easier to
'                access and manipulate them.
'
'                This is a generic process to load WAV, MIDI, GIF, JPG,
'                AVI, etc type files from a resource file.
'
' Parameters:    lngResID - long integer index assigned to this data source
'                           in the resource file
'                strResCustomName - User assigns this name if they do not want
'                           to use the name "CUSTOM" in the resource file
'                strTEMP_file - full path and filename of the temp data file
'                strFileExt - name of the temp file extension when finished.
'
' Return Values: TRUE or FALSE based on completion
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 03-APR-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intFile    As Integer
  Dim bytData()  As Byte
    
  On Error GoTo Save_Res_Data_Errors
' ---------------------------------------------------------------------------
' Retrieve the resource data into the byte array
' ---------------------------------------------------------------------------
  bytData = LoadResData(lngResID, strResCustomName)
  
' ---------------------------------------------------------------------------
' Write data to a temp file
' ---------------------------------------------------------------------------
  intFile = FreeFile
  Open strTEMP_File For Binary Access Write As #intFile
  Put #intFile, , bytData()
  Close #intFile
    
' ---------------------------------------------------------------------------
' Success  :-)
' ---------------------------------------------------------------------------
  Rename_File strTEMP_File, strFileExt
  Save_Data_As_File = True
  Exit Function

Save_Res_Data_Errors:
' ---------------------------------------------------------------------------
' Failure  :-(
' ---------------------------------------------------------------------------
  Save_Data_As_File = False

End Function
