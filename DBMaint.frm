VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDBMaint 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Compact/Repair Database v2.1"
   ClientHeight    =   1545
   ClientLeft      =   2775
   ClientTop       =   3375
   ClientWidth     =   5820
   Icon            =   "DBMaint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.Animation aniAVI 
      Height          =   1065
      Left            =   75
      TabIndex        =   3
      Top             =   150
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1879
      _Version        =   393216
      BackColor       =   255
      FullWidth       =   96
      FullHeight      =   71
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   5325
      Top             =   1125
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   1575
      ScaleHeight     =   1005
      ScaleWidth      =   4080
      TabIndex        =   0
      Top             =   150
      Width           =   4140
      Begin VB.Label lblDBName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   150
         TabIndex        =   2
         Top             =   525
         Width           =   3840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Now performing maintenance on:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   1
         Top             =   150
         Width           =   3915
      End
   End
   Begin VB.Label lblCredit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   1275
      Width           =   5640
   End
End
Attribute VB_Name = "frmDBMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------
' Define module level variables
' ---------------------------------------------------------------------------
  Private m_clsRes As clsResFile
  
Private Function GetDatabase() As Boolean

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim sAppPath As String
  Dim sFilters As String
  Dim sCmdLine As String
  Dim iCmdLnLen As Integer

' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  sAppPath = "C:\"
  sFilters = "Access Files (*.mdb)|*.mdb" & "All Files (*.*)|*.*"
  g_strDatabase = ""

' ---------------------------------------------------------------------------
' Get command line arguments.
' ---------------------------------------------------------------------------
  sCmdLine = Command()
  iCmdLnLen = Len(sCmdLine)
 
' ---------------------------------------------------------------------------
' See if there is the name of a database on the command line
' ---------------------------------------------------------------------------
  If iCmdLnLen > 0 Then
      g_strDatabase = sCmdLine
      GoTo Normal_Exit
  End If
  
' ---------------------------------------------------------------------------
' Get the location of the database.  Display the File Open dialog box
' ---------------------------------------------------------------------------
  ' Set CancelError is True
  frmDBMaint.CDialog.CancelError = True
  On Error GoTo Cancel_Was_Pressed
  
  With frmDBMaint.CDialog
       .DialogTitle = "Select database to compact"
       .DefaultExt = "*.mdb"
       .Filter = sFilters
       .flags = cdlOFNHideReadOnly
       .InitDir = sAppPath
       .FilterIndex = 1                 ' Specify default filter
       .FileName = "*.mdb"
       .ShowOpen                        ' Display the Open dialog box
  End With
  
' ---------------------------------------------------------------------------
' Save the name of the item selected
' ---------------------------------------------------------------------------
  g_strDatabase = CDialog.FileName
  
  
Normal_Exit:
  GetDatabase = True
  Exit Function
  
  
Cancel_Was_Pressed:
  On Error GoTo 0
  GetDatabase = False
  g_strDatabase = ""

End Function
Public Sub Reset_frmDBMaint()
  
' ---------------------------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------------------------
  g_strResFolder = ""
  g_strResFile = ""
  
' ---------------------------------------------------------------------------
' Get the database
' ---------------------------------------------------------------------------
  If GetDatabase Then
      If ValidDatabase Then
          
          ' get the avi file from the RES file
          Set m_clsRes = New clsResFile
          If Not m_clsRes.Load_Res_Data(101, "AVI", "AVI", g_strResFolder, g_strResFile) Then
              Set m_clsRes = Nothing
              StopTheProgram
              Exit Sub
          End If
          
          ' see if the temp file was created
          If Len(Trim(g_strResFile)) = 0 Then
              Set m_clsRes = Nothing
              StopTheProgram
              Exit Sub
          Else
              Set m_clsRes = Nothing
          End If
          
          ' Show the form and start compacting the database
          With frmDBMaint
               .aniAVI.Open g_strResFile
               .aniAVI.AutoPlay = True
               .lblCredit = "Freeware by Kenneth Ives  kenaso@home.com"
               .lblDBName.Caption = g_clsDFI.Shrink_2_Fit(g_strDatabase, 40)
               .Show vbModeless
               .Refresh
          End With
          '
          CompactMDB    ' compact the database
      Else
          Set m_clsRes = Nothing
          StopTheProgram
      End If
  Else
      Set m_clsRes = Nothing
      StopTheProgram
  End If
  
End Sub

Private Function ValidDatabase() As Boolean

' ---------------------------------------------------------------------------
' Define local variaables
' ---------------------------------------------------------------------------
  Dim n               As Integer
  Dim intResponse     As Integer
  Dim strMsgBoxText   As String
  Dim strTmpPath      As String
  
ValidDatabase_StartOver:
' ---------------------------------------------------------------------------
' Remove trailing spaces
' ---------------------------------------------------------------------------
  g_strDatabase = Trim(g_strDatabase)
  n = InStrRev(g_strDatabase, "\")
  strTmpPath = Left(g_strDatabase, n)
  
' ---------------------------------------------------------------------------
' Is there something there
' ---------------------------------------------------------------------------
  If Len(g_strDatabase) = 0 Then
      MsgBox "No database selected", vbOKOnly, "No MDB selected"
      ValidDatabase = False
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Is this a database
' ---------------------------------------------------------------------------
  If StrConv(Right(g_strDatabase, 4), vbUpperCase) <> ".MDB" Then
      MsgBox "This is not a database." & vbLf & UCase(g_strDatabase), _
             vbOKOnly, "No MDB selected"
      ValidDatabase = False
      Exit Function
  End If

' ---------------------------------------------------------------------------
' Does the database exist
' ---------------------------------------------------------------------------
  If Not g_clsDFI.File_Exist(g_strDatabase) Then
      MsgBox "Database cannot be found at this location." & vbLf & g_strDatabase, _
             vbOKOnly, "No MDB selected"
      ValidDatabase = False
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' See if the database area is restricted.  Do we have update authority?
' ---------------------------------------------------------------------------
  If g_clsDFI.IsThisRestricted(strTmpPath) Then
      strMsgBoxText = "This is a restricted area.  The database cannot be compacted and replaced." & vbCrLf & _
                      "Do you want to select a different database?"
      intResponse = MsgBox(strMsgBoxText, vbQuestion + vbOKCancel + vbDefaultButton1, "Restricted Area")
      If intResponse = vbOK Then
          GoTo ValidDatabase_StartOver
      Else
          ValidDatabase = False
          Exit Function
      End If
  End If
        
' ---------------------------------------------------------------------------
' We are ready to go
' ---------------------------------------------------------------------------
  ValidDatabase = True
  
End Function

Private Sub Form_Load()

' ---------------------------------------------------------------------------
' I am now using classes to save memory.  A class is only loaded when it is
' accessed.  When you are finished with it, it is unloaded from memory.
' Once a BAS module is accessed, it is loaded into memory for the duration
' of the application.
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim clsPrep As clsMiscForm
  
' ---------------------------------------------------------------------------
' Center the form on the screen
' ---------------------------------------------------------------------------
  Set clsPrep = New clsMiscForm
  clsPrep.CenterForm frmDBMaint
  clsPrep.RemoveX frmDBMaint
  Set clsPrep = Nothing
  
' ---------------------------------------------------------------------------
' Hide this form temporarily
' ---------------------------------------------------------------------------
  Hide
    
End Sub

