VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   ClientHeight    =   1500
   ClientLeft      =   48
   ClientTop       =   48
   ClientWidth     =   4464
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   4464
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar prgDatabase 
      Height          =   312
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4272
      _ExtentX        =   7535
      _ExtentY        =   550
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar prgTable 
      Height          =   312
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   4272
      _ExtentX        =   7535
      _ExtentY        =   550
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblDatabaseProgress 
      Caption         =   "Database progress"
      Height          =   192
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1512
   End
   Begin VB.Label lblTableTableProgress 
      AutoSize        =   -1  'True
      Height          =   192
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   36
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim mDB As Database
  Dim mRS(3) As Recordset
'  Dim TableNames
  'Center the screen and display it
  frmProgress.Left = (Screen.Width - frmProgress.Width) / 2
  frmProgress.Top = (Screen.Height - frmProgress.Height) / 2
  frmProgress.Show
  
  MsgBox "Click ok to start the program", vbOKOnly + vbInformation
  
  frmProgress.MousePointer = vbHourglass
  
  'Create an array of the table names to assign to the mRS array
'  TableNames = Array("Table1", "Operation", "Maintenance", "F1")
  'Pass the total number of records in the database to initializeProgressBar
  InitializeProgressBar GetTotalRecordCount(App.Path & "\MyDatabase97.mdb", TableNames)
  'open the database for processing
  Set mDB = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\MyDatabase97.mdb")
  
  'open the tables for processing
  For i = LBound(mRS) To UBound(mRS)
    Set mRS(i) = mDB.OpenRecordset("Table" & CStr(i + 1))
  Next
  
  ReadDatabase mRS
  mDB.Close
  frmProgress.MousePointer = vbNormal
  MsgBox "Click on the form to exit program", vbOKOnly + vbInformation
  
End Sub

Private Sub lblDatabaseProgress_Click()
  Unload Me
End Sub

Private Sub lblTableTableProgress_Click()
  Unload Me
End Sub

Private Sub prgDatabase_Click()
  Unload Me
End Sub

Private Sub prgTable_Click()
  Unload Me
End Sub

Private Sub InitializeProgressBar(MaxValue As Long)
  With frmProgress
    .prgDatabase.Min = 0
    .prgDatabase.Value = 0
    If MaxValue = 0 Then MaxValue = 1
    .prgDatabase.Max = MaxValue
    .Refresh
  End With
    
End Sub

Private Function GetTotalRecordCount(DatabaseName As String, TableNames) As Long
  Dim lDb As Database
  Dim lRs As Recordset
  Dim i As Integer
  
  Set lDb = DBEngine.Workspaces(0).OpenDatabase(DatabaseName)
  
  For i = 1 To 4
    Set lRs = lDb.OpenRecordset("Table" & CStr(i))
    GetTotalRecordCount = GetTotalRecordCount + lRs.RecordCount
    lRs.Close
  Next
End Function


Private Sub ReadDatabase(mRS() As Recordset)
    
  For i = LBound(mRS) To UBound(mRS)
    With mRS(i)
      'Make sure the table has records in it.
      If .RecordCount > 0 Then
        'Show user which table is being read
        frmProgress.lblTableTableProgress.Caption = "Now loading " & .Name
        frmProgress.Refresh
        
        'initialize the progress bar to 0 for the table
        frmProgress.prgTable.Value = 0
        'Make sure we are starting at the first location of the table
        .MoveFirst
        While Not .EOF
          'Sleep is used just to simulate time elapsing for processing the data
          Sleep 5
          'Process your data from the table here.
          .MoveNext
          frmProgress.prgDatabase.Value = frmProgress.prgDatabase.Value + 1
          frmProgress.prgTable.Value = .PercentPosition
        Wend
        mRS(i).Close
      End If
    End With
  Next
End Sub

