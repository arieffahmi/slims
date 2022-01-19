VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SLIMS"
   ClientHeight    =   8460
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   720
      Top             =   8520
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   12
      Top             =   7680
      Width           =   6255
   End
   Begin VB.CommandButton cmdDisplayActiveLoanApplications 
      Caption         =   "Display Active Loan Applications"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   11
      Top             =   6840
      Width           =   6255
   End
   Begin VB.CommandButton cmdLoanTermination 
      Caption         =   "Loan Termination"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   6000
      Width           =   6255
   End
   Begin VB.CommandButton cmdLoanApplication 
      Caption         =   "Loan Application"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   5160
      Width           =   6255
   End
   Begin VB.CommandButton cmdDisplayProgrammeList 
      Caption         =   "Display Programme List"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   8
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton cmdDeleteProgrammeData 
      Caption         =   "Delete Programme Data"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   7
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton cmdUpdateProgrammeData 
      Caption         =   "Update Programme Data"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdAddNewProgramme 
      Caption         =   "Add New Programme"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cmdDisplayStudentList 
      Caption         =   "Display Student List"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton cmdDeleteStudentData 
      Caption         =   "Delete Student Data"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton cmdUpdateStudentData 
      Caption         =   "Update Student Data"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdAddNewStudent 
      Caption         =   "Add New Student"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label lblSLIMS 
      Alignment       =   2  'Center
      Caption         =   "Scholarship and Loan Information Management System"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Public CON As New ADODB.Connection
    Public RS As New ADODB.Recordset
    Public strSQL As String
    
    Public strInput As String
    
    Public strStudentID As String
    Public strStudentName As String
    Public strStudentPhone As String
    Public strStudentEmail As String
    
    Public strProgrammeCode As String
    Public strProgrammeName As String
    Public strProgrammeSemesters As String
    Public curProgrammeFee As Currency

Private Sub cmdDeleteProgrammeData_Click()
    strInput = InputBox("Programme Code:")
    strSQL = "select * from Programme where ProgrammeCode = '" & strInput & "'"
    
    OpenConnection
    
    If RS.EOF Then
        MsgBox "Programme not found"
    Else
        RS.MoveFirst
        
        strProgrammeCode = RS.Fields!ProgrammeCode
        strProgrammeName = RS.Fields!ProgrammeName
        curProgrammeFee = Val(RS.Fields!ProgrammeFee)
        strProgrammeSemesters = RS.Fields!ProgrammeSemesters
        
        confirmDelete = MsgBox("Confirm delete?" & vbNewLine & vbNewLine & "Programme Code: " & strProgrammeCode & vbNewLine & "Programme Name: " & strProgrammeName & vbNewLine & "Programme Fee: " & FormatCurrency(curProgrammeFee) & vbNewLine & "Programme Semesters: " & strProgrammeSemesters, vbYesNo)
        
        If confirmDelete = vbYes Then
            RS.Delete
            MsgBox "Programe data deleted"
        End If
        
    End If
    
    CloseConnection
    
End Sub

Private Sub cmdDeleteStudentData_Click()
    strInput = InputBox("Student ID:")
    strSQL = "select * from Student where studentID = '" & strInput & "'"
    
    OpenConnection
     
    If RS.EOF Then
        MsgBox "Student not found"
    Else
        RS.MoveFirst
        
        strStudentID = RS.Fields!studentID
        strStudentName = RS.Fields!studentName
        strStudentPhone = RS.Fields!studentPhone
        strStudentEmail = RS.Fields!studentEmail
        
        confirmDelete = MsgBox("Confirm delete?" & vbNewLine & vbNewLine & "Student ID: " & strStudentID & vbNewLine & "Student Name: " & strStudentName & vbNewLine & "Student Phone: " & strStudentPhone & vbNewLine & "Student Email: " & strStudentEmail, vbYesNo)
        
        If confirmDelete = vbYes Then
            RS.Delete
            MsgBox "Student data deleted"
        End If
    End If
    
    CloseConnection

End Sub

Private Sub cmdDisplayProgrammeList_Click()
    Unload Me
    frmDisplayProgrammeList.Show
End Sub

Private Sub cmdDisplayActiveLoanApplications_Click()
    Unload Me
    frmDisplayActiveLoanApplications.Show
End Sub

Private Sub cmdDisplayStudentList_Click()
    Unload Me
    frmDisplayStudentList.Show
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdAddNewProgramme_Click()
    Unload Me
    frmAddNewProgramme.Show
End Sub

Private Sub cmdAddNewStudent_Click()
    Unload Me
    frmAddNewStudent.Show
End Sub

Private Sub cmdLoanTermination_Click()
    Unload Me
    frmLoanTermination.Show
End Sub

Private Sub cmdLoanApplication_Click()
    Unload Me
    frmLoanApplication.Show
End Sub

Private Sub cmdUpdateProgrammeData_Click()

    Dim fucd As New frmUpdateProgrammeData

    strInput = InputBox("Programme Code:")
    strSQL = "select * from Programme where ProgrammeCode = '" & strInput & "'"
    
    OpenConnection

    If RS.EOF Then
        MsgBox "Programme not found"
        CloseConnection
    Else
        
        RS.MoveFirst
        
        fucd.strProgrammeCode = RS.Fields!ProgrammeCode.Value
        fucd.strProgrammeName = RS.Fields!ProgrammeName.Value
        fucd.strProgrammeSemesters = RS.Fields!ProgrammeSemesters.Value
        fucd.curProgrammeFee = RS.Fields!ProgrammeFee.Value
        
        CloseConnection
        
        Unload Me
        fucd.Show
    End If
    
End Sub

Private Sub cmdUpdateStudentData_Click()

    Dim fusd As New frmUpdateStudentData

    strInput = InputBox("Student ID:")
    strSQL = "select * from Student where studentID = '" & strInput & "'"
    
    OpenConnection

    If RS.EOF Then
        MsgBox "Student not found"
        CloseConnection
    Else
        
        RS.MoveFirst
        
        fusd.strStudentID = RS.Fields!studentID.Value
        fusd.strStudentName = RS.Fields!studentName.Value
        fusd.strStudentPhone = RS.Fields!studentPhone.Value
        fusd.strStudentEmail = RS.Fields!studentEmail.Value
        
        CloseConnection
        
        Unload Me
        fusd.Show
    End If
    
End Sub

Public Sub CloseConnection()
    If RS.State And adStateOpen Then
        RS.Close
        CON.Close
    End If
End Sub

Public Sub OpenConnection()
    If RS.State = False Then
        CON.Open ("provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\data\SLIMS.mdb")
        RS.Open strSQL, CON, adOpenDynamic, adLockOptimistic
    End If
End Sub
