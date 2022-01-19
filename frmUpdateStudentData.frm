VERSION 5.00
Begin VB.Form frmUpdateStudentData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Student Data"
   ClientHeight    =   5790
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtStudentID 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox txtStudentName 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox txtStudentPhone 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   4095
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   6
      Top             =   4800
      Width           =   2775
   End
   Begin VB.TextBox txtStudentEmail 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Label lblStudentID 
      Caption         =   "Student ID:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblStudentName 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblStudentPhone 
      Caption         =   "Phone:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblStudentEmail 
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblUpdateStudentData 
      Alignment       =   2  'Center
      Caption         =   "Update Student Data"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileHome 
         Caption         =   "&Home"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmUpdateStudentData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim strInput As String
    Dim strSQL As String
    
    Public strStudentID As String
    Public strStudentName As String
    Public strStudentPhone As String
    Public strStudentEmail As String
    
Private Sub cmdClear_Click()

    txtStudentID.Text = ""
    txtStudentName.Text = ""
    txtStudentPhone.Text = ""
    txtStudentEmail.Text = ""
    
End Sub

Private Sub cmdSubmit_Click()
    
    frmMain.OpenConnection
    
    With frmMain.RS
    
    !studentID = txtStudentID.Text
    !studentName = txtStudentName.Text
    !studentPhone = txtStudentPhone.Text
    !studentEmail = txtStudentEmail.Text
    .Update
    End With
    
    MsgBox "Student data updated"
    
    cmdClear_Click
    
    mnuFileHome_Click
End Sub

Private Sub Form_Load()
    strSQL = "select * from Student where studentID = '" & strStudentID & "'"
    
    frmMain.OpenConnection

    txtStudentID.Text = strStudentID
    txtStudentName.Text = strStudentName
    txtStudentPhone.Text = strStudentPhone
    txtStudentEmail.Text = strStudentEmail
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileHome_Click()
    frmMain.CloseConnection
    Unload Me
    frmMain.Show
End Sub
