VERSION 5.00
Begin VB.Form frmUpdateProgrammeData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Programme Data"
   ClientHeight    =   5865
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   5880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProgrammeSemesters 
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
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3720
      Width           =   4095
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
      Left            =   2880
      TabIndex        =   4
      Top             =   4800
      Width           =   2775
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
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox txtProgrammeFee 
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
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2760
      Width           =   4095
   End
   Begin VB.TextBox txtProgrammeName 
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
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox txtProgrammeCode 
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
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label lblUpdateProgrammeData 
      Alignment       =   2  'Center
      Caption         =   "Update Programme Data"
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
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lblProgrammeSemesters 
      Caption         =   "Programme Semesters:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblProgrammeFee 
      Caption         =   "Programme Fee:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblProgrammeName 
      Caption         =   "Programme Name:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblProgrammeCode 
      Caption         =   "Programme Code:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1335
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
Attribute VB_Name = "frmUpdateProgrammeData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim strInput As String
    Dim strSQL As String
    
    Public strProgrammeCode As String
    Public strProgrammeName As String
    Public strProgrammeSemesters As String
    Public curProgrammeFee As Currency
    
Private Sub cmdClear_Click()
    txtProgrammeCode.Text = ""
    txtProgrammeName.Text = ""
    txtProgrammeSemesters.Text = ""
    txtProgrammeFee.Text = ""
    
End Sub

Private Sub cmdSubmit_Click()
    
    frmMain.OpenConnection
    
    With frmMain.RS
    !ProgrammeCode = txtProgrammeCode.Text
    !ProgrammeName = txtProgrammeName.Text
    !ProgrammeSemesters = txtProgrammeSemesters.Text
    !ProgrammeFee = Val(txtProgrammeFee.Text)
    
    .Update
    End With
    
    MsgBox "Programme data updated"
    
    cmdClear_Click
    mnuFileHome_Click
    
End Sub

Private Sub Form_Load()
    strSQL = "select * from Programme where ProgrammeCode = '" & strInput & "'"
    
    frmMain.OpenConnection
    
    txtProgrammeCode.Text = strProgrammeCode
    txtProgrammeName.Text = strProgrammeName
    txtProgrammeSemesters.Text = strProgrammeSemesters
    txtProgrammeFee.Text = curProgrammeFee
    
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileHome_Click()
    frmMain.CloseConnection
    Unload Me
    frmMain.Show
End Sub
