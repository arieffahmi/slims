VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAddNewStudent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register New Student"
   ClientHeight    =   6105
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2880
      TabIndex        =   6
      Top             =   3960
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11280
      Top             =   3240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "Student"
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
      TabIndex        =   9
      Top             =   5040
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
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   2655
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
      Left            =   2880
      TabIndex        =   5
      Top             =   3000
      Width           =   2895
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
      Left            =   2880
      TabIndex        =   4
      Top             =   2040
      Width           =   2895
   End
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
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
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
      TabIndex        =   10
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label lblAddNewStudent 
      Alignment       =   2  'Center
      Caption         =   "Add New Student"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   5535
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
      TabIndex        =   2
      Top             =   3000
      Width           =   2535
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
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
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
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
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
Attribute VB_Name = "frmAddNewStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ValidateInputFlag As Boolean

Private Sub cmdClear_Click()
    txtStudentID.Text = ""
    txtStudentName.Text = ""
    txtStudentPhone.Text = ""
    txtStudentEmail.Text = ""
    
End Sub

Private Sub cmdSubmit_Click()

    If ValidateInputFlag Then
        Adodc1.Refresh
        
        With Adodc1.Recordset
        .AddNew
        !studentID = txtStudentID.Text
        !studentName = txtStudentName.Text
        !studentPhone = txtStudentPhone.Text
        !studentEmail = txtStudentEmail.Text
        .Update
        End With
        
        MsgBox "New student added"
        
        cmdClear_Click
    Else
        ValidateInput
    End If
    
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileHome_Click()
    Unload Me
    frmMain.Show
End Sub

Private Sub Form_Load()
    Adodc1.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\data\SLIMS.mdb"
    Adodc1.Refresh
End Sub

Private Sub ValidateInput()

        If txtStudentID.Text = "" Then
            ValidateInputFlag = False
        ElseIf txtStudentName.Text = "" Then
            ValidateInputFlag = False
        ElseIf txtStudentPhone.Text = "" Then
            ValidateInputFlag = False
        ElseIf txtStudentEmail.Text = "" Then
            ValidateInputFlag = False
        Else
            ValidateInputFlag = True
        End If
    
        If ValidateInputFlag = False Then
            MsgBox ("Please fill in the required details.")
        End If
        
End Sub
