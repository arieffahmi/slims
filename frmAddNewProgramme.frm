VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAddNewProgramme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Programme"
   ClientHeight    =   6180
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   6195
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
   ScaleHeight     =   6180
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3000
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
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
      Left            =   3000
      TabIndex        =   1
      Top             =   2160
      Width           =   3015
   End
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
      Left            =   3000
      TabIndex        =   2
      Top             =   3120
      Width           =   3015
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
      TabIndex        =   4
      Top             =   5160
      Width           =   2895
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
      Left            =   3120
      TabIndex        =   5
      Top             =   5160
      Width           =   2895
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
      Left            =   3000
      TabIndex        =   3
      Top             =   4080
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9480
      Top             =   12960
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
      RecordSource    =   "Programme"
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
   Begin VB.Label lblAddNewProgramme 
      Alignment       =   2  'Center
      Caption         =   "Add New Programme"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   5895
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
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   2655
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
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   2655
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
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lblProgrammeFee 
      Caption         =   "Programme Fees:"
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
      TabIndex        =   6
      Top             =   4080
      Width           =   2655
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
Attribute VB_Name = "frmAddNewProgramme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ValidateInputFlag As Boolean

Private Sub cmdClear_Click()
    txtProgrammeCode.Text = ""
    txtProgrammeName.Text = ""
    txtProgrammeSemesters.Text = ""
    txtProgrammeFee.Text = ""
    
End Sub

Private Sub cmdSubmit_Click()

    If ValidateInputFlag Then
        Adodc1.Refresh
        
        With Adodc1.Recordset
        .AddNew
        !ProgrammeCode = txtProgrammeCode.Text
        !ProgrammeName = txtProgrammeName.Text
        !ProgrammeSemesters = txtProgrammeSemesters.Text
        !ProgrammeFee = Val(txtProgrammeFee.Text)
        
        .Update
        End With
        
        MsgBox "New programme added"
        
        cmdClear_Click
    Else
        ValidateInput
    End If
    
End Sub


Private Sub Form_Load()
    Adodc1.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\data\SLIMS.mdb"
    Adodc1.Refresh
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileHome_Click()
    Unload Me
    frmMain.Show
    
End Sub

Private Sub ValidateInput()

        If txtProgrammeCode.Text = "" Then
            ValidateInputFlag = False
        ElseIf txtProgrammeName.Text = "" Then
            ValidateInputFlag = False
        ElseIf txtProgrammeFee.Text = "" Then
            ValidateInputFlag = False
        ElseIf txtProgrammeSemesters.Text = "" Then
            ValidateInputFlag = False
        Else
            ValidateInputFlag = True
        End If
    
        If ValidateInputFlag = False Then
            MsgBox ("Please fill in the required details.")
        End If
        
End Sub
