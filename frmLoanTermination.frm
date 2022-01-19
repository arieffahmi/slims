VERSION 5.00
Begin VB.Form frmLoanTermination 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Termination"
   ClientHeight    =   12660
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   19380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12660
   ScaleWidth      =   19380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStudentDetails 
      Caption         =   "Student Details"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   41
      Top             =   1320
      Width           =   9495
      Begin VB.TextBox txtStudentName 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1560
         Width           =   7095
      End
      Begin VB.ComboBox comboStudentID 
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2160
         TabIndex        =   2
         Top             =   840
         Width           =   7095
      End
      Begin VB.TextBox txtStudentPhone 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2280
         Width           =   7095
      End
      Begin VB.TextBox txtStudentEmail 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3000
         Width           =   7095
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
         Height          =   495
         Left            =   240
         TabIndex        =   45
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblStudentName 
         Caption         =   "Student Name:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   44
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblStudentPhone 
         Caption         =   "Student Phone:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   43
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lblStudentEmail 
         Caption         =   "Student Email:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   42
         Top             =   3000
         Width           =   1815
      End
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
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   10920
      Width           =   9495
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
      Height          =   1575
      Left            =   9720
      TabIndex        =   4
      Top             =   10920
      Width           =   9495
   End
   Begin VB.Frame fraProgrammeDetails 
      Caption         =   "Programme Details"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   9720
      TabIndex        =   36
      Top             =   1320
      Width           =   9495
      Begin VB.TextBox txtProgrammeCode 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Width           =   6015
      End
      Begin VB.TextBox txtProgrammeName 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1560
         Width           =   6015
      End
      Begin VB.TextBox txtProgrammeFee 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2280
         Width           =   6015
      End
      Begin VB.TextBox txtProgrammeSemesters 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3000
         Width           =   6015
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
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   840
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
         Height          =   495
         Left            =   240
         TabIndex        =   39
         Top             =   1560
         Width           =   2535
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
         Height          =   495
         Left            =   240
         TabIndex        =   38
         Top             =   2280
         Width           =   2535
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
         Height          =   495
         Left            =   240
         TabIndex        =   37
         Top             =   3000
         Width           =   2655
      End
   End
   Begin VB.Frame fraLoanDetails 
      Caption         =   "Loan Details"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   29
      Top             =   5160
      Width           =   9495
      Begin VB.TextBox txtAdditionalLoanAmount 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3000
         Width           =   5055
      End
      Begin VB.TextBox txtProgrammeFeeRemaining 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox txtProgrammeFeeUnderLoan 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5055
      End
      Begin VB.TextBox txtProgrammeFeeUnderLoanPercentage 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox txtTotalLoanAmount 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3720
         Width           =   5055
      End
      Begin VB.TextBox txtLoanAmountPerSemester 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   4920
         Width           =   5055
      End
      Begin VB.Label lblProgrammeFeeUnderLoan 
         Caption         =   "Programme Fee Under Loan:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   35
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label lblProgrammeFeeUnderLoanPercentage 
         Caption         =   "Programme Fee Under Loan (%):"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   34
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label lblProgrammeFeeRemaining 
         Caption         =   "Programme Fee Remaining:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   2280
         Width           =   3495
      End
      Begin VB.Label lblTotalLoanAmount 
         Caption         =   "Total Loan Amount:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label lblAdditionalLoanAmount 
         Caption         =   "Additional Loan Amount:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   31
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label lblLoanAmountPerSemester 
         Caption         =   "Loan Amount Per Semester:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   4920
         Width           =   3375
      End
   End
   Begin VB.Frame fraRepaymentDetails 
      Caption         =   "Repayment Details"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   9720
      TabIndex        =   22
      Top             =   5160
      Width           =   9495
      Begin VB.TextBox txtAdditionalFees 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3000
         Width           =   5055
      End
      Begin VB.TextBox txtLoanProcessingFee 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox txtTotalRepaymentAmount 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3720
         Width           =   5055
      End
      Begin VB.TextBox txtAnnualInterestRate 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5055
      End
      Begin VB.TextBox txtRepaymentPeriod 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox txtAnnualRepaymentAmount 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   4800
         Width           =   5055
      End
      Begin VB.Label lblAnnualRepaymentAmount 
         Caption         =   "Annual Repayment Amount:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   4920
         Width           =   3135
      End
      Begin VB.Label lblRepaymentPeriod 
         Caption         =   "Repayment Period (Years):"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lblAnnualInterestRate 
         Caption         =   "Annual Interest Rate (%):"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label lblTotalRepaymentAmount 
         Caption         =   "Total Repayment Amount:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   3720
         Width           =   3015
      End
      Begin VB.Label lblLoanProcessingFee 
         Caption         =   "Loan Processing Fee:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label lblAdditionalFees 
         Caption         =   "Additional Fees:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   3000
         Width           =   2055
      End
   End
   Begin VB.Label lblLoanApplication 
      Alignment       =   2  'Center
      Caption         =   "Loan Termination"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   46
      Top             =   120
      Width           =   6615
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
Attribute VB_Name = "frmLoanTermination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim CON As New ADODB.Connection
    Dim RS As New ADODB.Recordset
    Dim strSQL As String
    
    Public strStudentID As String
    Public strStudentName As String
    Public strStudentPhone As String
    Public strStudentEmail As String
    
    Dim strProgrammeCode As String
    Dim strProgrammeName As String
    Dim curProgrammeFee As Currency
    Dim strProgrammeSemesters As String
      
    Dim dblProgrammeFeeUnderLoanPercentage As Double
    Dim dblProgrammeFeeUnderLoan As Currency
    Dim curProgrammeFeeRemaining As Currency
    Dim curAdditionalLoanAmount As Currency
    Dim curTotalLoanAmount As Currency
    
    Dim curLoanAmountPerSemester As Currency
    
    Dim intRepaymentPeriod As Integer
    Dim dblAnnualInterestRate As Double
    Dim curLoanProcessingFee As Currency
    Dim curAdditionalFees As Currency
    Dim curtotalRepaymentAmount As Currency
    Dim curAnnualRepaymentAmount As Currency
    
    Dim clearFlag As Boolean
    Dim ValidateStudentIDFlag As Boolean
    
Private Sub cmdClear_Click()
    clearFlag = True

    comboStudentID.Text = ""
    txtProgrammeCode.Text = ""
    
    txtStudentName.Text = ""
    txtStudentPhone.Text = ""
    txtStudentEmail.Text = ""
    
    txtProgrammeName.Text = ""
    txtProgrammeFee.Text = ""
    txtProgrammeSemesters.Text = ""
    
    txtProgrammeFeeUnderLoanPercentage.Text = ""
    txtProgrammeFeeUnderLoan.Text = ""
    txtProgrammeFeeRemaining.Text = ""
    txtAdditionalLoanAmount.Text = ""
    txtTotalLoanAmount.Text = ""
    
    txtLoanAmountPerSemester.Text = ""
    
    txtRepaymentPeriod.Text = ""
    txtAnnualInterestRate.Text = ""
    txtLoanProcessingFee.Text = ""
    txtAdditionalFees.Text = ""
    txtTotalRepaymentAmount.Text = ""
    
    txtAnnualRepaymentAmount.Text = ""
End Sub

Private Sub cmdSubmit_Click()

    ValidateStudentID

    If ValidateStudentIDFlag Then
        strSQL = "select * from Loan where studentID = '" & comboStudentID.Text & "'"
        
        OpenConnection
        
        If RS.EOF Then
            MsgBox "Student not found"
        Else
            
            RS.MoveFirst
            
            confirmDelete = MsgBox("Confirm loan termination?", vbYesNo)
            
            If confirmDelete = vbYes Then
            
                Dim dblCGPA As Double
                Dim CGPAFlag As Boolean
                
                dblCGPA = InputBox("Enter CGPA:")
                
                If IsNumeric(dblCGPA) Then
                    If (dblCGPA = 4) Then
                        CGPAFlag = True
                        curtotalRepaymentAmount = 0
                    ElseIf (dblCGPA < 4) And (dblCGPA >= 3.5) Then
                        CGPAFlag = True
                        curtotalRepaymentAmount = curtotalRepaymentAmount * 0.25
                    ElseIf (dblCGPA < 3.5) And (dblCGPA >= 3) Then
                        CGPAFlag = True
                        curtotalRepaymentAmount = curtotalRepaymentAmount * 0.5
                    ElseIf (dblCGPA < 3) And (dblCGPA >= 0) Then
                        CGPAFlag = True
                        curtotalRepaymentAmount = curtotalRepaymentAmount
                    Else
                        CGPAFlag = False
                        MsgBox ("Invalid Input")
                    End If
                End If
    
                If CGPAFlag Then
                    
                    RS.Delete
                    
                    MsgBox ("Loan termination successful." & vbNewLine & vbNewLine & "Total Repayment Amount: " & FormatCurrency(curtotalRepaymentAmount))
                End If
                
            End If
        End If
        
        CloseConnection
        
        Unload Me
        frmMain.Show
    End If
    
End Sub

Private Sub comboStudentID_Click()

    strSQL = "select * from Student where studentID = '" & comboStudentID.Text & "'"
    
    OpenConnection
    
    With RS
        .MoveFirst
        
        txtStudentName.Text = !studentName
        txtStudentPhone.Text = !studentPhone
        txtStudentEmail.Text = !studentEmail
    End With
    
    CloseConnection
    
    strSQL = "select * from Loan where studentID = '" & comboStudentID.Text & "'"
    
    OpenConnection
    
    With RS
        .MoveFirst
        
        strProgrammeCode = !ProgrammeCode
        txtProgrammeCode.Text = strProgrammeCode
        
        dblProgrammeFeeUnderLoanPercentage = Val(!ProgrammeFeeUnderLoanPercentage)
        curAdditionalLoanAmount = Val(!AdditionalLoanAmount)
        curTotalLoanAmount = Val(!TotalLoanAmount)
        
        intRepaymentPeriod = Val(!RepaymentPeriod)
        dblAnnualInterestRate = Val(!AnnualInterestRate)
        curAdditionalFees = Val(!AdditionalFees)
        curtotalRepaymentAmount = Val(!TotalRepaymentAmount)
        
    End With
    
    CloseConnection
    
    strSQL = "select * from Programme where ProgrammeCode = '" & txtProgrammeCode.Text & "'"
    
    OpenConnection
    
    With RS
        .MoveFirst
        
        txtProgrammeName.Text = !ProgrammeName
        txtProgrammeFee.Text = FormatCurrency(!ProgrammeFee)
        txtProgrammeSemesters.Text = !ProgrammeSemesters
        
        curProgrammeFee = !ProgrammeFee
    End With
    
    CloseConnection
    
    Calculate
    
End Sub

Private Sub Form_Load()
    LoadComboStudentID
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileHome_Click()
    Unload Me
    frmMain.Show
End Sub


Private Sub LoadComboStudentID()
    strSQL = "select * from Loan"
    
    OpenConnection
    
    If RS.EOF = False Then
        With RS
            .MoveFirst
            Do Until .EOF
                comboStudentID.AddItem !studentID
                .MoveNext
            Loop
        End With
    End If
    
    CloseConnection
    
End Sub

Private Sub CloseConnection()
    If RS.State And adStateOpen Then
        RS.Close
        CON.Close
    End If
End Sub

Private Sub OpenConnection()
    If RS.State = False Then
        CON.Open ("provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\data\SLIMS.mdb")
        RS.Open strSQL, CON, adOpenDynamic, adLockOptimistic
    End If
End Sub

Private Sub Calculate()

    ValidateStudentID
    
    dblProgrammeFeeUnderLoan = (curProgrammeFee / 100) * dblProgrammeFeeUnderLoanPercentage
    curProgrammeFeeRemaining = curProgrammeFee - dblProgrammeFeeUnderLoan
       
    curLoanAmountPerSemester = curTotalLoanAmount / Val(txtProgrammeSemesters.Text)
    
    
    curLoanProcessingFee = (curTotalLoanAmount * intRepaymentPeriod) / 100 * dblAnnualInterestRate
       
    curAnnualRepaymentAmount = curtotalRepaymentAmount / intRepaymentPeriod
    
    txtProgrammeFeeUnderLoanPercentage.Text = dblProgrammeFeeUnderLoanPercentage
    txtProgrammeFeeUnderLoan.Text = FormatCurrency(dblProgrammeFeeUnderLoan)
    txtProgrammeFeeRemaining.Text = FormatCurrency(curProgrammeFeeRemaining)
    txtAdditionalLoanAmount.Text = FormatCurrency(curAdditionalLoanAmount)
    txtTotalLoanAmount.Text = FormatCurrency(curTotalLoanAmount)
    
    txtLoanAmountPerSemester.Text = FormatCurrency(curLoanAmountPerSemester)
    
    txtRepaymentPeriod.Text = intRepaymentPeriod
    txtAnnualInterestRate.Text = dblAnnualInterestRate
    txtLoanProcessingFee.Text = FormatCurrency(curLoanProcessingFee)
    txtAdditionalFees.Text = FormatCurrency(curAdditionalFees)
    txtTotalRepaymentAmount.Text = FormatCurrency(curtotalRepaymentAmount)
    
    txtAnnualRepaymentAmount.Text = FormatCurrency(curAnnualRepaymentAmount)
    
End Sub

Private Sub ValidateStudentID()
    If comboStudentID.Text = "" Then
        ValidateStudentIDFlag = False
    Else
        ValidateStudentIDFlag = True
    End If
    
    If ValidateStudentIDFlag = False Then
        MsgBox ("Please enter Student ID")
    End If
End Sub
