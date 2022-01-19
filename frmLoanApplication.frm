VERSION 5.00
Begin VB.Form frmLoanApplication 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Application"
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
      TabIndex        =   41
      Top             =   5160
      Width           =   9495
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
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   4800
         Width           =   5055
      End
      Begin VB.TextBox txtRepaymentPeriod 
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
         TabIndex        =   18
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox txtAnnualInterestRate 
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
         TabIndex        =   19
         Top             =   1560
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
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3720
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
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox txtAdditionalFees 
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
         TabIndex        =   20
         Top             =   3000
         Width           =   5055
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
         TabIndex        =   47
         Top             =   3000
         Width           =   2055
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
         TabIndex        =   46
         Top             =   2280
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
         TabIndex        =   45
         Top             =   3720
         Width           =   3015
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
         TabIndex        =   44
         Top             =   1560
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
         TabIndex        =   43
         Top             =   840
         Width           =   3135
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
         TabIndex        =   42
         Top             =   4920
         Width           =   3135
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
      TabIndex        =   34
      Top             =   5160
      Width           =   9495
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
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   4920
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3720
         Width           =   5055
      End
      Begin VB.TextBox txtProgrammeFeeUnderLoanPercentage 
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
         TabIndex        =   16
         Top             =   840
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1560
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox txtAdditionalLoanAmount 
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
         TabIndex        =   17
         Top             =   3000
         Width           =   5055
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
         TabIndex        =   40
         Top             =   4920
         Width           =   3375
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
         TabIndex        =   39
         Top             =   3000
         Width           =   3375
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
         TabIndex        =   38
         Top             =   3720
         Width           =   2895
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
         TabIndex        =   37
         Top             =   2280
         Width           =   3495
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
         TabIndex        =   36
         Top             =   840
         Width           =   3615
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
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
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
      TabIndex        =   21
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
      TabIndex        =   28
      Top             =   1320
      Width           =   9495
      Begin VB.ComboBox comboProgrammeCode 
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
         Left            =   3240
         TabIndex        =   15
         Top             =   840
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3000
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2280
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1560
         Width           =   6015
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
         TabIndex        =   32
         Top             =   3000
         Width           =   2655
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
         TabIndex        =   31
         Top             =   2280
         Width           =   2535
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
         TabIndex        =   30
         Top             =   1560
         Width           =   2535
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
         TabIndex        =   29
         Top             =   840
         Width           =   2655
      End
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
      Left            =   14640
      TabIndex        =   23
      Top             =   10920
      Width           =   4575
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
      Left            =   9720
      TabIndex        =   22
      Top             =   10920
      Width           =   4695
   End
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
      TabIndex        =   24
      Top             =   1320
      Width           =   9495
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
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3000
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2280
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
         TabIndex        =   14
         Top             =   840
         Width           =   7095
      End
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1560
         Width           =   7095
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
         TabIndex        =   33
         Top             =   3000
         Width           =   1815
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
         TabIndex        =   27
         Top             =   2280
         Width           =   1695
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
         TabIndex        =   26
         Top             =   1560
         Width           =   1695
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
         TabIndex        =   25
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Label lblLoanApplication 
      Alignment       =   2  'Center
      Caption         =   "Loan Application"
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
      TabIndex        =   13
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
Attribute VB_Name = "frmLoanApplication"
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
    Dim curProgrammeFeeUnderLoan As Currency
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
    Dim ValidateInputFlag As Boolean
    Dim ValidateStudentIDFlag As Boolean
    Dim ValidateProgrammeCodeFlag As Boolean
    

Private Sub Form_Load()

    LoadComboStudentID
    LoadComboProgrammeCode
    
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

Private Sub LoadComboStudentID()
    strSQL = "select * from Student where not exists (select * from Loan where studentID = '" & comboStudentID.Text & "')"
    
    OpenConnection
    
    With RS
        .MoveFirst
        Do Until .EOF
            comboStudentID.AddItem !studentID
            .MoveNext
        Loop
    End With
    
    CloseConnection

End Sub

Private Sub LoadComboProgrammeCode()
    strSQL = "select * from Programme"
    
    OpenConnection
    
    With RS
        .MoveFirst
        Do Until .EOF
            comboProgrammeCode.AddItem !ProgrammeCode
            .MoveNext
        Loop
    End With
    
    CloseConnection
End Sub

Private Sub comboProgrammeCode_Click()
    strSQL = "select * from Programme where ProgrammeCode = '" & comboProgrammeCode.Text & "'"
    
    OpenConnection
    
    With RS
        .MoveFirst
        
        txtProgrammeName.Text = !ProgrammeName
        txtProgrammeFee.Text = FormatCurrency(!ProgrammeFee)
        txtProgrammeSemesters.Text = !ProgrammeSemesters
        
        curProgrammeFee = !ProgrammeFee
    End With
    
    CloseConnection
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
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileHome_Click()
    Unload Me
    frmMain.Show
End Sub


Private Sub cmdCalculate_Click()

    ValidateInput

    If ValidateInputFlag Then
    
        curProgrammeFeeUnderLoan = (curProgrammeFee / 100) * dblProgrammeFeeUnderLoanPercentage
        curProgrammeFeeRemaining = curProgrammeFee - curProgrammeFeeUnderLoan
        
        curTotalLoanAmount = curProgrammeFeeUnderLoan + curAdditionalLoanAmount
        
        curLoanAmountPerSemester = curTotalLoanAmount / Val(txtProgrammeSemesters.Text)
        
        
        curLoanProcessingFee = (curTotalLoanAmount * intRepaymentPeriod) / 100 * dblAnnualInterestRate
        
        curtotalRepaymentAmount = curTotalLoanAmount + curLoanProcessingFee + curAdditionalFees
        
        curAnnualRepaymentAmount = curtotalRepaymentAmount / intRepaymentPeriod
        
        
        
        txtProgrammeFeeUnderLoan.Text = FormatCurrency(curProgrammeFeeUnderLoan)
        txtProgrammeFeeRemaining.Text = FormatCurrency(curProgrammeFeeRemaining)
        txtTotalLoanAmount.Text = FormatCurrency(curTotalLoanAmount)
        
        txtLoanAmountPerSemester.Text = FormatCurrency(curLoanAmountPerSemester)
        
        txtLoanProcessingFee.Text = FormatCurrency(curLoanProcessingFee)
        txtTotalRepaymentAmount.Text = FormatCurrency(curtotalRepaymentAmount)
        
        txtAnnualRepaymentAmount.Text = FormatCurrency(curAnnualRepaymentAmount)
        
    End If
    
End Sub

Private Sub cmdSubmit_Click()

    ValidateInput
    
    If ValidateInputFlag = True Then
        confirmSubmit = MsgBox("Confirm loan application?", vbYesNo)
    End If
    
    If confirmSubmit = vbYes Then
        cmdCalculate_Click
        
        strSQL = "select * from Loan where studentID = '" & comboStudentID.Text & "'"
        
        OpenConnection
        
        If RS.EOF Then
            With RS
                .AddNew
                
                !studentID = comboStudentID.Text
                !ProgrammeCode = comboProgrammeCode.Text
                
                !ProgrammeFeeUnderLoanPercentage = dblProgrammeFeeUnderLoanPercentage
                !AdditionalLoanAmount = curAdditionalLoanAmount
                !TotalLoanAmount = curTotalLoanAmount
                
                !RepaymentPeriod = intRepaymentPeriod
                !AnnualInterestRate = dblAnnualInterestRate
                !AdditionalFees = curAdditionalFees
                !TotalRepaymentAmount = curtotalRepaymentAmount
                
                .Update
            End With
            
            MsgBox "Loan application successful." & vbNewLine & vbNewLine & "Total Loan Amount: " & FormatCurrency(curTotalLoanAmount) & vbNewLine & "Total Repayment Amount: " & FormatCurrency(curtotalRepaymentAmount)
            SaveToFile
            
            cmdClear_Click
        ElseIf ValidateInputFlag = False Then
            ValidateInput
        Else
            MsgBox "Loan application failed." & vbNewLine & vbNewLine & "This student already has an active loan of " & FormatCurrency(RS.Fields!TotalRepaymentAmount)
        End If
        
        CloseConnection
    End If
End Sub

Private Sub cmdClear_Click()
    clearFlag = True

    comboStudentID.Text = ""
    comboProgrammeCode.Text = ""
    
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

Private Sub SaveToFile()

    Dim strFilePath As String
    strFilePath = App.Path & "\data\loan.txt"
    Open strFilePath For Output As #1
    
    Write #1, "Student ID: " & comboStudentID.Text
    Write #1, "Student Name: " & txtStudentName.Text
    Write #1, "Student Phone: " & txtStudentPhone.Text
    Write #1,
    Write #1, "Programme Code: " & comboProgrammeCode.Text
    Write #1, "Programme Name: " & txtProgrammeName.Text
    Write #1, "Programme Fee: " & txtProgrammeFee.Text
    Write #1, "Programme Semesters: " & txtProgrammeSemesters.Text
    Write #1,
    Write #1, "Total Loan Amount: " & FormatCurrency(curTotalLoanAmount)
    Write #1, "Total Repayment Amount: " & FormatCurrency(curtotalRepaymentAmount)
    
    Close #1
    MsgBox "Loan application details are saved at " & App.Path & "\data\loan.txt"
End Sub

Private Sub txtAdditionalFees_Change()
    If clearFlag = False Then
        If IsNumeric(txtAdditionalFees.Text) Then
            curAdditionalFees = Val(txtAdditionalFees.Text)
        Else
            MsgBox ("Please enter a numerical value only")
        End If
    End If
End Sub

Private Sub txtAdditionalLoanAmount_Change()
    If clearFlag = False Then
        If IsNumeric(txtAdditionalLoanAmount.Text) Then
            curAdditionalLoanAmount = Val(txtAdditionalLoanAmount.Text)
        Else
            MsgBox ("Please enter a numerical value only")
        End If
    End If
End Sub

Private Sub txtAnnualInterestRate_Change()
    If clearFlag = False Then
        If IsNumeric(txtAnnualInterestRate.Text) Then
            If txtAnnualInterestRate.Text >= 0 And txtAnnualInterestRate.Text <= 100 Then
                dblAnnualInterestRate = Val(txtAnnualInterestRate.Text)
            Else
                MsgBox ("Please enter a numerical value from 0-100 only")
            End If
        Else
            MsgBox ("Please enter a numerical value from 0-100 only")
        End If
    End If
End Sub

Private Sub txtProgrammeFeeUnderLoanPercentage_Change()
    If clearFlag = False Then
        If IsNumeric(txtProgrammeFeeUnderLoanPercentage.Text) Then
            If txtProgrammeFeeUnderLoanPercentage.Text >= 0 And txtProgrammeFeeUnderLoanPercentage.Text <= 100 Then
                dblProgrammeFeeUnderLoanPercentage = Val(txtProgrammeFeeUnderLoanPercentage.Text)
            Else
            MsgBox ("Please enter a numerical value from 0-100 only")
            End If
        Else
            MsgBox ("Please enter a numerical value from 0-100 only")
        End If
    End If
End Sub

Private Sub txtRepaymentPeriod_Change()
    If clearFlag = False Then
        If IsNumeric(txtRepaymentPeriod.Text) Then
            intRepaymentPeriod = Val(txtRepaymentPeriod.Text)
        Else
            MsgBox ("Please enter a numerical value only")
        End If
    End If
End Sub

Private Sub ValidateInput()
    
    ValidateStudentID
    ValidateProgrammeCode

    If ValidateStudentIDFlag And ValidateProgrammeCodeFlag Then

        If txtProgrammeFeeUnderLoanPercentage.Text = "" Then
            ValidateInputFlag = False
        ElseIf txtAdditionalLoanAmount.Text = "" Then
            ValidateInputFlag = False
        ElseIf txtRepaymentPeriod.Text = "" Then
            ValidateInputFlag = False
        ElseIf txtAnnualInterestRate.Text = "" Then
            ValidateInputFlag = False
        ElseIf txtAdditionalFees.Text = "" Then
            ValidateInputFlag = False
        Else
            ValidateInputFlag = True
        End If
    
        If ValidateInputFlag = False Then
            MsgBox ("Please fill in the following details: " & vbNewLine & vbNewLine & "Programme Fee Under Loan (%)" & vbNewLine & "Additional Loan Amount" & vbNewLine & "Repayment Period (Years)" & vbNewLine & "Annual Interest Rate (%)" & vbNewLine & "Additional Fees")
        End If
    
    End If
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

Private Sub ValidateProgrammeCode()
    If comboProgrammeCode.Text = "" Then
        ValidateProgrammeCodeFlag = False
    Else
        ValidateProgrammeCodeFlag = True
    End If
    
    If ValidateProgrammeCodeFlag = False Then
        MsgBox ("Please enter Programme Code")
    End If
End Sub
