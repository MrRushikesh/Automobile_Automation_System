VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm4 
   Caption         =   "Employees"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm4.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CommandButton Command8 
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13320
      TabIndex        =   34
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "REPORT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11400
      TabIndex        =   33
      Top             =   6720
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm4.frx":3CBBC
      Height          =   4815
      Left            =   0
      OleObjectBlob   =   "frm4.frx":3CBD0
      TabIndex        =   32
      Top             =   7560
      Width           =   15375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Hrishikesh Patil\OneDrive\Desktop\Automobile Automation System\Fake\Employees.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   15960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Table4"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton Command6 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   29
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   28
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   27
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   26
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   25
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADDNEW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Employee Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   15495
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10800
         TabIndex        =   36
         Top             =   4440
         Width           =   3855
      End
      Begin VB.TextBox Text11 
         DataField       =   "Gender"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10800
         TabIndex        =   30
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox Text10 
         DataField       =   "EmpDOB"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10800
         TabIndex        =   23
         Top             =   3600
         Width           =   3735
      End
      Begin VB.TextBox Text9 
         DataField       =   "EmpAddress"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   10800
         TabIndex        =   22
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox Text8 
         DataField       =   "Email"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10800
         TabIndex        =   21
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox Text7 
         DataField       =   "DeptName"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   10800
         TabIndex        =   20
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox Text6 
         DataField       =   "JoinDate"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3120
         TabIndex        =   14
         Top             =   4440
         Width           =   3375
      End
      Begin VB.TextBox Text5 
         DataField       =   "EmpAge"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3120
         TabIndex        =   13
         Top             =   3600
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         DataField       =   "ContactNo"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3120
         TabIndex        =   12
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         DataField       =   "FatherName"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3120
         TabIndex        =   11
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         DataField       =   "EmpName"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3120
         TabIndex        =   10
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         DataField       =   "EmpID"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3120
         TabIndex        =   9
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Salary"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   35
         Top             =   4440
         Width           =   2655
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee DOB"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   19
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Gender"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   18
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   17
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee E-mail"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7080
         TabIndex        =   16
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7080
         TabIndex        =   15
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Join Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Age"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Emp Contact No."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   6
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   480
         TabIndex        =   5
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Father Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000002&
      Height          =   855
      Left            =   0
      TabIndex        =   31
      Top             =   6600
      Width           =   15375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "EMPLOYEE ENTRY FORM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   15495
   End
End
Attribute VB_Name = "frm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
MsgBox "Insert Data"
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Data1.Recordset.MovePrevious
MsgBox "Saved"
End Sub

Private Sub Command3_Click()
Data1.Recordset.Update
MsgBox "Update"
End Sub

Private Sub Command4_Click()
Data1.Recordset.MoveNext
MsgBox "Move Next"
End Sub

Private Sub Command5_Click()
Data1.Recordset.Delete
MsgBox "Delete"
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Command8_Click()
Data1.Recordset.Edit
MsgBox "Successful"
End Sub
