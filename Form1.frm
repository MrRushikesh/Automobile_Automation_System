VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm1 
   Caption         =   "Customers"
   ClientHeight    =   9960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   9960
   ScaleWidth      =   20250
   Visible         =   0   'False
   Begin VB.Data Table1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\MyProject\Database\AutomationDatabase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   16440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customer"
      Top             =   6840
      Width           =   1740
   End
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
      Left            =   13920
      TabIndex        =   23
      Top             =   4800
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
      Left            =   12000
      TabIndex        =   22
      Top             =   4800
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":3CBBC
      Height          =   4935
      Left            =   0
      OleObjectBlob   =   "Form1.frx":3CBD0
      TabIndex        =   21
      Top             =   5640
      Width           =   15735
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
      Left            =   9840
      TabIndex        =   19
      Top             =   4800
      Width           =   2055
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
      Left            =   7920
      TabIndex        =   18
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00404080&
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
      Left            =   4080
      MaskColor       =   &H00000000&
      TabIndex        =   17
      Top             =   4800
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
      Left            =   2040
      MaskColor       =   &H00C0C000&
      TabIndex        =   16
      Top             =   4800
      Width           =   1935
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
      Left            =   5880
      MaskColor       =   &H0080FFFF&
      TabIndex        =   15
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
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
      Left            =   0
      MaskColor       =   &H00404040&
      TabIndex        =   14
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Customer Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   15735
      Begin VB.TextBox Text6 
         BackColor       =   &H80000004&
         DataField       =   "VehicleName"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   11160
         TabIndex        =   13
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000004&
         DataField       =   "VehicleType"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   11160
         TabIndex        =   12
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000004&
         DataField       =   "CustomerContact"
         DataSource      =   "Data1"
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
         Left            =   11160
         TabIndex        =   11
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000004&
         DataField       =   "CustomerAddress"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   3000
         TabIndex        =   7
         Top             =   2400
         Width           =   5055
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000004&
         DataField       =   "CustomerName"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000004&
         DataField       =   "CustomerID"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   10
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Type"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         TabIndex        =   9
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Contact"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   8760
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
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
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000002&
      Height          =   855
      Left            =   -120
      TabIndex        =   20
      Top             =   4680
      Width           =   15855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      Caption         =   "CUSTOMER ENQUIRY FORM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15735
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
table1.Recordset.AddNew
MsgBox "Insert Data"
Text1.SetFocus

End Sub

Private Sub Command2_Click()
table1.Recordset.MovePrevious
MsgBox "Saved"

End Sub

Private Sub Command3_Click()
table1.Recordset.Update
MsgBox "Update"
End Sub

Private Sub Command4_Click()
table1.Recordset.MoveNext
MsgBox "Move Next"
End Sub

Private Sub Command5_Click()
table1.Recordset.Delete
MsgBox "Delete"
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Command7_Click()
DataReport1.Show
End Sub

Private Sub Command8_Click()
Data1.Recordset.Edit
MsgBox "Succesful"
End Sub

