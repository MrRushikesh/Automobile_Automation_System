VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm7 
   Caption         =   "Salesman"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm7.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\MyProject\Database\AutomationDatabase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   14280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Salesman"
      Top             =   6840
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm7.frx":3CBBC
      Height          =   6015
      Left            =   0
      OleObjectBlob   =   "frm7.frx":3CBD0
      TabIndex        =   21
      Top             =   6360
      Width           =   14055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000002&
      Height          =   975
      Left            =   0
      TabIndex        =   14
      Top             =   5160
      Width           =   14055
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
         Left            =   12720
         TabIndex        =   23
         Top             =   240
         Width           =   1215
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
         Left            =   10920
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   9240
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7320
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5520
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PREVIOUS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3720
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADDNEW"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Salesman Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   14055
      Begin VB.TextBox Text6 
         DataField       =   "SalesmanAge"
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
         Height          =   1095
         Left            =   9360
         TabIndex        =   13
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox Text5 
         DataField       =   "SalesmanField"
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
         Left            =   9360
         TabIndex        =   12
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         DataField       =   "SalesmanAge"
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
         Left            =   9360
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         DataField       =   "SalesmanContact"
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
         Left            =   3240
         TabIndex        =   7
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         DataField       =   "SalesmanName"
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
         Left            =   3240
         TabIndex        =   6
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         DataField       =   "SalesmanID"
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
         Left            =   3240
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Target"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6960
         TabIndex        =   10
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman Field"
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
         Left            =   6840
         TabIndex        =   9
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman Age"
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
         Left            =   6840
         TabIndex        =   8
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman Contact"
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
         TabIndex        =   4
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman Name"
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
         TabIndex        =   3
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman ID"
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
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SALESMAN DETAILS FORM"
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
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "frm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Text1.SetFocus

End Sub

Private Sub Command2_Click()
Data1.Recordset.MovePrevious

End Sub

Private Sub Command3_Click()
Data1.Recordset.Update

End Sub

Private Sub Command4_Click()
Data1.Recordset.MoveNext

End Sub

Private Sub Command5_Click()
Data1.Recordset.Delete

End Sub

Private Sub Command6_Click()
End

End Sub

