VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm9 
   Caption         =   "Stocks"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm9.frx":0000
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
      Left            =   13200
      TabIndex        =   23
      Top             =   4800
      Width           =   1335
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
      Left            =   11280
      TabIndex        =   22
      Top             =   4800
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm9.frx":3CBBC
      Height          =   4695
      Left            =   0
      OleObjectBlob   =   "frm9.frx":3CBD0
      TabIndex        =   21
      Top             =   5760
      Width           =   14655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\MyProject\Database\AutomationDatabase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   15000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Stock"
      Top             =   6840
      Width           =   2895
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
      Height          =   615
      Left            =   9480
      TabIndex        =   19
      Top             =   4800
      Width           =   1695
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
      Height          =   615
      Left            =   7560
      TabIndex        =   18
      Top             =   4800
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
      Height          =   615
      Left            =   5640
      TabIndex        =   17
      Top             =   4800
      Width           =   1815
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
      Height          =   615
      Left            =   1800
      TabIndex        =   16
      Top             =   4800
      Width           =   1695
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
      Height          =   615
      Left            =   3600
      TabIndex        =   15
      Top             =   4800
      Width           =   1935
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
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Stock Details"
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
      Width           =   14655
      Begin VB.TextBox Text6 
         DataField       =   "VehicleTyre"
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
         Left            =   9240
         TabIndex        =   13
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         DataField       =   "VehicleColor"
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
         Left            =   9240
         TabIndex        =   12
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         DataField       =   "VehicleName"
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
         Left            =   9240
         TabIndex        =   11
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         DataField       =   "Quantity"
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
         Left            =   2400
         TabIndex        =   7
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         DataField       =   "ModelNo"
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
         Left            =   2400
         TabIndex        =   6
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         DataField       =   "StockID"
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
         Left            =   2400
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Type"
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
         Left            =   6120
         TabIndex        =   10
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Color"
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
         Left            =   6120
         TabIndex        =   9
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Name"
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
         Left            =   6120
         TabIndex        =   8
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   360
         TabIndex        =   4
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Model No"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock ID"
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
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000002&
      Height          =   855
      Left            =   0
      TabIndex        =   20
      Top             =   4680
      Width           =   14655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "STOCK DETAILS FORM"
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
      Width           =   14655
   End
End
Attribute VB_Name = "frm9"
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

