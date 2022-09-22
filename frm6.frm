VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm6 
   Caption         =   "Sales"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm6.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Hrishikesh Patil\OneDrive\Desktop\Automobile Automation System\Database\AutomobileDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   14400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Sales"
      Top             =   9480
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm6.frx":3CBBC
      Height          =   5175
      Left            =   -240
      OleObjectBlob   =   "frm6.frx":3CBD0
      TabIndex        =   21
      Top             =   5520
      Width           =   14295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000002&
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   4440
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
         Left            =   12600
         TabIndex        =   23
         Top             =   120
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
         Left            =   11040
         TabIndex        =   22
         Top             =   120
         Width           =   1455
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
         Top             =   120
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
         Height          =   555
         Left            =   7440
         TabIndex        =   19
         Top             =   120
         Width           =   1695
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
         Left            =   5760
         TabIndex        =   18
         Top             =   120
         Width           =   1575
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
         Left            =   3840
         TabIndex        =   17
         Top             =   120
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
         Height          =   615
         Left            =   1800
         TabIndex        =   16
         Top             =   120
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
         Height          =   675
         Left            =   0
         MaskColor       =   &H00FF0000&
         TabIndex        =   15
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Save Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   14055
      Begin VB.TextBox Text6 
         DataField       =   "VehicleType"
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
         Left            =   8520
         TabIndex        =   13
         Top             =   2280
         Width           =   2535
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
         Height          =   525
         Left            =   8520
         TabIndex        =   12
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         DataField       =   "SalePrice"
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
         Left            =   8520
         TabIndex        =   11
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         DataField       =   "SaleDate"
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
         Top             =   2160
         Width           =   2775
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
         Height          =   525
         Left            =   2400
         TabIndex        =   6
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         DataField       =   "VehicleID"
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
         Top             =   720
         Width           =   1695
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
         Height          =   615
         Left            =   6120
         TabIndex        =   10
         Top             =   2280
         Width           =   2775
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
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Price"
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
         TabIndex        =   8
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Date"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Model No."
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
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle ID"
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
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SALE DETAILS FORM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "frm6"
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

