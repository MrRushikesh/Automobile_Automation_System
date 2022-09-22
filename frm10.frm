VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm10 
   Caption         =   "Vehicles"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   FillColor       =   &H00404080&
   ForeColor       =   &H80000013&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm10.frx":0000
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
      Height          =   555
      Left            =   12840
      TabIndex        =   28
      Top             =   5160
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
      TabIndex        =   27
      Top             =   5160
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm10.frx":3CBBC
      Height          =   6375
      Left            =   -600
      OleObjectBlob   =   "frm10.frx":3CBD0
      TabIndex        =   26
      Top             =   6000
      Width           =   14895
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\MyProject\Database\AutomationDatabase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   15240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "VehicleEntry"
      Top             =   8640
      Width           =   2055
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
      Left            =   9240
      TabIndex        =   24
      Top             =   5160
      Width           =   1935
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
      TabIndex        =   23
      Top             =   5160
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
      Left            =   5640
      TabIndex        =   22
      Top             =   5160
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
      Height          =   615
      Left            =   1680
      TabIndex        =   21
      Top             =   5160
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
      Left            =   3600
      TabIndex        =   20
      Top             =   5160
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
      TabIndex        =   19
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Vehicle Details"
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
      Width           =   14295
      Begin VB.TextBox Text8 
         DataField       =   "VehicleType"
         DataSource      =   "Data1"
         Height          =   525
         Left            =   10200
         TabIndex        =   18
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         DataField       =   "Discount"
         DataSource      =   "Data1"
         Height          =   525
         Left            =   10200
         TabIndex        =   17
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         DataField       =   "VehicleQuantity"
         DataSource      =   "Data1"
         Height          =   525
         Left            =   10200
         TabIndex        =   16
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         DataField       =   "VehiclePrice"
         DataSource      =   "Data1"
         Height          =   525
         Left            =   10200
         TabIndex        =   15
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         DataField       =   "VehicleColor"
         DataSource      =   "Data1"
         Height          =   525
         Left            =   3360
         TabIndex        =   10
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         DataField       =   "VehicleModelNo"
         DataSource      =   "Data1"
         Height          =   525
         Left            =   3360
         TabIndex        =   9
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         DataField       =   "VehicleName"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   3360
         TabIndex        =   8
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         DataField       =   "VehicleID"
         DataSource      =   "Data1"
         Height          =   615
         Left            =   3360
         TabIndex        =   7
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label10 
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
         Height          =   375
         Left            =   6960
         TabIndex        =   14
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Discount"
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
         TabIndex        =   13
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Quantity"
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
         TabIndex        =   12
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Price"
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
         TabIndex        =   11
         Top             =   720
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
         Left            =   480
         TabIndex        =   6
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Model No"
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
         TabIndex        =   5
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   15
         Left            =   360
         TabIndex        =   4
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1560
         Width           =   2295
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
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000002&
      Height          =   855
      Left            =   0
      TabIndex        =   25
      Top             =   5040
      Width           =   14295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "VEHICLE ENTRY FORM"
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
      Width           =   14295
   End
End
Attribute VB_Name = "frm10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Text1.SetFocus
MsgBox "INSERT DATA"



End Sub

Private Sub Command2_Click()
Data1.Recordset.MovePrevious
MsgBox "MOVE PREVIOUS"

End Sub

Private Sub Command3_Click()
Data1.Recordset.Update
MsgBox "SAVED"

End Sub

Private Sub Command4_Click()
Data1.Recordset.MoveNext
MsgBox "MOVE NEXT"


End Sub

Private Sub Command5_Click()
Data1.Recordset.Delete
MsgBox "DELETE"

End Sub

Private Sub Command6_Click()
End

End Sub

Private Sub Command7_Click()
DataReport1.Show

End Sub

Private Sub Command8_Click()
Data1.Recordset.Edit
MsgBox "EDIT"

End Sub

