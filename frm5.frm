VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm5 
   Caption         =   "Purchases"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm5.frx":0000
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
      Left            =   11880
      TabIndex        =   27
      Top             =   5880
      Width           =   1815
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
      Left            =   10200
      TabIndex        =   26
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Hrishikesh Patil\OneDrive\Desktop\Automobile Automation System\Database\AutomobileDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   855
      Left            =   15000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Purchase"
      Top             =   11280
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm5.frx":3CBBC
      Height          =   5535
      Left            =   0
      OleObjectBlob   =   "frm5.frx":3CBD0
      TabIndex        =   25
      Top             =   6840
      Width           =   13815
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
      Left            =   8640
      TabIndex        =   23
      Top             =   5880
      Width           =   1455
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
      Left            =   6960
      TabIndex        =   22
      Top             =   5880
      Width           =   1575
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
      Left            =   5400
      TabIndex        =   21
      Top             =   5880
      Width           =   1455
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
      Left            =   1800
      TabIndex        =   20
      Top             =   5880
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
      Left            =   3600
      TabIndex        =   19
      Top             =   5880
      Width           =   1695
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
      Left            =   0
      TabIndex        =   18
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Purchase Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   13815
      Begin VB.TextBox Text8 
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
         Height          =   555
         Left            =   9720
         TabIndex        =   17
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox Text7 
         DataField       =   "PurchaseDate"
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
         Left            =   9720
         TabIndex        =   16
         Top             =   2400
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         DataField       =   "VehicleQuantity"
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
         Left            =   9720
         TabIndex        =   15
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         DataField       =   "VehiclePrice"
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
         Left            =   9720
         TabIndex        =   14
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text4 
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
         Left            =   3120
         TabIndex        =   9
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         DataField       =   "VehicleModelNo"
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
         TabIndex        =   8
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text2 
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
         Height          =   525
         Left            =   3120
         TabIndex        =   7
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         DataField       =   "CustomerID"
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
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label9 
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
         Left            =   6960
         TabIndex        =   13
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Date"
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
         Left            =   6960
         TabIndex        =   12
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label Label7 
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
         Left            =   6960
         TabIndex        =   11
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   6960
         TabIndex        =   10
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label5 
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
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Model No."
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
         TabIndex        =   4
         Top             =   2400
         Width           =   3255
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
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
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
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000002&
      Height          =   855
      Left            =   0
      TabIndex        =   24
      Top             =   5760
      Width           =   13815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "PURCHASE ENTRY FORM"
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
      Top             =   120
      Width           =   13815
   End
End
Attribute VB_Name = "frm5"
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

