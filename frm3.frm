VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm3 
   Caption         =   "Cash"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   27.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm3.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   14880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Width           =   2295
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   14115
      TabIndex        =   23
      Top             =   5640
      Width           =   14175
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
      Height          =   630
      Left            =   12480
      TabIndex        =   20
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      TabIndex        =   19
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "REPORT"
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
      Left            =   10680
      TabIndex        =   17
      Top             =   4800
      Width           =   1695
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
      Left            =   9000
      TabIndex        =   16
      Top             =   4800
      Width           =   1575
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
      Left            =   7200
      TabIndex        =   15
      Top             =   4800
      Width           =   1695
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
      TabIndex        =   14
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
      Left            =   3600
      TabIndex        =   13
      Top             =   4800
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
      Left            =   1800
      TabIndex        =   12
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Cash Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   14175
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   8760
         TabIndex        =   22
         Top             =   2520
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111214593
         CurrentDate     =   44682
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   615
         Left            =   3240
         TabIndex        =   21
         Top             =   1680
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111214593
         CurrentDate     =   44682
      End
      Begin VB.TextBox Text5 
         DataField       =   "PaymentAmount"
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
         Left            =   8760
         TabIndex        =   11
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         DataField       =   "PaymentMode"
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
         Left            =   8760
         TabIndex        =   10
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox Text3 
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
         Left            =   3240
         TabIndex        =   6
         Top             =   2520
         Width           =   1935
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
         Left            =   3240
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Amount"
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
         Left            =   6000
         TabIndex        =   9
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Date"
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
         Left            =   6000
         TabIndex        =   8
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Mode"
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
         Left            =   6000
         TabIndex        =   7
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label4 
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
         Left            =   480
         TabIndex        =   4
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1800
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
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   18
      Top             =   4680
      Width           =   14175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "VEHICLE CASH FORM"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rec As New ADODB.Recordset

Private Sub Command1_Click()
rec.AddNew

End Sub

Private Sub Command2_Click()
rec.MovePrevious
End Sub

Private Sub Command3_Click()
rec.Update

End Sub

Private Sub Command4_Click()
rec.MoveNext
End Sub

Private Sub Command5_Click()
rec.Delete

End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Command8_Click()
rec.Update
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
con.Open "Provider = Microsoft.jet.OLED.4.0; Data source=C:\Users\Intel\Desktop\Automobile Automation System\Database\Customer.mdb;"
rec.CursorLocation = adUseClient
rec.Open "select * from Customer", con, adOpenDynamic, Add
Set Text1.DataSource = rec
Text1.DataField = "CustomerID"
Set Text2.DataSource = rec
Text2.DataField = "CustomerName"
Set Text3.DataSource = rec
Text3.DataField = "CustomerAddress"
Set Text4.DataSource = rec
Text4.DataField = "VehicleType"
Set Text5.DataSource = rec
Text5.DataField = "ContactNumber"
Set Text6.DataSource = rec
Text6.DataField = "VehicleName"



End Sub

