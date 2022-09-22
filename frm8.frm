VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm8 
   Caption         =   "Specification"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm8.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\MyProject\Database\AutomationDatabase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   14880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Specification"
      Top             =   8160
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm8.frx":3CBBC
      Height          =   3495
      Left            =   0
      OleObjectBlob   =   "frm8.frx":3CBD0
      TabIndex        =   33
      Top             =   8040
      Width           =   14655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000002&
      Height          =   1095
      Left            =   0
      TabIndex        =   26
      Top             =   6720
      Width           =   14655
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
         TabIndex        =   35
         Top             =   240
         Width           =   1095
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
         TabIndex        =   34
         Top             =   240
         Width           =   1935
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
         Left            =   9360
         TabIndex        =   32
         Top             =   240
         Width           =   1815
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
         Height          =   675
         Left            =   7440
         TabIndex        =   31
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
         Height          =   675
         Left            =   5640
         TabIndex        =   30
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
         Height          =   675
         Left            =   3720
         TabIndex        =   29
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
         Height          =   675
         Left            =   1800
         TabIndex        =   28
         Top             =   240
         Width           =   1815
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
         TabIndex        =   27
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Specification Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   14655
      Begin VB.TextBox Text12 
         DataField       =   "Capacity"
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
         TabIndex        =   25
         Top             =   4680
         Width           =   3135
      End
      Begin VB.TextBox Text11 
         DataField       =   "Weight"
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
         TabIndex        =   24
         Top             =   3840
         Width           =   3135
      End
      Begin VB.TextBox Text10 
         DataField       =   "RearBreak"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   10800
         TabIndex        =   23
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox Text9 
         DataField       =   "FrontBreak"
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
         TabIndex        =   22
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox Text8 
         DataField       =   "RearTyre"
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
         TabIndex        =   21
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox Text7 
         DataField       =   "FrontTyre"
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
         TabIndex        =   20
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         DataField       =   "WheelBase"
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
         TabIndex        =   13
         Top             =   4800
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         DataField       =   "Headlight"
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
         Height          =   405
         Left            =   3240
         TabIndex        =   12
         Top             =   3960
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         DataField       =   "MaxSpeed"
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
         TabIndex        =   11
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         DataField       =   "MaxPower"
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
         TabIndex        =   10
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         DataField       =   "Displacement"
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
         TabIndex        =   9
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         DataField       =   "EngineID"
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
         TabIndex        =   8
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Tank Capacity"
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
         Left            =   7440
         TabIndex        =   19
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Weight"
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
         Left            =   7440
         TabIndex        =   18
         Top             =   3960
         Width           =   3015
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Rear Break"
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
         Left            =   7440
         TabIndex        =   17
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Front Break"
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
         Left            =   7440
         TabIndex        =   16
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Rear Tyers"
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
         Left            =   7440
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Front Tyers"
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
         Left            =   7440
         TabIndex        =   14
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "WheelBase"
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
         TabIndex        =   7
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Headlight"
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
         Left            =   240
         TabIndex        =   6
         Top             =   3960
         Width           =   3495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Speed"
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
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Power"
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
         TabIndex        =   4
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Displacement"
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
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine ID"
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
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SPECIFICATION DETAILS"
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
Attribute VB_Name = "frm8"
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

