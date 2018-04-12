VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FF8080&
   Caption         =   "Travel Details"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12495
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmdback0 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dateofj 
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   4320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   16744576
      CalendarTitleBackColor=   16744576
      CalendarTitleForeColor=   16777215
      CalendarTrailingForeColor=   -2147483645
      Format          =   85983235
      CurrentDate     =   42488
      MaxDate         =   43100
      MinDate         =   42475
   End
   Begin VB.CommandButton Cmdnext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   5535
   End
   Begin VB.TextBox nop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   525
      Left            =   7920
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   4440
      X2              =   8520
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Passenger Details"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   4200
      TabIndex        =   5
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Lbldate 
      BackColor       =   &H00FF8080&
      Caption         =   "Date Of Journey:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label Lblnop 
      BackColor       =   &H00FF8080&
      Caption         =   "Number Of Passengers:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   2760
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdback0_Click()
Form1.Show
Form2.Hide
End Sub

Private Sub Cmdnext_Click()
If nop.Text = "" Then
MsgBox "Give some Input"
ElseIf Val(nop.Text) > 6 Then
MsgBox "Maximum 6 Passenger"
ElseIf (Val(nop.Text) > 0 And Val(nop.Text) <= 6) Then
Form3.Show
End If
End Sub

Private Sub dateofj_Change()
Cmdnext.SetFocus
End Sub

Private Sub Form_Load()
Form1.Hide
dateofj.MinDate = DateValue(Now)
End Sub


