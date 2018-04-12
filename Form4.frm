VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FF8080&
   Caption         =   "Bill"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12495
   LinkTopic       =   "Form4"
   ScaleHeight     =   8385
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
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
      Left            =   11160
      TabIndex        =   21
      Top             =   120
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   615
      Left            =   5760
      TabIndex        =   20
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox gtotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   525
      Left            =   10320
      TabIndex        =   17
      Top             =   7560
      Width           =   1455
   End
   Begin VB.TextBox cfee 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   525
      Left            =   10320
      TabIndex        =   15
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox stax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   555
      Left            =   10320
      TabIndex        =   13
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox total 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   525
      Left            =   10320
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   5160
      X2              =   7200
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   480
      X2              =   11760
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   11760
      X2              =   11760
      Y1              =   2040
      Y2              =   5520
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   8520
      X2              =   8520
      Y1              =   2040
      Y2              =   5520
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   1560
      X2              =   1560
      Y1              =   2040
      Y2              =   5520
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   11760
      X2              =   480
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   480
      X2              =   11760
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label bktime 
      BackColor       =   &H00FF8080&
      Caption         =   "Booking Time :"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6840
      TabIndex        =   19
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label sdet 
      BackColor       =   &H00FF8080&
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
      Height          =   1935
      Left            =   600
      TabIndex        =   18
      Top             =   5880
      Width           =   4815
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Grand Total:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Convenience Fee:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7200
      TabIndex        =   14
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Service Tax @ 3.5%:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   6360
      Width           =   3015
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label charges 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2055
      Left            =   9120
      TabIndex        =   9
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label pasname 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   1680
      TabIndex        =   8
      Top             =   3000
      Width           =   6495
   End
   Begin VB.Label sno 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   720
      TabIndex        =   7
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "Charges"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9600
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Passenger Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "S.No."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label traveldt 
      BackColor       =   &H00FF8080&
      Caption         =   "Travel Date :"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label bookdt 
      BackColor       =   &H00FF8080&
      Caption         =   "Booking Date :"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Label packdet 
      BackColor       =   &H00FF8080&
      Caption         =   "Package Details : "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "INVOICE"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim cn1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim sql1 As String
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql As String

Private Sub Command1_Click()
CommonDialog1.ShowPrinter
Command1.Visible = False
Me.PrintForm
Command1.Visible = True
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
Form3.Hide
i = 1
sdet.Caption = "Support Details: " + vbCrLf + "Tour Booking System" + vbCrLf + "XYZ Area" + vbCrLf + "Phone No.: +91-9874987666"
cn.Open ("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + App.Path + "\Passenger.mdb;")
cn1.Open ("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + App.Path + "\Package.mdb;")
sql = "select * from passenger where ID =" + Str(Form3.q)
Set rs = cn.Execute(sql)
sql1 = "select price from Package where destination='" + rs(10) + "' and nightsdays='" + rs(11) + "'"
Set rs1 = cn1.Execute(sql1)
packdet.Caption = packdet.Caption + " " + rs(10) + " (" + rs(11) + ")"
bookdt.Caption = bookdt.Caption + " " + Str(DateValue(Now))
bktime.Caption = bktime.Caption + " " + Str(TimeValue(Now))
traveldt.Caption = traveldt.Caption + " " + Str(rs(9))
While rs.EOF = False
sno.Caption = sno.Caption + vbCrLf + Str(i) + "."
pasname.Caption = pasname.Caption + vbCrLf + rs(1)
charges.Caption = charges.Caption + vbCrLf + Str(rs1(0))
rs.MoveNext
i = i + 1
Wend
i = i - 1
total.Text = rs1(0) * i
stax.Text = (rs1(0) / 100) * 3.5
cfee.Text = 200 * i
gtotal.Text = Val(total.Text) + Val(stax.Text) + Val(cfee.Text)
End Sub

Private Sub Form_Terminate()
cn.Close
cn1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
cn1.Close
End Sub

