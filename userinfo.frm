VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF8080&
   Caption         =   "Passenger Details"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12495
   LinkTopic       =   "Form3"
   ScaleHeight     =   8430
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmdback 
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
      TabIndex        =   18
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7080
      Width           =   4335
   End
   Begin VB.TextBox Txtid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   540
      Left            =   8760
      TabIndex        =   16
      Top             =   5520
      Width           =   2895
   End
   Begin VB.ComboBox Comboid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      ItemData        =   "userinfo.frx":0000
      Left            =   2280
      List            =   "userinfo.frx":0002
      TabIndex        =   14
      Top             =   5520
      Width           =   3615
   End
   Begin VB.TextBox Txtemail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   2280
      TabIndex        =   12
      Top             =   4680
      Width           =   9375
   End
   Begin VB.TextBox Txtad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   2280
      TabIndex        =   10
      Top             =   3840
      Width           =   9375
   End
   Begin VB.TextBox Txtcontact 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
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
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   8760
      TabIndex        =   8
      Top             =   3000
      Width           =   2895
   End
   Begin VB.ComboBox Combogen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   540
      Left            =   2280
      TabIndex        =   6
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox Txtage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
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
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   8760
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Txtname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   4560
      X2              =   8040
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "ID No. :"
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
      Left            =   7080
      TabIndex        =   15
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "ID Proof :"
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
      Left            =   480
      TabIndex        =   13
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Email ID :"
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
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Address :"
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
      Left            =   480
      TabIndex        =   9
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "Contact :"
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
      Left            =   6960
      TabIndex        =   7
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Gender :"
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
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Age :"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Name :"
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
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblno 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql As String
Public q As Integer
Dim i As Integer



Private Sub Cmdback_Click()
Form2.Show
Form3.Hide
End Sub



Private Sub Command1_Click()
Cmdback.Visible = False
Dim myAnswer As Integer
If (i = Val(Form2.nop.Text) - 1) Then
Command1.Caption = "Confirm Booking"
End If
If (i < Val(Form2.nop.Text)) Then
sql = "insert into Passenger values('" + Str(q) + "','" + Txtname.Text + "','" + Txtage.Text + "','" + Combogen.Text + "','" + Txtad.Text + "','" + Txtcontact.Text + "','" + Txtemail.Text + "','" + Comboid.Text + "','" + Txtid.Text + "',#" + Str(Form2.dateofj.Value) + "#,'" + Form1.Combodest.Text + "','" + Form1.Combopack.Text + "')"
cn.Execute (sql)
i = i + 1
Txtname.Text = ""
Txtage.Text = ""
Txtcontact.Text = ""
Txtad.Text = ""
Txtemail.Text = ""
Txtid.Text = ""
Combogen.Text = ""
Comboid.Text = ""
lblno.Caption = "Passenger" + Str(i)
Txtname.SetFocus
Else
sql = "insert into Passenger values('" + Str(q) + "','" + Txtname.Text + "','" + Txtage.Text + "','" + Combogen.Text + "','" + Txtad.Text + "','" + Txtcontact.Text + "','" + Txtemail.Text + "','" + Comboid.Text + "','" + Txtid.Text + "',#" + Str(Form2.dateofj.Value) + "#,'" + Form1.Combodest.Text + "','" + Form1.Combopack.Text + "')"
cn.Execute (sql)
cn.Close
myAnswer = MsgBox("Are You Sure?", vbYesNo + vbQuestion + vbDefaultButton2)
If (myAnswer = vbYes) Then
Form4.Show
End If
End If
End Sub

Private Sub Form_Load()
Form2.Hide
Combogen.AddItem ("Male")
Combogen.AddItem ("Female")
Comboid.AddItem ("Voter ID Card")
Comboid.AddItem ("Aadhar Card")
Comboid.AddItem ("Passport")
Comboid.AddItem ("Pancard")
i = 1
If (i = Val(Form2.nop.Text)) Then
Command1.Caption = "Confirm Booking"
Else
Command1.Caption = "NEXT"
End If
lblno.Caption = "Passenger" + Str(i)
cn.Open ("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + App.Path + "\Passenger.mdb;")
sql = "select id from passenger"
Set rs = cn.Execute(sql)
If (rs.EOF = True) Then
q = 1
Else
While (rs.EOF = False)
q = rs(0) + 1
rs.MoveNext
Wend
End If
End Sub

Private Sub Form_Terminate()
cn.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
cn.Close
End Sub






