VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Package"
   ClientHeight    =   7950
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12495
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5775
      Left            =   5760
      TabIndex        =   10
      Top             =   1560
      Width           =   6015
   End
   Begin VB.CommandButton cmdbook 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Book Now!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Confirm Selection"
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   5295
   End
   Begin VB.ComboBox Combopack 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   5400
      Width           =   4095
   End
   Begin VB.ComboBox Combodest 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000E&
      X1              =   6000
      X2              =   11280
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      X1              =   6000
      X2              =   11280
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      X1              =   11280
      X2              =   11280
      Y1              =   2400
      Y2              =   6120
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Tour Details :"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   7200
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   4320
      X2              =   7800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Lbltravel 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Tour && Travels"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4080
      TabIndex        =   11
      Top             =   240
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   6000
      X2              =   6000
      Y1              =   2400
      Y2              =   6120
   End
   Begin VB.Label Label7 
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
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   5520
      Width           =   5175
   End
   Begin VB.Label Label6 
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
      Height          =   495
      Left            =   6360
      TabIndex        =   7
      Top             =   4800
      Width           =   5175
   End
   Begin VB.Label Label5 
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
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   4080
      Width           =   5175
   End
   Begin VB.Label Label4 
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
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   2640
      Width           =   5175
   End
   Begin VB.Label Label3 
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
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   3360
      Width           =   5175
   End
   Begin VB.Label Lblpack 
      BackColor       =   &H00FF8080&
      Caption         =   "Enter the Package:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Label Lbldest 
      BackColor       =   &H00FF8080&
      Caption         =   "Enter the Destination:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2640
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql As String

Private Sub cmdbook_Click()
Form2.Show
End Sub

Private Sub Combodest_Change()
Dim i As Integer
i = Combodest.ListCount - 1
Do Until (i < 0)
Combodest.RemoveItem (i)
i = i - 1
Loop
sql = "select destination from package where destination like '" + Combodest.Text + "%' order by destination"
Set rs = cn.Execute(sql)
If rs.EOF = False Then
Dim chk As String
chk = rs(0)
Combodest.AddItem (rs(0))
Do Until (rs.EOF)
If (rs(0) <> chk) Then
Combodest.AddItem (rs(0))
chk = rs(0)
rs.MoveNext
Else
chk = rs(0)
rs.MoveNext
End If
Loop
End If
Set rs = Nothing
End Sub

Private Sub Combodest_Click()

Frame1.Visible = True
Lblpack.Visible = True
If (Combopack.Visible = True) Then
Combopack.Clear
End If
Combopack.Visible = True
Dim i As Integer
Dim j As Integer
sql = "select * from package where destination = '" + Combodest.Text + "'"
Set rs = cn.Execute(sql)
entries = rs.GetRows
For i = 0 To UBound(entries, 2)
  Combopack.AddItem (entries(2, i))
 Next
Set rs = Nothing
Combopack.SetFocus
End Sub
Private Sub Combopack_Click()
Frame1.Visible = False
cmdbook.Visible = True
Label9.Visible = True
Line1.Visible = True
sql = "select * from package where destination= '" + Combodest.Text + "' and nightsdays= '" + Combopack.Text + "'"
Set rs = cn.Execute(sql)
Label4.Caption = "Package Price:          Rs" + Str(rs(3))
Label3.Caption = "Transport Mode:            " + rs(4)
Label5.Caption = "Meals Included:            " + rs(5)
Label6.Caption = "Sightseeing:                   " + rs(7)
Label7.Caption = "Hotel Rating:                   " + Str(rs(6))
Set rs = Nothing
cmdbook.SetFocus
End Sub


Private Sub Form_Load()
cn.Open ("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + App.Path + "\Package.mdb;")
Lblpack.Visible = False
cmdbook.Visible = False
Combopack.Visible = False
sql = "select destination from package order by destination"
Set rs = cn.Execute(sql)
Dim chk As String
chk = rs(0)
Combodest.AddItem (rs(0))
Do Until (rs.EOF)
If (rs(0) <> chk) Then
Combodest.AddItem (rs(0))
chk = rs(0)
rs.MoveNext
Else
chk = rs(0)
rs.MoveNext
End If
Loop
Set rs = Nothing
End Sub

Private Sub Form_Terminate()
cn.Close
End Sub


