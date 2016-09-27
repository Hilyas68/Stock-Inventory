VERSION 5.00
Begin VB.Form frmsecurity 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LOGIN"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmsecurity.frx":0000
   ScaleHeight     =   2490
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   1440
      TabIndex        =   8
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmsecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Private Sub cmdcancel_Click()
Combo1.Text = ""
txtpassword.Text = ""
End Sub

Private Sub cmdok_Click()

Set rs = conn.Execute("select * from security_tbl where user_name='" & Combo1.Text & "' and password='" & txtpassword.Text & "'")
If Not rs.EOF Then
frmsecurity.Hide
mdimainmenu.Show
mdimainmenu.StatusBar1.Panels(3) = "Currently Login: " + rs!user_name
Else
MsgBox "Access Denied", vbCritical, "Confirmation"
Combo1.Text = ""
txtpassword.Text = ""
End If

End Sub

Private Sub Form_Load()

connstrinG = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
conn.Open connstrinG

Set rs1 = conn.Execute("select user_name from security_tbl")
Do While Not rs1.EOF
Combo1.AddItem (rs1!user_name)
rs1.MoveNext
Loop

End Sub
