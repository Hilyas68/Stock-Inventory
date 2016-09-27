VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtmonthnow 
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtnextyear 
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtmonth 
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtdate2 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtdate1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Date2"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Date1"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
x = txtmonthnow.Text + txtmonth
txtdate1.Text = x
x = Val(txtmonth.Text) + Val(1)
txtnextyear.Text = x
x = txtmonthnow.Text + txtnextyear
txtdate2.Text = x
End Sub

Private Sub Form_Load()
txtdate1.Text = Format(Now, "mm/dd/yyyy")

txtmonth.Text = Format(Now, "yyyy")
txtmonthnow.Text = Format(Now, "mm/dd/")
End Sub
