VERSION 5.00
Begin VB.Form FRMPROGRESSBAR 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   360
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Generating reports......"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      FillColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "FRMPROGRESSBAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

