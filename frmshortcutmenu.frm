VERSION 5.00
Begin VB.Form frmshortcutmenu 
   BackColor       =   &H00808080&
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   13635
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcustret 
      Height          =   1815
      Left            =   11640
      Picture         =   "frmshortcutmenu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Customer Return"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   1695
      Left            =   9960
      Picture         =   "frmshortcutmenu.frx":0FEE
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Supplier Delivery"
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdcustomermaintenance 
      Height          =   1455
      Left            =   10440
      Picture         =   "frmshortcutmenu.frx":1B56
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Customer Maintenance"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton cmdproduct 
      Height          =   1455
      Left            =   5760
      Picture         =   "frmshortcutmenu.frx":2846
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Product Maintenance"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdsupplierdelivery 
      Height          =   1695
      Left            =   1440
      Picture         =   "frmshortcutmenu.frx":4015
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Supplier Delivery"
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdsuppliermaintenance 
      Height          =   1695
      Left            =   2280
      Picture         =   "frmshortcutmenu.frx":4A69
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Supplier Maintenance"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Record Customer Return transaction"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11640
      TabIndex        =   11
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FFFFFF&
      X1              =   12240
      X2              =   12360
      Y1              =   4440
      Y2              =   4320
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FFFFFF&
      X1              =   12240
      X2              =   12120
      Y1              =   4440
      Y2              =   4320
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FFFFFF&
      X1              =   10440
      X2              =   10560
      Y1              =   4440
      Y2              =   4320
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FFFFFF&
      X1              =   10320
      X2              =   10440
      Y1              =   4320
      Y2              =   4440
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FFFFFF&
      X1              =   3960
      X2              =   3840
      Y1              =   4320
      Y2              =   4440
   End
   Begin VB.Line Line17 
      BorderColor     =   &H80000009&
      X1              =   3720
      X2              =   3840
      Y1              =   4320
      Y2              =   4440
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00FFFFFF&
      X1              =   1920
      X2              =   2040
      Y1              =   4320
      Y2              =   4440
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FFFFFF&
      X1              =   2160
      X2              =   2040
      Y1              =   4320
      Y2              =   4440
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FFFFFF&
      X1              =   7800
      X2              =   7920
      Y1              =   1800
      Y2              =   1920
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000009&
      X1              =   7920
      X2              =   7800
      Y1              =   1680
      Y2              =   1800
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFFFF&
      X1              =   5520
      X2              =   5640
      Y1              =   1560
      Y2              =   1680
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      X1              =   5640
      X2              =   5520
      Y1              =   1680
      Y2              =   1800
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Record Customer Delivery transaction"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9840
      TabIndex        =   9
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      BorderStyle     =   3  'Dot
      X1              =   3840
      X2              =   3840
      Y1              =   3480
      Y2              =   4440
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000009&
      BorderStyle     =   3  'Dot
      X1              =   2040
      X2              =   2040
      Y1              =   3480
      Y2              =   4440
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      X1              =   2040
      X2              =   3840
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      BorderStyle     =   3  'Dot
      X1              =   12240
      X2              =   12240
      Y1              =   3480
      Y2              =   4440
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   10440
      Y1              =   3480
      Y2              =   4440
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      BorderStyle     =   3  'Dot
      X1              =   11280
      X2              =   11280
      Y1              =   2520
      Y2              =   3480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   12240
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Add new customer, edit, and update information"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      X1              =   7800
      X2              =   10440
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Add new product, edit, and update information"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      X1              =   3480
      X2              =   5640
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Record Supplier Delivery transaction"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Add new supplier, edit and update information"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      X1              =   2880
      X2              =   2880
      Y1              =   2760
      Y2              =   3480
   End
End
Attribute VB_Name = "frmshortcutmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcustomermaintenance_Click()
frmcustomermaintenance.Show
End Sub

Private Sub cmdcustret_Click()
frmcustomerreturn.Show
End Sub

Private Sub cmdproduct_Click()
frmproductmaintenance.Show
End Sub

Private Sub cmdsupplierdelivery_Click()
frmsupplierdelivery.Show
End Sub

Private Sub cmdsuppliermaintenance_Click()
frmsuppliermaintenance.Show
End Sub

Private Sub Command1_Click()
frmincomingdeliveries.Show
End Sub

