VERSION 5.00
Begin VB.Form frmemail 
   Caption         =   "Send E-mail"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5760
   LinkTopic       =   "Form2"
   ScaleHeight     =   4410
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send Mail"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   2295
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   240
   End
End
Attribute VB_Name = "frmemail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim oSmtp As New EASendMailObjLib.Mail
    oSmtp.LicenseCode = "TryIt"
    
    ' Set your sender email address
    oSmtp.FromAddr = "your@yahoo.com"
    
    
    ' Add recipient email address
    oSmtp.AddRecipientEx Text1.Text, 0
    
    ' Set email subject
    oSmtp.Subject = Text2.Text
    
    ' Set email body
    oSmtp.BodyText = Text3.Text
    
    ' Your SMTP server address
    oSmtp.ServerAddr = "smtp.mail.yahoo.com"
    

    
    ' User and password for ESMTP authentication, if your server doesn't require
    ' User authentication, please remove the following codes.
    oSmtp.UserName = "your@yahoo.com"
    oSmtp.Password = "your yahoo pasword"

    ' Set port to 465.
    oSmtp.ServerPort = 465

    ' If your smtp server requires SSL connection, please add this line
    oSmtp.SSL_init
    
    MsgBox "start to send email ..."

    If oSmtp.SendMail() = 0 Then
        MsgBox "email was sent successfully!"
    Else
        MsgBox "failed to send email with the following error:" & oSmtp.GetLastErrDescription()
    End If
End Sub

