VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 

Private Sub Command1_Click()

    Winsock1.Close

    Winsock1.Protocol = sckUDPProtocol

    Winsock1.RemoteHost = "255.255.255.255"

    Winsock1.RemotePort = 420

    Winsock1.SendData "lvmsg" & Me.Text1.Text

 

End Sub

 

Private Sub Form_Load()

     Winsock1.Bind 420

End Sub

 

Private Sub Form_Unload(Cancel As Integer)

    Winsock1.Close

End Sub
