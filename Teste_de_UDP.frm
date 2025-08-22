VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
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
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3000
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   40028
   End
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

    Winsock1.RemoteHost = "192.168.1.255"

    Winsock1.RemotePort = 40000

    Winsock1.SendData "lvmsg" & Me.Text1.Text

 

End Sub

 

Private Sub Form_Load()

     Winsock1.Bind 40000

End Sub

 

Private Sub Form_Unload(Cancel As Integer)

    Winsock1.Close

End Sub

