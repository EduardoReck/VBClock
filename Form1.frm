VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form Form1 
   Caption         =   "AnalogClckMX"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3540
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
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   236
   Begin VB.CheckBox editAlarm 
      Caption         =   "Edita Alarme"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   855
   End
   Begin VB.CheckBox editTime 
      Caption         =   "Edita Hora"
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   735
   End
   Begin VB.CheckBox opt1 
      Caption         =   "24h"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   1920
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   95
      ImageHeight     =   185
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1709
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C75
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":42EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5AB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":72F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8B3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A422
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BBF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D5AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EE78
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   1080
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   2760
      Picture         =   "Form1.frx":10359
      Stretch         =   -1  'True
      Top             =   120
      Width           =   630
   End
   Begin VB.Image ImgNum 
      Height          =   495
      Index           =   5
      Left            =   2040
      Picture         =   "Form1.frx":16D23
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   345
   End
   Begin VB.Image ImgNum 
      Height          =   495
      Index           =   4
      Left            =   1680
      Picture         =   "Form1.frx":181F4
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   345
   End
   Begin VB.Image ImgNum 
      Height          =   495
      Index           =   7
      Left            =   2760
      Picture         =   "Form1.frx":19750
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   345
   End
   Begin VB.Image ImgNum 
      Height          =   495
      Index           =   2
      Left            =   960
      Picture         =   "Form1.frx":1ACAC
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   345
   End
   Begin VB.Image ImgNum 
      Height          =   495
      Index           =   3
      Left            =   1320
      Picture         =   "Form1.frx":1C17D
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   345
   End
   Begin VB.Image ImgNum 
      Height          =   495
      Index           =   6
      Left            =   2400
      Picture         =   "Form1.frx":1D876
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   345
   End
   Begin VB.Image ImgNum 
      Height          =   495
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":1EF6F
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   345
   End
   Begin VB.Image ImgNum 
      Height          =   495
      Index           =   1
      Left            =   600
      Picture         =   "Form1.frx":204CB
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   345
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      Height          =   195
      Index           =   10
      Left            =   1440
      TabIndex        =   13
      Top             =   360
      Width           =   195
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   195
      Index           =   9
      Left            =   1200
      TabIndex        =   12
      Top             =   600
      Width           =   195
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Height          =   195
      Index           =   8
      Left            =   960
      TabIndex        =   11
      Top             =   960
      Width           =   195
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   195
      Index           =   7
      Left            =   1080
      TabIndex        =   10
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   195
      Index           =   6
      Left            =   1320
      TabIndex        =   9
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   195
      Index           =   5
      Left            =   1800
      TabIndex        =   8
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   195
      Index           =   4
      Left            =   2280
      TabIndex        =   7
      Top             =   1680
      Width           =   195
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   195
      Index           =   3
      Left            =   2400
      TabIndex        =   6
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   195
      Index           =   2
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   195
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   195
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      Height          =   195
      Index           =   11
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   195
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   360
      Shape           =   3  'Circle
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblHoraClck 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HoraRelógio"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   870
   End
   Begin VB.Line linSeg 
      BorderColor     =   &H000000FF&
      X1              =   128
      X2              =   136
      Y1              =   80
      Y2              =   128
   End
   Begin VB.Line linHora 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   128
      X2              =   128
      Y1              =   81
      Y2              =   56
   End
   Begin VB.Line linMin 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   128
      X2              =   80
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Shape Shape1 
      Height          =   1845
      Left            =   840
      Shape           =   3  'Circle
      Top             =   240
      Width           =   2085
   End
   Begin VB.Label lblHoraAtual 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atual"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
   
    Const PI        As Double = 3.14159265
    Const ConvRad   As Double = PI / 180
 
    Public OutroDia    As Date
    Public meuDia      As Date
    
    Public Ehmin       As Double
    Public Ehhoras     As Double
    Public Hora        As Long
    Public Min         As Long
    Public Seg         As Long
    
    Public CentroX     As Long
    Public CentroY     As Long
    
    Public PreFH       As Long
    Public PreFW       As Long
    
    Public PHoraTamanho As Long
    Public PMinTamanho As Long
    Public PSegTamanho As Long
    
    Public AnguloHora  As Long
    Public AnguloMin   As Long
    Public AnguloSeg   As Long
    Public Angu        As Double
    
    Public imgH        As Long
    Public imgW        As Long
    
    Public img1H       As Long
    Public img1W       As Long
    
    Public formH       As Long
    Public formW       As Long
    
    Public toplbl      As Long
    Public lftlbl      As Long
    Public Plbl        As Long
    
    Dim TAOKEYHR       As Boolean
    Dim TAOKEYMIN      As Boolean
    
    Dim flg(11) As Boolean
    
    Public PMHora As Byte
    Public TSeg As Long
    
    Public SegundosTotais As Long
    
Private Sub ResetFlags(flagIndex As Integer)
    Static i As Integer
    For i = 0 To 11
        If i = flagIndex Then
            flg(i) = False
        Else
            flg(i) = True
        End If
    Next i
End Sub
Private Function QualLbl(X As Single, Y As Single) As Integer
    Static i As Integer
    Static apertou As Integer
    apertou = 404
    For i = 0 To 11
        If Y > lblNum(i).Top And Y < (lblNum(i).Top + lblNum(i).Height) And X > lblNum(i).Left And X < (lblNum(i).Left + lblNum(i).Width) Then
            apertou = i
        End If
    Next i
    QualLbl = apertou
End Function

Private Sub Form_Load()
    Static i            As Byte
    Static filealarm    As Long
    Static AHora2       As Long
    Static AMin2        As Long
    Static ASeg2        As Long
    
    Dim styles As Long
    styles = GetWindowLongW(Form1.hWnd, GWL_EXSTYLE)
    styles = styles Or WS_EX_COMPOSITED
    SetWindowLongW Form1.hWnd, GWL_EXSTYLE, styles
    
    
    imgH = ImgNum(1).Height
    imgW = ImgNum(1).Width
    img1H = Image1.Height
    img1W = Image1.Width
    formH = Form1.Height
    formW = Form1.Width
    
    filealarm = FreeFile()
    
    If Len(Dir$("Alarme.txt")) <> 0 Then
           
        Open "Alarme.txt" For Input As #filealarm 'Abre o arquivo do alarme
        
        Input #filealarm, AHora2, AMin2, ASeg2  'Carrega os dados do arquivo nas variáveis do alarme
        
        Form2.AHora = AHora2
        Form2.AMin = AMin2
        Form2.ASeg = AMin2
        
        Close #filealarm
    
    End If
    Form2.txtAlarm.Text = CStr(AHora2) + ":" + CStr(AMin2) + ":" + CStr(ASeg2)
    
    Static filehour    As Long
    Static oldHora     As Long
    Static oldMin      As Long
    Static oldSeg      As Long
    Static oldDia      As Date
    Static dia         As Date
    
    
    If Len(Dir$("Hora.txt")) <> 0 Then
        filehour = FreeFile()
        
        Open "Hora.txt" For Input As #filehour 'Abre o arquivo hora
        
        Input #filehour, oldHora, oldMin, oldSeg, oldDia 'Recebe as variáveis
        meuDia = Now 'Atualiza o horário atual
        dia = meuDia - oldDia 'Compara a diferença do tempo em que o arquivo foi salvo para calcular quanto tempo passou
        Hora = oldHora + Hour(dia) 'Adiciona a diferença para contar o tempo que passou
        Min = oldMin + Minute(dia)
        Seg = oldSeg + Second(dia)
        Close #filehour
    End If
    
    CentroX = Shape1.Left + (Shape1.Width / 2)
    CentroY = Shape1.Top + (Shape1.Height / 2)
    
    If Shape1.Height > Shape1.Width Then
        PHoraTamanho = Shape1.Width / 2 * 0.4
        PMinTamanho = Shape1.Width / 2 * 0.65
        PSegTamanho = Shape1.Width / 2 * 0.75
        Plbl = Shape1.Width / 2 * 0.85
    Else
        PHoraTamanho = Shape1.Height / 2 * 0.4
        PMinTamanho = Shape1.Height / 2 * 0.65
        PSegTamanho = Shape1.Height / 2 * 0.75
        Plbl = Shape1.Height / 2 * 0.85
    End If
    
    linSeg.X1 = CentroX
    linSeg.Y1 = CentroY
    
    linMin.X1 = CentroX
    linMin.Y1 = CentroY
    
    
    linHora.X1 = CentroX
    linHora.Y1 = CentroY
    Const numDgree As Byte = 150
    
    For i = 0 To 11
        lblNum(i).Enabled = False
        lblNum(i).Left = CentroX + (Plbl * Sin((-30 * i + numDgree) * ConvRad)) - (lblNum(i).Width / 2)
        lblNum(i).Top = CentroY + (Plbl * Cos((-30 * i + numDgree) * ConvRad)) - (lblNum(i).Height / 2)
        lblNum(i).FontSize = Shape1.Height / 2 * 0.15
    Next
    PreFW = Form1.Width
    PreFH = Form1.Height
    
    lblHoraAtual.Caption = meuDia 'Mostra a data e horário atual na label Dia
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static apertado As Single
    
    apertado = QualLbl(X, Y)
    Debug.Print CStr(apertado)
    
    If (apertado >= 0 And apertado < 12) And Button = vbLeftButton Then
    
        TAOKEYHR = True
        
    End If
        
         If (apertado >= 0 And apertado < 12) And Button = vbRightButton Then
        
            TAOKEYMIN = True
        
        End If
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static Hip          As Double
    Static Xdis         As Double
    Static Ydis         As Double
    Static Angux        As Double
    Static Anguy        As Double
    Static cosenu       As Double
    Static senu         As Double
    Static quadradu     As Double
    
    If TAOKEYHR Or TAOKEYMIN Then
    
    Xdis = X - CentroX
    Ydis = Y - CentroY
    
    Hip = Math.Sqr((Xdis ^ 2) + (Ydis ^ 2))
  
    cosenu = Ydis / Hip
    senu = Xdis / Hip
    On Error GoTo mantemomsm
    
    Angux = (Atn(senu / Sqr(-senu * senu + 1))) / ConvRad
    
    Anguy = ((Atn(-cosenu / Sqr(-cosenu * cosenu + 1)) + 2 * Atn(1)) / ConvRad)
mantemomsm:
    ' End If
    
        If Angux > 0 Then
            Angu = 180 - Anguy
        Else
            Angu = Anguy + 180
        End If
        
        If TAOKEYHR Then
            Ehhoras = (Angu / 30) + 0.5
        End If
        
        If TAOKEYMIN Then
            Ehmin = (Angu / 6) + 0.5
        End If
    
    
    Debug.Print CStr(Angux) + "        " + CStr(Anguy) + "        " + CStr(Angu) + "        " + CStr(Ehhoras) + "        " + CStr(Ehmin)
            If TAOKEYHR Then
                linHora.X2 = CentroX + (PHoraTamanho * Sin((180 - Angu) * ConvRad))
                linHora.Y2 = CentroY + (PHoraTamanho * Cos((180 - Angu) * ConvRad))
            End If
            If TAOKEYMIN Then
                linMin.X2 = CentroX + (PMinTamanho * Sin((180 - Angu) * ConvRad))
                linMin.Y2 = CentroY + (PMinTamanho * Cos((180 - Angu) * ConvRad))
            End If
     
    
    End If
                
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        If TAOKEYHR Or TAOKEYMIN Then
            If TAOKEYHR Then
                Hora = Fix(Ehhoras)
            End If
            If TAOKEYMIN Then
                Min = Fix(Ehmin)
            End If
        
        
        TAOKEYHR = False
        TAOKEYMIN = False
        Debug.Print "Desligou tudo"
        End If
        
        
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Unload Form2
    
End Sub

Private Sub AdjustElement(ByRef control As Object)

    control.Top = Form1.Height * (control.Top / PreFH)
    control.Left = Form1.Width * (control.Left / PreFW)
    
End Sub

Private Sub AdjustSize(ByRef control As Object)

    control.Height = Form1.Height * (control.Height / PreFH)
    control.Width = Form1.Width * (control.Width / PreFW)
    
End Sub

Private Sub Form_Resize()

    Static i            As Byte

    AdjustElement Shape1
    AdjustSize Shape1
    AdjustElement Shape2
    AdjustSize Shape2
    AdjustElement lblHoraClck
    AdjustElement lblHoraAtual
    AdjustElement opt1
    AdjustElement editAlarm
    AdjustElement editTime
    AdjustElement Image1

     If Form1.Height > Form1.Width Then

        Image1.Width = Form1.Width * (img1W / formW)
        Image1.Height = Form1.Width * (img1H / formH)
        
        Else
        
        Image1.Height = Form1.Height * (img1H / formH)
        Image1.Width = Form1.Height * (img1W / formW)
        
        End If
        
    CentroX = Shape1.Left + (Shape1.Width / 2)
    CentroY = Shape1.Top + (Shape1.Height / 2)
    
    If Shape1.Height > Shape1.Width Then
    
    PHoraTamanho = Shape1.Width / 2 * 0.4
    PMinTamanho = Shape1.Width / 2 * 0.65
    PSegTamanho = Shape1.Width / 2 * 0.75
    Plbl = Shape1.Width / 2 * 0.85
    
    Else
    
    PHoraTamanho = Shape1.Height / 2 * 0.4
    PMinTamanho = Shape1.Height / 2 * 0.65
    PSegTamanho = Shape1.Height / 2 * 0.75
    Plbl = Shape1.Height / 2 * 0.85
    
    End If
    
    linSeg.X1 = CentroX
    linSeg.Y1 = CentroY

    linMin.X1 = CentroX
    linMin.Y1 = CentroY

    linHora.X1 = CentroX
    linHora.Y1 = CentroY
    
    linSeg.X2 = CentroX + (PSegTamanho * Sin(-AnguloSeg * ConvRad))
    linSeg.Y2 = CentroY + (PSegTamanho * Cos(-AnguloSeg * ConvRad))
    
    linMin.X2 = CentroX + (PMinTamanho * Sin(-AnguloMin * ConvRad))
    linMin.Y2 = CentroY + (PMinTamanho * Cos(-AnguloMin * ConvRad))
    
    linHora.X2 = CentroX + (PHoraTamanho * Sin(-AnguloHora * ConvRad))
    linHora.Y2 = CentroY + (PHoraTamanho * Cos(-AnguloHora * ConvRad))
    
    Const numDgree As Byte = 150

    For i = 0 To 11
    
        If Shape1.Height > Shape1.Width Then
        
        
        lblNum(i).FontSize = Shape1.Width / 2 * 0.15
        lblHoraClck.FontSize = Shape1.Width / 2 * 0.15
        lblHoraAtual.FontSize = Shape1.Width / 2 * 0.15
        
        Else
        
        lblNum(i).FontSize = Shape1.Height / 2 * 0.15
        lblHoraClck.FontSize = Shape1.Height / 2 * 0.15
        lblHoraAtual.FontSize = Shape1.Height / 2 * 0.15
        
        End If
    
        lblNum(i).Left = CentroX + (Plbl * Sin((-30 * i + numDgree) * ConvRad)) - (lblNum(i).Width / 2)
        lblNum(i).Top = CentroY + (Plbl * Cos((-30 * i + numDgree) * ConvRad)) - (lblNum(i).Height / 2)
        
    Next
    
    For i = 0 To 7
        If i = 0 Then
    
            AdjustElement ImgNum(i)
    
        Else
        
            ImgNum(i).Top = ImgNum(0).Top
            ImgNum(i).Left = ImgNum(i - 1).Left + (ImgNum(i - 1).Width + 2)
        
        End If
    
        If Form1.Height > Form1.Width Then

            ImgNum(i).Width = Form1.Width * (imgW / formW)
            ImgNum(i).Height = ImgNum(i).Width * (imgH / imgW)
        
        Else
        
            ImgNum(i).Height = Form1.Height * (imgH / formH)
            ImgNum(i).Width = ImgNum(i).Height * (imgW / imgH)
        
        End If
    
    Next
    
    PreFW = Form1.Width
    
    PreFH = Form1.Height
    
End Sub

Private Sub Image1_Click()
 Form2.Visible = True
End Sub

Private Sub ImgNum_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    
    If editTime.Value And Not editAlarm.Value Then
        Select Case index
    
            Case 0, 1
        
                Select Case Button
        
                    Case vbLeftButton
                    
                        Hora = Hora + 1
                        
                    If Hora > 23 Then
                    
                        Hora = 0
                        
                    End If
                        
                    Case vbRightButton
                    
                        Hora = Hora - 1
                        
                    If Hora < 0 Then
                    
                        Hora = 23
                        
                    End If
                    
                End Select
        
            Case 3, 4
        
                Select Case Button
                    Case vbLeftButton
                        Min = Min + 1
                    If Min > 59 Then
                        Min = 0
                    End If
                    Case vbRightButton
                        Min = Min - 1
                    If Min < 0 Then
                        Min = 59
                    End If
                End Select
        
            Case 6, 7
        
                Select Case Button
                    Case vbLeftButton
                        Seg = Seg + 1
                        If Seg > 59 Then
                        Seg = 0
                        End If
                    Case vbRightButton
                        Seg = Seg - 1
                        If Seg < 0 Then
                        Seg = 59
                        End If
                End Select
        End Select
    End If
    
    If editAlarm.Value And Not editTime.Value Then
    
        Select Case index
    
        Case 0, 1
        
        Select Case Button
        
            Case vbLeftButton
               Form2.AHora = Form2.AHora + 1
               If Form2.AHora > 23 Then
               Form2.AHora = 0
               End If
            Case vbRightButton
               Form2.AHora = Form2.AHora - 1
               If Form2.AHora < 0 Then
               Form2.AHora = 23
               End If
        End Select
        
        Case 3, 4
        
         Select Case Button
            Case vbLeftButton
               Form2.AMin = Form2.AMin + 1
               If Form2.AMin > 59 Then
               Form2.AMin = 0
               End If
            Case vbRightButton
               Form2.AMin = Form2.AMin - 1
               If Form2.AMin < 0 Then
               Form2.AMin = 59
               End If
        End Select
        
        Case 6, 7
        
         Select Case Button
            Case vbLeftButton
            Form2.ASeg = Form2.ASeg + 1
               If Form2.ASeg > 59 Then
               Form2.ASeg = 0
               End If
            Case vbRightButton
               Form2.ASeg = Form2.ASeg - 1
               If Form2.ASeg < 0 Then
               Form2.ASeg = 59
               End If
        End Select
        
        End Select
    
    If Form2.AHora >= 10 Then
            Set ImgNum(1).Picture = ImageList1.ListImages(Form2.AHora - 9).Picture
            Set ImgNum(0).Picture = ImageList1.ListImages(2).Picture
        Else
            Set ImgNum(1).Picture = ImageList1.ListImages(Form2.AHora + 1).Picture
            Set ImgNum(0).Picture = ImageList1.ListImages(1).Picture
        End If
        
    Set ImgNum(3).Picture = ImageList1.ListImages((Form2.AMin \ 10) + 1).Picture
    Set ImgNum(4).Picture = ImageList1.ListImages((Form2.AMin Mod 10) + 1).Picture
    
    Set ImgNum(6).Picture = ImageList1.ListImages((Form2.ASeg \ 10) + 1).Picture
    Set ImgNum(7).Picture = ImageList1.ListImages((Form2.ASeg Mod 10) + 1).Picture
    
    Form2.txtAlarm.Text = CStr(Form2.AHora) + ":" + CStr(Form2.AMin) + ":" + CStr(Form2.ASeg)
    
    End If
    
End Sub




Private Sub Timer1_Timer()
    
    
    'Função de tempo, a cada tick aumenta um segundo na variavel Seg
    
    If editAlarm.Value = 0 And editTime.Value = 0 And TAOKEYHR = 0 And TAOKEYMIN = 0 Then
    
        If (Seg >= 59) Then
            Seg = 0
            Min = Min + 1
        Else
            Seg = Seg + 1
        End If
    
        If (Min > 59) Then
        Min = 0
        Hora = Hora + 1
        End If
    
        If Hora > 23 Then
        Hora = 0
        End If
    
    End If
    
    meuDia = Now
    lblHoraAtual.Caption = meuDia
    
    If Hora > 12 Then
    TSeg = ((Hora - 12) * 3600) + (Min * 60) + Seg
    Else
    TSeg = (Hora * 3600) + (Min * 60) + Seg 'Calcula os segundos totais para ajudar na angulação das horas e minutos
    End If
    
    'Calculo de cada um dos angulos
    AnguloHora = (TSeg / 120) + 180
    AnguloMin = ((Min * 60) + Seg) / 10 + 180
    AnguloSeg = (Seg * 6) + 180

    
    lblHoraClck.Caption = CStr(Hora) + ":" + CStr(Min) + ":" + CStr(Seg)
    If editAlarm.Value = 0 Then
    
    If opt1.Value = 1 Then
    
        If Hora >= 10 Then
        
            Set ImgNum(1).Picture = ImageList1.ListImages((Hora Mod 10) + 1).Picture
            Set ImgNum(0).Picture = ImageList1.ListImages(Hora \ 10 + 1).Picture
            
        Else
        
            Set ImgNum(1).Picture = ImageList1.ListImages(Hora + 1).Picture
            Set ImgNum(0).Picture = ImageList1.ListImages(1).Picture
            
        End If
        
     End If
     If opt1.Value = 0 Then
        PMHora = Hora
        If Hora > 12 Then
            PMHora = Hora - 12
        End If
        If PMHora >= 10 Then
            Set ImgNum(1).Picture = ImageList1.ListImages(PMHora - 9).Picture
            Set ImgNum(0).Picture = ImageList1.ListImages(2).Picture
        Else
            Set ImgNum(1).Picture = ImageList1.ListImages(PMHora + 1).Picture
            Set ImgNum(0).Picture = ImageList1.ListImages(1).Picture
        End If
    End If
    
    Set ImgNum(3).Picture = ImageList1.ListImages((Min \ 10) + 1).Picture
    Set ImgNum(4).Picture = ImageList1.ListImages((Min Mod 10) + 1).Picture
    
    Set ImgNum(6).Picture = ImageList1.ListImages((Seg \ 10) + 1).Picture
    Set ImgNum(7).Picture = ImageList1.ListImages((Seg Mod 10) + 1).Picture
    
    End If
    
    If editAlarm.Value = 0 And editTime.Value = 0 And TAOKEYHR = 0 And TAOKEYMIN = 0 Then
    linSeg.X2 = CentroX + (PSegTamanho * Sin(-AnguloSeg * ConvRad))
    linSeg.Y2 = CentroY + (PSegTamanho * Cos(-AnguloSeg * ConvRad))
    
    linMin.X2 = CentroX + (PMinTamanho * Sin(-AnguloMin * ConvRad))
    linMin.Y2 = CentroY + (PMinTamanho * Cos(-AnguloMin * ConvRad))
    
    linHora.X2 = CentroX + (PHoraTamanho * Sin(-AnguloHora * ConvRad))
    linHora.Y2 = CentroY + (PHoraTamanho * Cos(-AnguloHora * ConvRad))
    End If
End Sub
