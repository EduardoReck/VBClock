VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurações"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BttnChange 
      Caption         =   "Troque a Interface"
      Height          =   360
      Left            =   4680
      TabIndex        =   12
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "0"
      DataMember      =   "assa"
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   4800
      List            =   "Form2.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtEegg 
      Appearance      =   0  'Flat
      Height          =   2085
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock sendSock 
      Left            =   3600
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock entrySock 
      Left            =   3480
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "192.168.1.133"
      RemotePort      =   12345
      LocalPort       =   40000
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   3600
      Top             =   1680
   End
   Begin VB.TextBox Entrada1 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton HoraAtual 
      Caption         =   "Clique para atualizar a hora atual"
      Height          =   720
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtAlarm 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton AlarmSave 
      Caption         =   "Salve o horário do alarme"
      Height          =   720
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   990
   End
   Begin VB.CommandButton TimeSave 
      Caption         =   "Salve a hora atual"
      Height          =   705
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   990
   End
   Begin VB.CommandButton LoadHora 
      Caption         =   "Carregue a hora do arquivo"
      Height          =   840
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   990
   End
   Begin VB.CommandButton LoadAlarm 
      Caption         =   "Carregue o Alarme do Arquivo"
      Height          =   840
      Left            =   2400
      TabIndex        =   0
      Top             =   1560
      Width           =   990
   End
   Begin VB.Label lblIp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      Top             =   3000
      Width           =   465
   End
   Begin VB.Label Titulo1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Defina o horário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label lblAlarm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Defina o horário do alarme:"
      Height          =   195
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   1965
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
     Option Explicit
    
    Public AHora       As Long
    Public AMin        As Long
    Public ASeg        As Long
    
    Dim WSAData        As WSAData
    
    Dim enviaSock      As Long
    Dim recebeSock     As Long
    
    Dim recv_addr      As SOCKADDR_IPV4
    Dim teste          As ip_mreq
    Dim pfd            As POLLFD
    
    Dim erro           As Long
    
    Dim pAddresses     As IP_ADAPTER_ADDRESSES
    Dim pCurrAddress   As IP_ADAPTER_ADDRESSES
    
    Dim esseaqui       As String
    
    Dim endereco       As IP_ADAPTER_UNICAST_ADDRESS
    
    Const outBufLen    As Long = 16 * 1024
    
    Dim lenTeste       As Integer
    
    Dim buf_addr(20)   As SOCKADDR_IPV4
    
    Public Alarme      As Date
    Public FLAGalarme  As Boolean
    Public FLAGhora    As Boolean

Private Sub AlarmSave_Click()
    Static filealarm   As Long
    filealarm = FreeFile()
    Open "Alarme.txt" For Output As #filealarm 'Abre o arquivo para salvar o tempo do alarme

    Write #filealarm, AHora, AMin, ASeg 'Escreve o tempo no arquivo

    Close #filealarm 'Fecha o arquivo
End Sub
Private Sub BttnChange_Click()
Static index As Integer

    WS2_32.closesocket (recebeSock)
    
    index = Combo1.ListIndex

    recv_addr.sin_addr = buf_addr(index).sin_addr
    
    recebeSock = WS2_32.socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
    If recebeSock = INVALID_SOCKET Then
    Debug.Print CStr(WS2_32.WSAGetLastError) + "ta errado no socket"
    End If
    
    If (WS2_32.setsockopt(recebeSock, SOL_SOCKET, SO_REUSEADDR, 1, 16)) <> 0 Then
     Debug.Print CStr(WS2_32.WSAGetLastError) + "  ta errado no reuseaddr"
    End If
     
    With recv_addr
            .sin_family = AF_INET
            .sin_port = WS2_32.htons(12405)
    End With


    If (WS2_32.bind(recebeSock, VarPtr(recv_addr), LenB(recv_addr))) <> 0 Then
        erro = WS2_32.WSAGetLastError
        Debug.Print CStr(erro) + "  ta errado no bind"
        txtEegg.Text = "Interface Inválida" + vbCrLf + txtEegg.Text
        
    End If

    With teste
        .imr_interface.S_addr = INADDR_ANY
        .imr_multiaddr.S_addr = &HE0      '&HE0000000
    End With

    lenTeste = LenB(teste)

    If (WS2_32.setsockopt(recebeSock, IPPROTO_IP, IP_ADD_MEMBERSHIP, teste, lenTeste)) < 0 Then
    Debug.Print CStr(WS2_32.WSAGetLastError) + "  ta errado no multicast if"
    End If

End Sub


Private Sub Entrada1_Click()
    Entrada1.Text = " "
    Titulo1.Caption = "Defina o horário:"
End Sub

Private Sub Entrada1_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then 'Após a tecla Return ser pressionada:
   
        If IsDate(Entrada1.Text) = False And Not (Entrada1.Text = "de ler ") Then
            Entrada1.Text = "Data inválida!"
            Exit Sub
        End If
        If Entrada1.Text = "de ler " Then
        txtEegg.Visible = True
        Form2.Width = 7275
        Exit Sub
        End If
        Form1.OutroDia = CDate(Entrada1.Text) 'A entrada do relógio se torna a hora colocada na caixa de texto
        Form1.Hora = Hour(Form1.OutroDia)
        Form1.Min = Minute(Form1.OutroDia)
        Form1.Seg = Second(Form1.OutroDia)
        KeyAscii = 0
    End If
    
    
End Sub


Private Sub Form_Load()
    Dim buffer As Long
    Static i As Integer

    With entrySock
            .Close
            .Protocol = sckUDPProtocol
            .RemoteHost = "192.168.1.255"
            .remotePort = 40000
            .bind 40000
    End With
    With sendSock
            .Close
            .Protocol = sckUDPProtocol
            .RemoteHost = "192.168.1.133"
            .bind 40001
    End With
    lblIp = entrySock.LocalIP

    If (WS2_32.WSAStartup(2.2, WSAData)) <> 0 Then
    Debug.Print CStr(WS2_32.WSAGetLastError) + "ta errado no startup"
    End If

    recebeSock = WS2_32.socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
    If recebeSock = INVALID_SOCKET Then
    Debug.Print CStr(WS2_32.WSAGetLastError) + "ta errado no socket"
    End If
    
     If (WS2_32.setsockopt(recebeSock, SOL_SOCKET, SO_REUSEADDR, 1, 16)) <> 0 Then
     Debug.Print CStr(WS2_32.WSAGetLastError) + "  ta errado no reuseaddr"
     End If
     
    With recv_addr
            '.sin_addr = '"r"  '"224.0.0.0"
            .sin_addr = 0 '&HC0A80185 '192.168.1.133
            .sin_family = AF_INET
            .sin_port = WS2_32.htons(12405)
    End With


    If (WS2_32.bind(recebeSock, VarPtr(recv_addr), LenB(recv_addr))) <> 0 Then
        erro = WS2_32.WSAGetLastError
        Debug.Print CStr(erro) + "  ta errado no bind"
    End If

    With teste
        .imr_interface.S_addr = INADDR_ANY
        .imr_multiaddr.S_addr = &HE0      '&HE0000000
    End With

    lenTeste = LenB(teste)

    If (WS2_32.setsockopt(recebeSock, IPPROTO_IP, IP_ADD_MEMBERSHIP, teste, lenTeste)) < 0 Then
    Debug.Print CStr(WS2_32.WSAGetLastError) + "  ta errado no multicast if"
    End If

    Form2.Width = 4620
    
    buffer = CoTaskMemAlloc(outBufLen)
    
    
    erro = Iphlpapi.GetAdaptersAddresses(AF_INET, GAA_FLAG_INCLUDE_PREFIX, 0, buffer, outBufLen)
    If erro <> 0 Then
        CoTaskMemFree buffer
        Exit Sub
    End If
    
    CopyMemory pAddresses, ByVal buffer, Len(pAddresses)
       
       i = 0
       
    Do
        'iterar
        Debug.Print SysAllocString(pAddresses.FriendlyNamePtr)
        
        
        CopyMemory endereco, ByVal pAddresses.FirstUnicastAddressPtr, LenB(endereco)
        
        esseaqui = String$(16, "a")
        
        WS2_32.InetNtopW AF_INET, ByVal endereco.Address.sockaddrptr + 4, esseaqui, Len(esseaqui)
        
        If InStr(1, esseaqui, vbNullChar) Then
            esseaqui = Left$(esseaqui, InStr(1, esseaqui, vbNullChar) - 1)
        End If
        
        Debug.Print esseaqui
        
        Combo1.AddItem SysAllocString(pAddresses.FriendlyNamePtr), i
        
        CopyMemory buf_addr(i), ByVal endereco.Address.sockaddrptr, LenB(recv_addr)
        
        i = i + 1
        If pAddresses.NextPtr = 0 Then Exit Do
        CopyMemory pAddresses, ByVal pAddresses.NextPtr, Len(pAddresses)
    Loop
    CoTaskMemFree buffer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        If UnloadMode = 0 Then
        Form2.Visible = False
        Cancel = 1
        End If
        
End Sub

Private Sub HoraAtual_Click()

    Form1.meuDia = Now 'Recebe o horário da maquina
    Form1.Hora = Hour(Form1.meuDia)
    Form1.Min = Minute(Form1.meuDia)
    Form1.Seg = Second(Form1.meuDia)
    
End Sub

Private Sub LoadAlarm_Click()
    Static filealarm   As Long
    If Len(Dir$("Alarme.txt")) <> 0 Then
    filealarm = FreeFile()
    
    Open "Alarme.txt" For Input As #filealarm 'Abre o arquivo do alarme
    
    Input #filealarm, AHora, AMin, ASeg  'Carrega os dados do arquivo nas variáveis do alarme
    
    Close #filealarm
    End If
    txtAlarm.Text = CStr(AHora) + ":" + CStr(AMin) + ":" + CStr(ASeg)
End Sub
Private Sub LoadHora_Click()
    Static filehour    As Long
    Static oldHora     As Long
    Static oldMin      As Long
    Static oldSeg      As Long
    Static oldDia      As Date
    Static dia         As Date
    
    If Len(Dir$("Hora.txt")) = 0 Then Exit Sub
    
    filehour = FreeFile()
    
    Open "Hora.txt" For Input As #filehour 'Abre o arquivo hora
    
    Input #filehour, oldHora, oldMin, oldSeg, oldDia 'Recebe as variáveis
    Form1.meuDia = Now 'Atualiza o horário atual
    dia = Form1.meuDia - oldDia 'Compara a diferença do tempo em que o arquivo foi salvo para calcular quanto tempo passou
    Form1.Hora = oldHora + Hour(dia) 'Adiciona a diferença para contar o tempo que passou
    Form1.Min = oldMin + Minute(dia)
    Form1.Seg = oldSeg + Second(dia)
    Close #filehour
    
    'dhaSock.SendData CStr(oldHora) + ":" + CStr(oldMin) + ":" + CStr(oldSeg)
    
    Debug.Print Combo1.ItemData(Combo1.TabIndex)
    
End Sub

Private Sub Timer1_Timer()
    Static i As Byte
    Static FLAG        As Boolean
    Static Hora2        As Long
    Static Min2         As Long
    Static Seg2         As Long
    Static enviei       As SOCKADDR_IPV4
    Dim strOut          As String
    Dim inLen As Long
    
    
    Dim buffer(1023) As Byte
    
    Hora2 = Form1.Hora
    Min2 = Form1.Min
    Seg2 = Form1.Seg
    
    pfd.socket = recebeSock
    pfd.events = 256
    
    erro = WS2_32.WSAPoll(pfd, 1, 0)
    If (erro < 0) Then
        erro = WS2_32.WSAGetLastError
        Debug.Print CStr(erro) + "  ta errado no poll  "
    End If
    If Entrada1.Text = "de ler " And pfd.revents = 256 Then
    
        inLen = LenB(enviei)
        erro = WS2_32.recvfrom(recebeSock, buffer(0), UBound(buffer) + 1, 0, enviei, inLen)
        
        If erro = -1 Then
            erro = WS2_32.WSAGetLastError
            Debug.Print CStr(erro) + "  ta errado no recieve  "
        End If
        If erro > 0 Then
        strOut = Space$(erro)
        erro = MultiByteToWideChar(CP_UTF8, 0, VarPtr(buffer(0)), erro, StrPtr(strOut), Len(strOut))
        strOut = Left$(strOut, erro)
        'strOut = Left$(StrConv(buffer, vbUnicode), erro)
        End If
        txtEegg.Text = strOut + vbCrLf + txtEegg.Text
    'Debug.Print enviei
    End If
    
    
    If FLAGalarme = True Then
    strOut = CStr(AHora) + ":" + CStr(AMin) + ":" + CStr(ASeg)
        With entrySock
            .Close
            .Protocol = sckUDPProtocol
            .RemoteHost = "192.168.1.255"
            .remotePort = 40000
            .bind 40000, "192.168.1.133"
            .SendData strOut
            End With
        FLAGalarme = False
    End If
    
    
    If AHora = Hora2 And AMin = Min2 And ASeg = Seg2 Then
        FLAG = True
        i = 0
        Beep
    End If
    
    If FLAG Then
    Form1.Shape2.Visible = Not Form1.Shape2.Visible
        i = i + 1
        If i = 11 Then
        FLAG = False
        i = 0
        End If
    
    End If

End Sub

Private Sub TimeSave_Click()

    Static filehour    As Long
    filehour = FreeFile()
    Open "Hora.txt" For Output As #filehour
    Form1.meuDia = Now
    Write #filehour, Form1.Hora, Form1.Min, Form1.Seg, Form1.meuDia
    Close #filehour
    
End Sub

Private Sub txtAlarm_Click()
    txtAlarm.Text = " "
End Sub

Private Sub txtAlarm_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
            If IsDate(txtAlarm.Text) = False Then
            txtAlarm.Text = "Data inválida!"
            Exit Sub
        End If
        Alarme = CDate(txtAlarm.Text)
        AHora = Hour(Alarme)
        AMin = Minute(Alarme)
        ASeg = Second(Alarme)
        KeyAscii = 0

    End If
    
End Sub

Private Sub entrySock_DataArrival(ByVal bytesTotal As Long)
    Dim strIn           As String
    Dim TDia            As Date
    Static THora        As Byte
    Static TMin         As Byte
    Static TSeg         As Byte
    Static filehour     As Long
    Static filealarm    As Long
    
    Const dhaIn As String = "A"
    Const shaIn As String = "B"
    Const chaIn As String = "C"
    Const ahaIn As String = "D"
    Const dhcIn As String = "E"
    Const shcIn As String = "F"
    Const chcIn As String = "G"
    Const mhcIn As String = "H"
    Const qheIn As String = "J"
    
    
    entrySock.GetData strIn, vbString
    
    Select Case Left$(strIn, 1)
        
        Case dhaIn
        
            If IsDate(Mid$(strIn, 2)) = False Then
            Entrada1.Text = "Data inválida!"
            Exit Sub
            End If
        
        Form1.OutroDia = CDate(Mid$(strIn, 2)) 'A entrada do relógio se torna a hora colocada na caixa de texto
        Form1.Hora = Hour(Form1.OutroDia)
        Form1.Min = Minute(Form1.OutroDia)
        Form1.Seg = Second(Form1.OutroDia)
        
        Case shaIn
            
            filehour = FreeFile()
            Open "Hora.txt" For Output As #filehour
            Form1.meuDia = Now
            Write #filehour, Form1.Hora, Form1.Min, Form1.Seg, Form1.meuDia
            Close #filehour
            
        Case chaIn
            Static oldHora     As Long
            Static oldMin      As Long
            Static oldSeg      As Long
            Static oldDia      As Date
            Static dia         As Date
    
            If Len(Dir$("Hora.txt")) = 0 Then Exit Sub
    
            filehour = FreeFile()
    
            Open "Hora.txt" For Input As #filehour 'Abre o arquivo hora
    
            Input #filehour, oldHora, oldMin, oldSeg, oldDia 'Recebe as variáveis
            Form1.meuDia = Now 'Atualiza o horário atual
            dia = Form1.meuDia - oldDia 'Compara a diferença do tempo em que o arquivo foi salvo para calcular quanto tempo passou
            Form1.Hora = oldHora + Hour(dia) 'Adiciona a diferença para contar o tempo que passou
            Form1.Min = oldMin + Minute(dia)
            Form1.Seg = oldSeg + Second(dia)
            Close #filehour
            
        Case ahaIn
            
            Form1.meuDia = Now 'Recebe o horário da maquina
            Form1.Hora = Hour(Form1.meuDia)
            Form1.Min = Minute(Form1.meuDia)
            Form1.Seg = Second(Form1.meuDia)
            
        Case dhcIn
        
            If IsDate(Mid$(strIn, 2)) = False Then
            txtAlarm.Text = "Data inválida!"
            Exit Sub
            End If
            
        Alarme = CDate(Mid$(strIn, 2))
        AHora = Hour(Alarme)
        AMin = Minute(Alarme)
        ASeg = Second(Alarme)
        
        txtAlarm.Text = CStr(AHora) + ":" + CStr(AMin) + ":" + CStr(ASeg)
        
        Case shcIn
        
        filealarm = FreeFile()
        Open "Alarme.txt" For Output As #filealarm 'Abre o arquivo para salvar o tempo do alarme

        Write #filealarm, AHora, AMin, ASeg 'Escreve o tempo no arquivo

        Close #filealarm 'Fecha o arquivo
        
        Case chcIn
        
        If Len(Dir$("Alarme.txt")) <> 0 Then
        filealarm = FreeFile()
    
        Open "Alarme.txt" For Input As #filealarm 'Abre o arquivo do alarme
    
        Input #filealarm, AHora, AMin, ASeg  'Carrega os dados do arquivo nas variáveis do alarme
    
        Close #filealarm
        End If
        txtAlarm.Text = CStr(AHora) + ":" + CStr(AMin) + ":" + CStr(ASeg)
        
        Case mhcIn
        
        FLAGalarme = True
        
    End Select
End Sub

