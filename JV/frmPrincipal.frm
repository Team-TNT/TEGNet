VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPrincipal 
   Caption         =   "Jugador Virtual"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10575
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   10575
   Visible         =   0   'False
   Begin MSComctlLib.ImageList imgLstSysTrayOK 
      Left            =   615
      Top             =   5745
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":02A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0402
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":055E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":06BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0816
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtMision 
      Height          =   1170
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "frmPrincipal.frx":0972
      Top             =   4320
      Width           =   1620
   End
   Begin VB.TextBox txtErrores 
      Height          =   1170
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "frmPrincipal.frx":0978
      Top             =   2880
      Width           =   1620
   End
   Begin VB.Timer tmr_TimeOut 
      Left            =   600
      Top             =   240
   End
   Begin VB.TextBox txtEnviado 
      Height          =   345
      Left            =   4110
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6495
      Width           =   3705
   End
   Begin VB.TextBox txtRecibido 
      Height          =   345
      Left            =   4125
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6090
      Width           =   3705
   End
   Begin MSWinsockLib.Winsock wskVP 
      Left            =   165
      Top             =   6090
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgLstSysTrayError 
      Left            =   1140
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":097E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0ADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0D92
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":104A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLstSysTrayActivo 
      Left            =   1890
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":11A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1302
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":145E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":15BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1716
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1872
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label12 
      Caption         =   "Actitud:"
      Height          =   270
      Left            =   8025
      TabIndex        =   22
      Top             =   6435
      Width           =   840
   End
   Begin VB.Label Label11 
      Caption         =   "Agresividad:"
      Height          =   270
      Left            =   8010
      TabIndex        =   21
      Top             =   6060
      Width           =   870
   End
   Begin VB.Label lblActitud 
      Caption         =   "Label10"
      Height          =   225
      Left            =   9030
      TabIndex        =   20
      Top             =   6465
      Width           =   930
   End
   Begin VB.Label lblAgresividad 
      Caption         =   "Label9"
      Height          =   270
      Left            =   9030
      TabIndex        =   19
      Top             =   6045
      Width           =   885
   End
   Begin VB.Image imgSysTray 
      Height          =   420
      Left            =   45
      Top             =   5520
      Width           =   420
   End
   Begin VB.Label Label8 
      Caption         =   "A. Norte"
      Height          =   240
      Left            =   75
      TabIndex        =   16
      Top             =   1035
      Width           =   720
   End
   Begin VB.Label Label7 
      Caption         =   "A. Sur"
      Height          =   240
      Left            =   75
      TabIndex        =   15
      Top             =   1425
      Width           =   720
   End
   Begin VB.Label Label6 
      Caption         =   "Asia"
      Height          =   240
      Left            =   75
      TabIndex        =   14
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label5 
      Caption         =   "Europa"
      Height          =   240
      Left            =   75
      TabIndex        =   13
      Top             =   2190
      Width           =   720
   End
   Begin VB.Label Label4 
      Caption         =   "Oceania"
      Height          =   240
      Left            =   75
      TabIndex        =   12
      Top             =   2565
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   "Africa"
      Height          =   240
      Left            =   75
      TabIndex        =   11
      Top             =   645
      Width           =   720
   End
   Begin VB.Label lblPorcentajeContinente 
      Caption         =   "Label3"
      Height          =   210
      Index           =   6
      Left            =   855
      TabIndex        =   10
      Top             =   2580
      Width           =   660
   End
   Begin VB.Label lblPorcentajeContinente 
      Caption         =   "Label3"
      Height          =   210
      Index           =   5
      Left            =   855
      TabIndex        =   9
      Top             =   2196
      Width           =   660
   End
   Begin VB.Label lblPorcentajeContinente 
      Caption         =   "Label3"
      Height          =   210
      Index           =   4
      Left            =   855
      TabIndex        =   8
      Top             =   1812
      Width           =   660
   End
   Begin VB.Label lblPorcentajeContinente 
      Caption         =   "Label3"
      Height          =   210
      Index           =   3
      Left            =   855
      TabIndex        =   7
      Top             =   1428
      Width           =   660
   End
   Begin VB.Label lblPorcentajeContinente 
      Caption         =   "Label3"
      Height          =   210
      Index           =   2
      Left            =   855
      TabIndex        =   6
      Top             =   1044
      Width           =   660
   End
   Begin VB.Label lblPorcentajeContinente 
      Caption         =   "Label3"
      Height          =   210
      Index           =   1
      Left            =   855
      TabIndex        =   5
      Top             =   660
      Width           =   660
   End
   Begin VB.Label lblEstado 
      Caption         =   "Label3"
      Height          =   255
      Left            =   30
      TabIndex        =   4
      Top             =   6540
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Recibido"
      Height          =   255
      Left            =   3225
      TabIndex        =   3
      Top             =   6150
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Enviado"
      Height          =   195
      Left            =   3225
      TabIndex        =   2
      Top             =   6525
      Width           =   900
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "PopUp"
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const MODO_DEBUG = 0

Private Sub Form_Unload(Cancel As Integer)
    
    '###Borrar Mensaje
    #If MODO_DEBUG = 1 Then
        Close #2
        Close #1
    #End If
    
    SysTrayUnload Me
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandle
    
    SysTrayMouseMove Me, Button, Shift, X, Y

    Exit Sub
ErrorHandle:
    ReportErr "Form_MouseMove", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuSalir_Click()
    SysTrayExit Me
End Sub

Private Sub tmr_TimeOut_Timer()
    On Error GoTo ErrorHandle
    
    GintSegRestantesTimeOut = GintSegRestantesTimeOut - 1
    
    If GintSegRestantesTimeOut <= 0 Then
        'Se agoto el tiempo
        tmr_TimeOut.Interval = 0
        EfectuarAccion
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "tmr_TimeOut_Timer", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub wskVP_Close()
    On Error GoTo ErrorHandle
    
    cDesconectar
    
    Unload Me
    
    Exit Sub
ErrorHandle:
    ReportErr "wskVP_Close", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub wskVP_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ErrorHandle
    
    Dim strMensaje As String
    
    wskVP.GetData strMensaje
    
    '###
    txtRecibido.Text = strMensaje
    
    SepararMensajes strMensaje
    
    Exit Sub
ErrorHandle:
    ReportErr "Winsock1_DataArrival", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    Dim vecParametros() As Variant
    
    'Obtiene los parametros de la linea de comando
    vecParametros = ObtenerLineaComando(6)
    
    GintPuerto = vecParametros(1)
    GstrServer = vecParametros(2)
    GstrJvNombre = vecParametros(3)
    GintJvColor = vecParametros(4)
    GsngUmbralAtaque = (vecParametros(5) / 100) * CsngMaxValorPais
    GsngUmbralAtaqueOriginal = GsngUmbralAtaque
    GsngActitud = vecParametros(6) / 100
    
    'Inicializa el SysTray
    imgSysTray.Picture = imgLstSysTrayOK.ListImages(GintJvColor).Picture
    SysTrayInicializar Me.hwnd, GstrJvNombre, imgSysTray
    'Cambia el icono del formulario
    frmPrincipal.Icon = imgLstSysTrayOK.ListImages(GintJvColor).Picture
    frmPrincipal.Caption = GstrJvNombre
    
    lblAgresividad.Caption = GsngUmbralAtaque
    lblActitud.Caption = GsngActitud
    
    '###
    Hardcodear
    
    CargarMatrizEstados
    DescripcionEstadosCliente
    GEstadoCliente = estDesconectado
    lblEstado.Caption = GvecEstadoCliente(GEstadoCliente)
    
    '###Borrar Mensaje
    #If MODO_DEBUG = 1 Then
        Open App.Path & "\mensaje_VP_Salida.txt" For Output As #2
        Open App.Path & "\mensaje_VP_Entrada.txt" For Output As #1
    #End If
    
    'Se conecta con el servidor
    cConectar GstrServer, GintPuerto
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Load", Me.Name, Err.Description, Err.Number, Err.Source
    'Si no logra conectarse al servidor muere
    End
End Sub

