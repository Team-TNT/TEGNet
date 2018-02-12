VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdifrmPrincipal 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "TEGNet"
   ClientHeight    =   3660
   ClientLeft      =   1665
   ClientTop       =   3000
   ClientWidth     =   6270
   Icon            =   "mdifrmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrSysTray 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3240
      Top             =   960
   End
   Begin MSComctlLib.ImageList imgLstFichas 
      Left            =   840
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":05F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":095E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":106A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":13EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":176E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgLstBotones"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Conectar/Desconectar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Resincronizar Partida"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Opciones"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Guardar partida"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pausar Juego"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mostrar Misión"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mostrar Tarjetas"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar Tropas"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atacar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mover Tropa"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Tomar Tarjeta"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fin Turno"
            ImageIndex      =   11
         EndProperty
      EndProperty
      Begin VB.Image imgSysTrayTurno 
         Height          =   240
         Left            =   5115
         Picture         =   "mdifrmPrincipal.frx":1BC2
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgSysTrayOK 
         Height          =   240
         Left            =   6285
         Picture         =   "mdifrmPrincipal.frx":1D0C
         Top             =   60
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList imgLstBotones 
      Left            =   1560
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":1E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":22AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":25C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":2722
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":2A3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":2B9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":2CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":2E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":316E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":348A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":35E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":3A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":3B96
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":3CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmPrincipal.frx":3E4E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   2520
      Top             =   960
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3405
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   529
            MinWidth        =   529
            Object.ToolTipText     =   "Intercambio de información con el Servidor"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   794
            MinWidth        =   794
            Object.ToolTipText     =   "Mi color"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   150
            MinWidth        =   18
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2672
            Object.ToolTipText     =   "Tipo de Ronda"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2672
            Object.ToolTipText     =   "Estado del Juego"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Object.ToolTipText     =   "Administración"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   477
            MinWidth        =   477
            Picture         =   "mdifrmPrincipal.frx":3FAA
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   882
            MinWidth        =   882
            Object.ToolTipText     =   "Tiempo del turno"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuPartida 
      Caption         =   "&Partida"
      Begin VB.Menu mnuPartidaConectar 
         Caption         =   "&Conectar..."
      End
      Begin VB.Menu mnuPartidaDesconectar 
         Caption         =   "&Desconectar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPartidaResincronizar 
         Caption         =   "&Resincronizar"
      End
      Begin VB.Menu mnuPartidaGuardar 
         Caption         =   "&Guardar..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuPartidaPausar 
         Caption         =   "&Pausar"
      End
      Begin VB.Menu MnuSepPartida1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPartidaIdioma 
         Caption         =   "&Idioma..."
      End
      Begin VB.Menu MnuOpciones 
         Caption         =   "&Opciones..."
      End
      Begin VB.Menu MnuSepPartida2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPartidaSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuVerMapa 
         Caption         =   "&Mapa"
      End
      Begin VB.Menu mnuVerJugadores 
         Caption         =   "&Jugadores"
      End
      Begin VB.Menu mnuVerChat 
         Caption         =   "&Chat"
      End
      Begin VB.Menu mnuVerSeleccion 
         Caption         =   "&Selección"
      End
      Begin VB.Menu mnuVerDados 
         Caption         =   "&Dados"
      End
      Begin VB.Menu mnuVerMision 
         Caption         =   "Mi&sión"
      End
      Begin VB.Menu mnuVerTarjetas 
         Caption         =   "&Tarjetas"
      End
      Begin VB.Menu mnuVerTropasDisponibles 
         Caption         =   "T&ropas Disponibles"
      End
      Begin VB.Menu mnuVerInfo 
         Caption         =   "&Información"
      End
      Begin VB.Menu mnuVerLog 
         Caption         =   "&Bitácora"
      End
      Begin VB.Menu mnuVerListaMisiones 
         Caption         =   "&Lista de Misiones"
      End
   End
   Begin VB.Menu mnuJuego 
      Caption         =   "&Juego"
      Begin VB.Menu mnuJuegoAgregar 
         Caption         =   "Agregar una Tro&pa"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuJuegoAtacar 
         Caption         =   "&Atacar"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuJuegoMover 
         Caption         =   "&Mover una Tropa"
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuSepJuego1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJuegoTomar 
         Caption         =   "&Tomar Tarjeta"
         Shortcut        =   ^T
      End
      Begin VB.Menu MnuSepJuego2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJuegoFinalizar 
         Caption         =   "&Finalizar Turno"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "&Ventana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuOrganizar 
         Caption         =   "&Organizar Ventanas"
      End
   End
   Begin VB.Menu mnuAccion 
      Caption         =   "A&ccion"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuAtacar 
         Caption         =   "&Atacar País"
      End
      Begin VB.Menu MnuSepAccion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMover 
         Caption         =   "&Mover una Tropa"
      End
      Begin VB.Menu mnuMoverTodas 
         Caption         =   "Mover &Todas las Tropas"
      End
   End
   Begin VB.Menu MnuAdministracion 
      Caption         =   "&Administración"
      Begin VB.Menu MnuCambiarAdministrador 
         Caption         =   "&Cambiar Administrador"
      End
      Begin VB.Menu mnuAsignarJV 
         Caption         =   "&Asignar Jugador Virtual"
      End
      Begin VB.Menu MnuSepAdm1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBajarServidor 
         Caption         =   "&Bajar Servidor"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu MnuSepAyuda1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "Acerca de..."
      End
   End
   Begin VB.Menu mnuAgregar 
      Caption         =   "Agregar Tropa"
      Visible         =   0   'False
      Begin VB.Menu mnuAgregar1 
         Caption         =   "Agregar 1 tropa"
      End
      Begin VB.Menu mnuAgregar5 
         Caption         =   "Agregar 5 tropas"
      End
      Begin VB.Menu mnuAgregar10 
         Caption         =   "Agregar 10 tropas"
      End
      Begin VB.Menu MnuSepAgregar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAgregarTodas 
         Caption         =   "Agregar todas las tropas"
      End
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "PopUp"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "mdifrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Regional - Variables de error
Private strMsgAbandonar As String
Private strMsgAbandonarCaption As String
Private strMsgDesconectar As String
Private strMsgDesconectarCaption As String
Private strMsgWinsockClose As String
Private strMsgWinsockCloseCaption As String
Private strMsgConexion1 As String
Private strMsgConexion1Caption As String
Private strMsgConexion2 As String
Private strMsgConexion2Caption As String
Private strMsgConexionOtro As String
Private strMsgConexionOtroCaption As String
Private strMsgBajarServidor As String
Private strMsgBajarServidorCaption As String

Public Sub CargaInicial()
    On Error GoTo ErrorHandle
    
    InicializarJugadores
    
    GsoyAdministrador = False
    MnuAdministracion.Visible = False
    
    DescripcionEstadosCliente
    CargarMatrizEstados
    CargarMatrizControles
    '###E
    
    'Setea el estado inicial como desconectado
    GEstadoCliente = estDesconectado
    
    'Setea las opciones de administracion
    SetearAdm False
    
    'Para que no rompa las bolas (tiene que cargarse primero)
    Load frmMapa
    
    ActualizarControles
    
    'Conecta con la base de datos
'    conectarBaseDatos
    
    'Setea los colores
    GvecColores(1) = vbBlack
    GvecColores(2) = vbMagenta
    GvecColores(3) = vbRed
    GvecColores(4) = vbBlue
    GvecColores(5) = vbYellow
    GvecColores(6) = vbGreen
    
    GvecColoresInv(1) = vbWhite
    GvecColoresInv(2) = vbWhite
    GvecColoresInv(3) = vbWhite
    GvecColoresInv(4) = vbWhite
    GvecColoresInv(5) = vbBlack
    GvecColoresInv(6) = vbBlack

    cargarInfoMapa
    
    CargarPersonalizacion
    
    'Muestra los formularios
'    frmTarjetas.Visible = False
'    frmPropiedades.Visible = False
'    frmMision.Visible = False
'    frmJugadores.Visible = True
'    frmChat.Visible = True
'    frmSeleccion.Visible = True
'    frmDados.Visible = False
    
    ActualizarMenu
    
    Exit Sub
ErrorHandle:
    ReportErr "CargaInicial", Me.Name, Err.Description, Err.Number, Err.Source

End Sub

Private Sub MDIForm_Load()
    
    On Error GoTo ErrorHandle
    'frmInicio.Show 'vbModal
    
    'Regional - Carga
    Me.Caption = ObtenerTextoRecurso(CintPrincipalCaption)
    mnuPartida.Caption = ObtenerTextoRecurso(CintPrincipalMnuPartida)
    mnuVer.Caption = ObtenerTextoRecurso(CintPrincipalMnuVer)
    mnuJuego.Caption = ObtenerTextoRecurso(CintPrincipalMnuJuego)
    mnuVentana.Caption = ObtenerTextoRecurso(CintPrincipalMnuVentana)
    MnuAdministracion.Caption = ObtenerTextoRecurso(CintPrincipalMnuAdministracion)
    mnuAyuda.Caption = ObtenerTextoRecurso(CintPrincipalMnuAyuda)
    mnuPartidaConectar.Caption = ObtenerTextoRecurso(CintPrincipalMnuPartidaConectar)
    mnuPartidaDesconectar.Caption = ObtenerTextoRecurso(CintPrincipalMnuPartidaDesconectar)
    mnuPartidaResincronizar.Caption = ObtenerTextoRecurso(CintPrincipalMnuPartidaResincronizar)
    mnuPartidaGuardar.Caption = ObtenerTextoRecurso(CintPrincipalMnuPartidaGuardar)
    mnuPartidaPausar.Caption = ObtenerTextoRecurso(CintPrincipalTipPausar)
    mnuPartidaIdioma.Caption = ObtenerTextoRecurso(CintPrincipalMnuPartidaIdioma)
    MnuOpciones.Caption = ObtenerTextoRecurso(CintPrincipalMnuPartidaOpciones)
    mnuPartidaSalir.Caption = ObtenerTextoRecurso(CintPrincipalMnuPartidaSalir)
    mnuVerMapa.Caption = ObtenerTextoRecurso(CintPrincipalMnuVerMapa)
    mnuVerJugadores.Caption = ObtenerTextoRecurso(CintPrincipalMnuVerJugadores)
    mnuVerChat.Caption = ObtenerTextoRecurso(CintPrincipalMnuVerChat)
    mnuVerSeleccion.Caption = ObtenerTextoRecurso(CintPrincipalMnuVerSeleccion)
    mnuVerDados.Caption = ObtenerTextoRecurso(CintPrincipalMnuVerDados)
    mnuVerMision.Caption = ObtenerTextoRecurso(CintPrincipalMnuVerMision)
    mnuVerTarjetas.Caption = ObtenerTextoRecurso(CintPrincipalMnuVerTarjetas)
    mnuVerTropasDisponibles.Caption = ObtenerTextoRecurso(CintPrincipalMnuVerTropas)
    mnuVerInfo.Caption = ObtenerTextoRecurso(CintPrincipalMnuVerInformacion)
    mnuVerLog.Caption = ObtenerTextoRecurso(CintPrincipalMnuVerLog)
    mnuVerListaMisiones.Caption = ObtenerTextoRecurso(CintPrincipalMnuVerListaMisiones)
    mnuJuegoAgregar.Caption = ObtenerTextoRecurso(CintPrincipalMnuJuegoAgregar)
    mnuJuegoAtacar.Caption = ObtenerTextoRecurso(CintPrincipalMnuJuegoAtacar)
    mnuJuegoMover.Caption = ObtenerTextoRecurso(CintPrincipalMnuJuegoMover)
    mnuJuegoTomar.Caption = ObtenerTextoRecurso(CintPrincipalMnuJuegoTomar)
    mnuJuegoFinalizar.Caption = ObtenerTextoRecurso(CintPrincipalMnuJuegoFinTurno)
    mnuOrganizar.Caption = ObtenerTextoRecurso(CintPrincipalMnuVentanaOrganizar)
    MnuCambiarAdministrador.Caption = ObtenerTextoRecurso(CintPrincipalMnuAdministracionCambiar)
    mnuAsignarJV.Caption = ObtenerTextoRecurso(CintPrincipalMnuAdministracionJV)
    MnuBajarServidor.Caption = ObtenerTextoRecurso(CintPrincipalMnuAdministracionBajar)
    mnuAcercaDe.Caption = ObtenerTextoRecurso(CintPrincipalMnuAyudaAcercaDe)
    mnuAtacar.Caption = ObtenerTextoRecurso(CintPrincipalMnuPopupAtacar)
    mnuMover.Caption = ObtenerTextoRecurso(CintPrincipalMnuPopupMover1)
    mnuMoverTodas.Caption = ObtenerTextoRecurso(CintPrincipalMnuPopupMoverTodas)
    mnuAgregar1.Caption = ObtenerTextoRecurso(CintPrincipalMnuPopupAgregar1)
    mnuAgregar5.Caption = ObtenerTextoRecurso(CintPrincipalMnuPopupAgregar5)
    mnuAgregar10.Caption = ObtenerTextoRecurso(CintPrincipalMnuPopupAgregar10)
    mnuAgregarTodas.Caption = ObtenerTextoRecurso(CintPrincipalMnuPopupAgregarTodas)
    Toolbar1.Buttons(enuToolBar.tbConexion).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipConectar)
    Toolbar1.Buttons(enuToolBar.tbResincronizacion).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipResincronizar)
    Toolbar1.Buttons(enuToolBar.tbOpciones).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipOpciones)
    Toolbar1.Buttons(enuToolBar.tbGuardar).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipGuardar)
    Toolbar1.Buttons(enuToolBar.tbPausa).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipPausar)
    Toolbar1.Buttons(enuToolBar.tbMision).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipVerMision)
    Toolbar1.Buttons(enuToolBar.tbVerTarjetas).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipVerTarjetas)
    Toolbar1.Buttons(enuToolBar.tbAgregar).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipAgregar)
    Toolbar1.Buttons(enuToolBar.tbAtacar).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipAtacar)
    Toolbar1.Buttons(enuToolBar.tbMover).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipMover)
    Toolbar1.Buttons(enuToolBar.tbTomarTarjeta).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipTomar)
    Toolbar1.Buttons(enuToolBar.tbFinTurno).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipFinTurno)
    StatusBar1.Panels(enuPaneles.panIcoIntercambio).ToolTipText = ObtenerTextoRecurso(CintPrincipalStatusIntercambio)
    StatusBar1.Panels(enuPaneles.panIcoColor).ToolTipText = ObtenerTextoRecurso(CintPrincipalStatusMiColor)
    StatusBar1.Panels(enuPaneles.panTipoRonda).ToolTipText = ObtenerTextoRecurso(CintPrincipalStatusTipoRonda)
    StatusBar1.Panels(enuPaneles.panEstado).ToolTipText = ObtenerTextoRecurso(CintPrincipalStatusEstadoJuego)
    StatusBar1.Panels(enuPaneles.panADM).ToolTipText = ObtenerTextoRecurso(CintPrincipalStatusAdministracion)
    StatusBar1.Panels(enuPaneles.panTimer).ToolTipText = ObtenerTextoRecurso(CintPrincipalStatusTiempoTurno)
    strMsgAbandonar = ObtenerTextoRecurso(CintPrincipalMsgAbandonar)
    strMsgAbandonarCaption = ObtenerTextoRecurso(CintPrincipalMsgAbandonarCaption)
    strMsgDesconectar = ObtenerTextoRecurso(CintPrincipalMsgDesconectar)
    strMsgDesconectarCaption = ObtenerTextoRecurso(CintPrincipalMsgDesconectarCaption)
    strMsgWinsockClose = ObtenerTextoRecurso(CintPrincipalMsgServidor)
    strMsgWinsockCloseCaption = ObtenerTextoRecurso(CintPrincipalMsgServidorCaption)
    strMsgConexion1 = ObtenerTextoRecurso(CintPrincipalMsgConexion)
    strMsgConexion1Caption = ObtenerTextoRecurso(CintPrincipalMsgConexionCaption)
    strMsgConexion2 = ObtenerTextoRecurso(CintPrincipalMsgConexion2)
    strMsgConexion2Caption = ObtenerTextoRecurso(CintPrincipalMsgConexion2Caption)
    strMsgConexionOtro = ObtenerTextoRecurso(CintPrincipalMsgOtros)
    strMsgConexionOtroCaption = ObtenerTextoRecurso(CintPrincipalMsgOtrosCaption)
    strMsgBajarServidor = ObtenerTextoRecurso(CintPrincipalMsgBajarServidor)
    strMsgBajarServidorCaption = ObtenerTextoRecurso(CintPrincipalMsgBajarServidorCaption)
    
    'Inicializa el SysTray
    SysTrayInicializar Me.hwnd, "TEGNet", imgSysTrayOK
    
    '###
    GblnSeCierra = False
    
    CargaInicial
    
    'Muestra el Wizard
    mnuPartidaConectar_Click
    
    Exit Sub
ErrorHandle:
    If frmInicio.Visible Then
        Unload frmInicio
    End If
    ReportErr "MDIForm_Load", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorHandle
    Dim blnDescargar As VbMsgBoxResult
    Dim blnCerrarServidor As VbMsgBoxResult

    'Si está desconectado no pregunta
    If GEstadoCliente = estDesconectado Then
        blnDescargar = vbYes
    Else
        blnDescargar = MsgBox(strMsgAbandonar, vbQuestion + vbYesNo + vbDefaultButton2, strMsgAbandonarCaption)
    End If
    
    If blnDescargar = vbYes Then
        
        GrabarPersonalizacion
        
        'Si está conectado y es ADM,
        'pregunta si quiere Bajar el Servidor.
        If GEstadoCliente <> estDesconectado And GEstadoCliente <> estInconsistente Then
            If GsoyAdministrador Then
                blnCerrarServidor = MsgBox(ObtenerTextoRecurso(CintPrincipalMsgCerrarServidor), vbQuestion + vbYesNo + vbDefaultButton2, ObtenerTextoRecurso(CintPrincipalMsgCerrarServidorCaption))
                'Para que no avise que se cerró el Server
                GintServidorCerrado = enuServidorCerrado.secSalida
                If blnCerrarServidor = vbYes Then
                    cBajarServidor
                    'Sin el DoEvents, el msg queda encolado y no sale hacia al Server.
                    DoEvents
                End If
            End If
        End If
        
        '###
        'Cierra todas las ventanas que puedan haber quedado abiertas
        GblnSeCierra = True
        Unload frmSeleccionColor
        Unload frmComienzo
        Unload frmCambioAdm
        Unload frmGuardarPartida
        Unload frmOpciones
        Unload frmConquista
        Unload frmIdioma
    
    Else
        Cancel = 1
    End If

    Exit Sub
ErrorHandle:
    ReportErr "MDIForm_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub


Public Sub cargarInfoMapa()
'Lee de la base de datos y arma la cuadrícula que indica
'a que pais corresponde cada punto

On Error GoTo ErrorHandle

    Dim strSQL As String
    Dim strNombrePais As String
    Dim byPais As Byte
    
    'Obtiene la cantidad de paises
    GbyCantidadPaises = 50
    
    'Por cada pais
    For byPais = 1 To GbyCantidadPaises
        strNombrePais = ObtenerTextoRecurso(enuIndiceArchivoRecurso.pmsPaises + byPais)
        frmMapa.objPais(byPais).Nombre = strNombrePais
        frmMapa.objPais(byPais).ToolTipText = strNombrePais
        frmMapa.objPais(byPais).MostrarFicha = False
    Next
        
    Exit Sub
ErrorHandle:
    ReportErr "cargarInfoMapa", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    
    SysTrayUnload Me
    
    Exit Sub
ErrorHandle:
    ReportErr "MDIForm_Unload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuAcercaDe_Click()
    On Error GoTo ErrorHandle
    
    MostrarFormulario frmCreditos, vbModal
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuAcercaDe_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuAgregar1_Click()
    On Error GoTo ErrorHandle
    
    cAgregarTropas frmMapa.PaisSeleccionadoOrigen, 1
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuAgregar1_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuAgregar10_Click()
        On Error GoTo ErrorHandle
    
    cAgregarTropas frmMapa.PaisSeleccionadoOrigen, 10
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuAgregar1_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuAgregar5_Click()
    On Error GoTo ErrorHandle
    
    cAgregarTropas frmMapa.PaisSeleccionadoOrigen, 5
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuAgregar1_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuAgregarTodas_Click()
    On Error GoTo ErrorHandle
    
    cAgregarTropas frmMapa.PaisSeleccionadoOrigen, GvecJugadores(GintMiColor).intTropasDisponibles
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuAgregar1_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuAsignarJV_Click()
    On Error GoTo ErrorHandle
    
    MostrarFormulario frmJugadorVirtual, vbModal
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuAsignarJV_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuAtacar_Click()
    On Error GoTo ErrorHandle
    
    mnuJuegoAtacar_Click
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuAtacar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub MnuBajarServidor_Click()
    On Error GoTo ErrorHandle
    
    If MsgBox(strMsgBajarServidor, vbYesNo + vbExclamation, strMsgBajarServidorCaption) = vbYes Then
        'Para que muestre mensaje de Servidor cerrado voluntariamente
        GintServidorCerrado = enuServidorCerrado.secVoluntaria
        cBajarServidor
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "MnuBajarServidor_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub MnuCambiarAdministrador_Click()
    On Error GoTo ErrorHandle
    
    MostrarFormulario frmCambioAdm, vbModal
    
    Exit Sub
ErrorHandle:
    ReportErr "MnuCambiarAdministrador_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuJuegoAgregar_Click()
    On Error GoTo ErrorHandle
    
    cAgregarTropas frmMapa.PaisSeleccionadoOrigen, 1
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuJuegoAgregar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuJuegoAtacar_Click()
    On Error GoTo ErrorHandle
    
    cAtacar frmMapa.PaisSeleccionadoOrigen, frmMapa.PaisSeleccionadoDestino
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuJuegoAtacar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuJuegoFinalizar_Click()
    On Error GoTo ErrorHandle
    
    cFinTurno
    
    Exit Sub
ErrorHandle:
    ReportErr "MnuJuegoFinalizar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuJuegoMover_Click()
    On Error GoTo ErrorHandle
    
    cMover frmMapa.PaisSeleccionadoOrigen, frmMapa.PaisSeleccionadoDestino, 1, tmMovimiento
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuJuegoMover_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuJuegoTomar_Click()
    On Error GoTo ErrorHandle

    cTomarTarjeta
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuJuegoTomar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuMover_Click()
    On Error GoTo ErrorHandle

    mnuJuegoMover_Click
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuMover_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuMoverTodas_Click()
    On Error GoTo ErrorHandle
    
    cMoverTodas frmMapa.PaisSeleccionadoOrigen, frmMapa.PaisSeleccionadoDestino
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuMoverTodas_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub MnuOpciones_Click()
    On Error GoTo ErrorHandle
    
    MostrarFormulario frmOpciones, vbModal
    
    Exit Sub
ErrorHandle:
    ReportErr "MnuOpciones_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub
Public Sub mnuOrganizar_Click()
'Es publica porque la llamo desde mdlCliente, en CargarPersonalizacion()
    On Error GoTo ErrorHandle
   
    'Mapa
    frmMapa.Top = 0
    frmMapa.Left = 0
    
    'Jugadores
    frmJugadores.Top = 0
    frmJugadores.Left = frmMapa.Width
      
    'Seleccion
    With frmSeleccion
        'No importa si esta Visible, su Heigth se calcula igual porque se usa como referencia para el Chat y el Log.
        'If .Visible Then
            .Left = frmMapa.Width
            .Top = frmJugadores.Height
            .Width = frmJugadores.Width
            'Si el MDI esta Maximizado, aprovecha la parte inferior al maximo
            'pero si el MDI esta Normal, usa TODO el espacio libre (con un Minimo)
            If mdifrmPrincipal.WindowState = vbMaximized Then
                'Si hay alto de sobra en el MDI, lo aprovecha
                If mdifrmPrincipal.ScaleHeight > (frmJugadores.Height + 930) Then
                    'Pero si el espacio sobrante es demasiado grande solo aprovecho hasta un máximo de 1500
                    If (mdifrmPrincipal.ScaleHeight - frmJugadores.Height) > (840 + (.Height - .ScaleHeight)) Then
                        .Height = 840 + (.Height - .ScaleHeight) 'Maximo (equivalente a 4 lineas del form Bitacora)
                    Else
                        .Height = mdifrmPrincipal.ScaleHeight - frmJugadores.Height
                    End If
                Else
                    .Height = 930 'Minimo
                End If
            ElseIf mdifrmPrincipal.WindowState = vbNormal Then
                If mdifrmPrincipal.ScaleHeight > (frmJugadores.Height + 930) Then
                    'Usa Todo
                    .Height = mdifrmPrincipal.ScaleHeight - frmJugadores.Height
                Else
                    .Height = 930 'Minimo
                End If
            End If
        'End If
    End With
    
    'Chat, luego de Seleccion
    With frmChat
        If .Visible Then
            .Left = 0
            .Height = frmSeleccion.Height
            .Top = frmMapa.Height
            'Si no está visible el form Seleccion,
            'se ocupa toda la parte inferior.
            If frmSeleccion.Visible Then
                .Width = frmMapa.Width
            Else
                .Width = frmMapa.Width + frmJugadores.Width
            End If
        End If
    End With
    
    'Log, luego de Seleccion
    With frmLog
        If .Visible And Not frmChat.Visible Then
            .Left = 0
            .Height = frmSeleccion.Height
            .Top = frmMapa.Height
            'Si no está visible el form Seleccion,
            'se ocupa toda la parte inferior.
            If frmSeleccion.Visible Then
                .Width = frmMapa.Width
            Else
                .Width = frmMapa.Width + frmJugadores.Width
            End If
        End If
    End With
    
    'Tropas Disponibles
    With frmTropasDisponibles
        If .Visible Then
            .Left = 45
            .Top = frmMapa.Height - .Height - 45
        End If
    End With
    
    'Dados
    With frmDados
        If .Visible Then
            .Left = 45
            .Top = frmMapa.Height - .Height - 45
        End If
    End With
    
    'MDI
    'Si no está maximizada, le da el tamaño suficiente
    'para que contenga los formularios MDIChild
    With mdifrmPrincipal
        If .WindowState = 0 Then
            'Por un problema con las Scrollbars...
            'segun el tipo de Ventana, aplica una correccion distinta.
            If (frmJugadores.Height - frmJugadores.ScaleHeight) = 480 Then
                'Ventanas XP
                .Width = frmMapa.Width + frmJugadores.Width + 180
                If frmSeleccion.Visible Then
                    .Height = frmMapa.Height + frmSeleccion.Height + 1550
                ElseIf frmChat.Visible Then
                    .Height = frmMapa.Height + frmChat.Height + 1550
                ElseIf frmLog.Visible Then
                    .Height = frmMapa.Height + frmLog.Height + 1550
                Else
                    .Height = frmMapa.Height + 1550
                End If
            ElseIf (frmJugadores.Height - frmJugadores.ScaleHeight) = 375 Then
                'Ventanas Clasicas
                .Width = frmMapa.Width + frmJugadores.Width + 180
                If frmSeleccion.Visible Then
                    .Height = frmMapa.Height + frmSeleccion.Height + 1430
                ElseIf frmChat.Visible Then
                    .Height = frmMapa.Height + frmChat.Height + 1430
                ElseIf frmLog.Visible Then
                    .Height = frmMapa.Height + frmLog.Height + 1430
                Else
                    .Height = frmMapa.Height + 1430
                End If
            Else
                'Otro (posible Fuentes Grandes...)
                'Se adapta, pero mantiene el problema de las Scrollbars
                .Width = frmMapa.Width + frmJugadores.Width + (.Width - .ScaleWidth)
                If frmSeleccion.Visible Then
                    .Height = frmMapa.Height + frmSeleccion.Height + (.Height - .ScaleHeight)
                ElseIf frmChat.Visible Then
                    .Height = frmMapa.Height + frmChat.Height + (.Height - .ScaleHeight)
                ElseIf frmLog.Visible Then
                    .Height = frmMapa.Height + frmLog.Height + (.Height - .ScaleHeight)
                Else
                    .Height = frmMapa.Height + (.Height - .ScaleHeight)
                End If
            End If
        End If
    End With
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuOrganizar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuPartidaConectar_Click()
    On Error GoTo ErrorHandle
    
    MostrarFormulario frmComienzo, vbModal
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuPartidaConectar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuPartidaDesconectar_Click()
    On Error GoTo ErrorHandle
    
    If MsgBox(strMsgDesconectar, vbQuestion + vbYesNo + vbDefaultButton2, strMsgDesconectarCaption) = vbYes Then
        cDesconectar
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuPartidaDesconectar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuPartidaGuardar_Click()
    On Error GoTo ErrorHandle
    
    cPedirPartidasGuardadas

    Exit Sub
ErrorHandle:
    ReportErr "mnuPartidaGuardar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuPartidaIdioma_Click()
    On Error GoTo ErrorHandle
    
    frmIdioma.IdiomaSeteado = True
    MostrarFormulario frmIdioma, vbModal
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuPartidaIdioma_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub mnuPartidaPausar_Click()
    On Error GoTo ErrorHandle

    If GEstadoCliente <> estPartidaPausada Then
        'Pausar
        cPausarPartida
    Else
        'Continuar
        cContinuarPartida
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuPartidaPausar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuPartidaResincronizar_Click()
    On Error GoTo ErrorHandle
    
    cResincronizar

    Exit Sub
ErrorHandle:
    ReportErr "mnuPartidaResincronizar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuPartidaSalir_Click()
    
    Unload Me
    
End Sub

Private Sub mnuVerChat_Click()
    On Error GoTo ErrorHandle
    
    mnuVerChat.Checked = Not mnuVerChat.Checked
    frmChat.Visible = mnuVerChat.Checked
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuVerChat_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuVerDados_Click()
    On Error GoTo ErrorHandle
    
    mnuVerDados.Checked = Not mnuVerDados.Checked
    frmDados.Visible = mnuVerDados.Checked

    Exit Sub
ErrorHandle:
    ReportErr "mnuVerDados_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuVerJugadores_Click()
    On Error GoTo ErrorHandle
    
    mnuVerJugadores.Checked = Not mnuVerJugadores.Checked
    frmJugadores.Visible = mnuVerJugadores.Checked

    Exit Sub
ErrorHandle:
    ReportErr "mnuVerJugadores_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuVerListaMisiones_Click()
    On Error GoTo ErrorHandle
    
    mnuVerListaMisiones.Checked = Not mnuVerListaMisiones.Checked
    frmMisiones.Visible = mnuVerListaMisiones.Checked

    Exit Sub
ErrorHandle:
    ReportErr "mnuVerListaMisiones_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuVerLog_Click()
    On Error GoTo ErrorHandle
    
    mnuVerLog.Checked = Not mnuVerLog.Checked
    frmLog.Visible = mnuVerLog.Checked

    Exit Sub
ErrorHandle:
    ReportErr "mnuVerLog_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuVerMision_Click()
    On Error GoTo ErrorHandle
    
    mnuVerMision.Checked = Not (mnuVerMision.Checked)
    frmMision.Visible = mnuVerMision.Checked

    Exit Sub
ErrorHandle:
    ReportErr "mnuVerMision_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuVerSeleccion_Click()
    On Error GoTo ErrorHandle
    
    mnuVerSeleccion.Checked = Not (mnuVerSeleccion.Checked)
    frmSeleccion.Visible = mnuVerSeleccion.Checked
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuVerSeleccion_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuVerTarjetas_Click()
    On Error GoTo ErrorHandle
    
    mnuVerTarjetas.Checked = Not (mnuVerTarjetas.Checked)
    frmTarjetas.Visible = mnuVerTarjetas.Checked
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuVerTarjetas_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuVerInfo_Click()
    On Error GoTo ErrorHandle
    
    mnuVerInfo.Checked = Not (mnuVerInfo.Checked)
    frmPropiedades.Visible = mnuVerInfo.Checked

    Exit Sub
ErrorHandle:
    ReportErr "mnuVerInfo_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuVerMapa_Click()
    On Error GoTo ErrorHandle
    
    mnuVerMapa.Checked = Not (mnuVerMapa.Checked)
    frmMapa.Visible = mnuVerMapa.Checked
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuVerMapa_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuVerTropasDisponibles_Click()
    On Error GoTo ErrorHandle
    
    mnuVerTropasDisponibles.Checked = Not (mnuVerTropasDisponibles.Checked)
    frmTropasDisponibles.Visible = mnuVerTropasDisponibles.Checked
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuVerTropasDisponibles_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub ActualizarMenu()
    On Error GoTo ErrorHandle
    
    mnuVerInfo.Checked = frmPropiedades.Visible
    mnuVerMapa.Checked = frmMapa.Visible
    mnuVerTarjetas.Checked = frmTarjetas.Visible
    mnuVerMision.Checked = frmMision.Visible
    mnuVerJugadores.Checked = frmJugadores.Visible
    mnuVerChat.Checked = frmChat.Visible
    mnuVerDados.Checked = frmDados.Visible
    mnuVerSeleccion.Checked = frmSeleccion.Visible
    mnuVerTropasDisponibles.Checked = frmTropasDisponibles.Visible
    mnuVerLog.Checked = frmLog.Visible
    mnuVerListaMisiones.Checked = frmMisiones.Visible
    
    Exit Sub
ErrorHandle:
    ReportErr "ActualizarMenu", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Timer1_Timer()
    On Error GoTo ErrorHandle
    
    'Dim NuevoValor As Integer
    Dim sngNuevoValorTimer As Single
    Dim sngTiempoRestante As Single
    
    'Actualiza el tiempo restante del turno
    sngNuevoValorTimer = Timer
    If sngNuevoValorTimer < GsngInicioTimerTurno Then
        GsngInicioTimerTurno = GsngInicioTimerTurno - (24# * 60# * 60#)
    End If
    
    sngTiempoRestante = GintTimerTurno - (sngNuevoValorTimer - GsngInicioTimerTurno)
    StatusBar1.Panels(enuPaneles.panTimer).Text = CLng(sngTiempoRestante)
    If sngTiempoRestante < 5 And GintColorActual = GintMiColor Then
        Beep
    End If
        
'    If IsNumeric(StatusBar1.Panels(enuPaneles.panTimer).Text) Then
'        NuevoValor = CInt(StatusBar1.Panels(enuPaneles.panTimer).Text) - 1
'        If NuevoValor < 0 Then NuevoValor = 0
'        StatusBar1.Panels(enuPaneles.panTimer).Text = NuevoValor
'    End If

    Exit Sub
ErrorHandle:
    ReportErr "Timer1_Timer", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub tmrSysTray_Timer()
    On Error GoTo ErrorHandle
    
    'Cuando le toca el turno el ícono titila y queda resaltado
    If intFlagSysTray >= 0 Then
        If intFlagSysTray Mod 2 = 0 Then
            'Si es par, muestra el resaltado
            SysTrayChangeIcon Me.hwnd, imgSysTrayTurno
        Else
            SysTrayChangeIcon Me.hwnd, imgSysTrayOK
        End If
        intFlagSysTray = intFlagSysTray - 1
    Else
        'Deja de titilar
        tmrSysTray.Enabled = False
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "Toolbar1_ButtonClick", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrorHandle
    
    Select Case Button.Index
        Case enuToolBar.tbConexion
            If GEstadoCliente = estDesconectado Then
                mnuPartidaConectar_Click
            Else
                mnuPartidaDesconectar_Click
            End If
        Case enuToolBar.tbResincronizacion
            mnuPartidaResincronizar_Click
        Case enuToolBar.tbOpciones
            MnuOpciones_Click
        Case enuToolBar.tbGuardar
            mnuPartidaGuardar_Click
        Case enuToolBar.tbPausa
            mnuPartidaPausar_Click
        Case enuToolBar.tbMision
            mnuVerMision_Click
        Case enuToolBar.tbVerTarjetas
            mnuVerTarjetas_Click
        Case enuToolBar.tbAgregar
            mnuJuegoAgregar_Click
        Case enuToolBar.tbAtacar
            mnuJuegoAtacar_Click
        Case enuToolBar.tbMover
            mnuJuegoMover_Click
        Case enuToolBar.tbTomarTarjeta
            mnuJuegoTomar_Click
        Case enuToolBar.tbFinTurno
            mnuJuegoFinalizar_Click
    End Select

    Exit Sub
ErrorHandle:
    ReportErr "Toolbar1_ButtonClick", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Winsock1_Close()
    On Error GoTo ErrorHandle
    
    cDesconectar
    
    '###
    'Segun como se haya Cerrado el Servidor,
    'muestra el mensaje apropiado.
    Select Case GintServidorCerrado
        Case enuServidorCerrado.secEventual
            MsgBox strMsgWinsockClose, vbExclamation, strMsgWinsockCloseCaption
        Case enuServidorCerrado.secVoluntaria
            MsgBox ObtenerTextoRecurso(CintPrincipalMsgServidorCerrado), vbExclamation, strMsgWinsockCloseCaption
        Case enuServidorCerrado.secSalida
            'No muestra mensaje alguno
    End Select
   
    Exit Sub
ErrorHandle:
    ReportErr "Winsock1_Close", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Winsock1_Connect()
    On Error GoTo ErrorHandle
    
    'Cambia el icono de la barra
    mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbConexion).Image = 13
    'Habilita y Deshabilita las opciones de menu
    mdifrmPrincipal.mnuPartidaDesconectar.Enabled = True
    mdifrmPrincipal.mnuPartidaConectar.Enabled = False
    'Resetea el valor de esta variable global
    GintServidorCerrado = secEventual
    
    If frmComienzo.Visible And frmComienzo.fraEtapas(enuEtapas.etaUnirse).Visible Then
        frmComienzo.Label4.Caption = ""
        Screen.MousePointer = vbDefault
        Unload frmComienzo
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "Winsock1_Connect", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ErrorHandle
    
    Dim strMensaje As String
    
    StatusBar1.Panels(enuPaneles.panIcoIntercambio).Picture = imgLstFichas.ListImages(7).Picture
    Winsock1.GetData strMensaje
    SepararMensajes strMensaje
    StatusBar1.Panels(enuPaneles.panIcoIntercambio).Picture = Nothing
    
    Exit Sub
ErrorHandle:
    ReportErr "Winsock1_DataArrival", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error GoTo ErrorHandle
    
    '### Cachear error
    Select Case Number
        Case 10061:
            'No existe un servidor activo en el host/puerto especificado
            If frmComienzo.Visible And frmComienzo.fraEtapas(enuEtapas.etaUnirse).Visible Then
                frmComienzo.Label4.Caption = Mid$(strMsgConexion1, 1, Len(strMsgConexion1) - 2) & "."
            Else
                MsgBox strMsgConexion1 & GstrServidor & ":" & GintPuerto & ").", vbCritical, strMsgConexion1Caption
            End If
        Case 11004:
            'No existe el host especificado
            If frmComienzo.Visible And frmComienzo.fraEtapas(enuEtapas.etaUnirse).Visible Then
                frmComienzo.Label4.Caption = strMsgConexion2
            Else
                MsgBox strMsgConexion2 & GstrServidor & ".", vbCritical, strMsgConexion2Caption
            End If
        Case Else
            If frmComienzo.Visible And frmComienzo.fraEtapas(enuEtapas.etaUnirse).Visible Then
                frmComienzo.Label4.Caption = strMsgConexionOtro & Number & " - " & Description
            Else
                MsgBox strMsgConexionOtro & Number & " - " & Description, vbCritical, strMsgConexionOtro
            End If
    End Select
    Screen.MousePointer = vbDefault
    
    '###E
    If GEstadoCliente < estConectado Then
        'Cierra el puerto para poder volver a conectarse
        cDesconectar
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "Winsock1_Error", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

