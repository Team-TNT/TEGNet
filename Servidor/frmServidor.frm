VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servidor TEGNet"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "frmServidor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Información de la Partida"
      Height          =   870
      Left            =   90
      TabIndex        =   13
      Top             =   1410
      Width           =   6075
      Begin VB.Label lblTimer 
         Height          =   210
         Left            =   2160
         TabIndex        =   17
         Top             =   555
         Width           =   1665
      End
      Begin VB.Label Label2 
         Caption         =   "Timer del Turno:"
         Height          =   210
         Left            =   105
         TabIndex        =   16
         Top             =   555
         Width           =   1470
      End
      Begin VB.Label lblEstadoJugadorActivo 
         Height          =   210
         Left            =   2160
         TabIndex        =   15
         Top             =   270
         Width           =   2235
      End
      Begin VB.Label Label5 
         Caption         =   "Estado del Jugador Activo:"
         Height          =   210
         Left            =   105
         TabIndex        =   14
         Top             =   270
         Width           =   2040
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Servidor"
      Height          =   1335
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   6075
      Begin VB.Label lblPuertoServer 
         Height          =   210
         Left            =   2205
         TabIndex        =   12
         Top             =   750
         Width           =   1665
      End
      Begin VB.Label Label9 
         Caption         =   "Puerto:"
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   750
         Width           =   1470
      End
      Begin VB.Label lblIPServer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2205
         TabIndex        =   10
         Top             =   510
         Width           =   1665
      End
      Begin VB.Label Label8 
         Caption         =   "Dirección IP:"
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   510
         Width           =   1470
      End
      Begin VB.Label lblNombreServer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2205
         TabIndex        =   8
         Top             =   255
         Width           =   1665
      End
      Begin VB.Label Label7 
         Caption         =   "Nombre:"
         Height          =   210
         Left            =   150
         TabIndex        =   7
         Top             =   255
         Width           =   1470
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   5280
         Picture         =   "frmServidor.frx":0442
         Top             =   450
         Width           =   480
      End
      Begin VB.Image imgSystray 
         Height          =   240
         Index           =   0
         Left            =   4560
         Picture         =   "frmServidor.frx":074C
         Top             =   225
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSystray 
         Height          =   240
         Index           =   1
         Left            =   4560
         Picture         =   "frmServidor.frx":0896
         Top             =   225
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Estado:"
         Height          =   210
         Left            =   150
         TabIndex        =   5
         Top             =   990
         Width           =   1425
      End
      Begin VB.Label lblEstado 
         Height          =   210
         Left            =   2205
         TabIndex        =   4
         Top             =   990
         Width           =   2235
      End
   End
   Begin VB.ListBox lstLog 
      Height          =   2205
      Left            =   75
      TabIndex        =   0
      Top             =   2640
      Width           =   6075
   End
   Begin VB.Timer tmrTurno 
      Left            =   3000
      Top             =   2160
   End
   Begin VB.Timer tmrAYA 
      Left            =   2520
      Top             =   2160
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   2040
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblCopyRight 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2001-2004 TEGNet Team"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   3165
      TabIndex        =   6
      Top             =   4920
      Width           =   2940
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      Caption         =   "Versión:"
      Height          =   240
      Left            =   4500
      TabIndex        =   2
      Top             =   2370
      Width           =   1530
   End
   Begin VB.Label Label6 
      Caption         =   "Bitácora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   1
      Top             =   2370
      Width           =   825
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "Menu Systray"
      Visible         =   0   'False
      Begin VB.Menu mnuRestaurar 
         Caption         =   "&Restaurar"
      End
      Begin VB.Menu mnuSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intClienteSeleccionado As Integer

Private Sub Form_Load()
    On Error GoTo ErrorHandle
        
    SysTrayInicializar Me.hwnd, "Servidor TEGNet", imgSystray(0)
    
    'Muestra la Versión del servidor
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    
    'Lee del Registro el Flag que habilita el DEBUG
    MODO_DEBUG = GetSetting("TEGNet", "Server", "Debug")
    
    'Para que no oculte el formulario al iniciarse, descomentar esta linea
    'Me.WindowState = vbNormal
    
    ConectarBaseDatos
    GintCantConexiones = 0
    wskServer(0).LocalPort = IIf(Trim(Command) = "", 5479, Command)
    wskServer(0).Listen
    
    lblIPServer.Caption = wskServer(0).LocalIP
    lblNombreServer.Caption = wskServer(0).LocalHostName
    lblPuertoServer.Caption = wskServer(0).LocalPort
    
    GintIndiceAdm = -1
    GEstadoServidor = estEsperandoAdm
    GintEsperaAckAYA = 8000
    
    CargarMatrizEstados
    
    CargarVectorColores
    CargarVectorPaises
    CargarVectorContinentes
    
    'Inicializa el timer en 9999 para que el cliente interprete un timer
    'infinito al comenzar una partida guardada
    GintValorTimerActual = CintValorTimerInfinito

    '###E Borrar
    vecEstadosServidor(enuEstadoSrv.estConfigurandoServidor) = "Configurando Servidor"
    vecEstadosServidor(enuEstadoSrv.estEjecutandoPartida) = "Ejecutando Partida"
    vecEstadosServidor(enuEstadoSrv.estEsperandoAdm) = "Esperando Adm"
    vecEstadosServidor(enuEstadoSrv.estEsperandoJugadores) = "Esperando Jugadores"
    vecEstadosServidor(enuEstadoSrv.estPartidaDetenida) = "Partida Pausada"
    
    CambiarEstadoServidor estEsperandoAdm
    
    blnPreguntar = True
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Load", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandle
    
    SysTrayMouseMove Me, Button, Shift, X, Y

    Exit Sub
ErrorHandle:
    ReportErr "Form_MouseMove", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorHandle
    If blnPreguntar Then
        If MsgBox("¿Está seguro que desea cerrar el Servidor?", vbQuestion + vbYesNo + vbDefaultButton2, "Salir") <> vbYes Then
            Cancel = 1
        End If
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
    
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrorHandle
    
    SysTrayResize Me

    Exit Sub
ErrorHandle:
    ReportErr "Form_Resize", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    
    'Crea en el Registro el Flag que habilita el DEBUG
    'SaveSetting "TEGNet", "Server", "Debug", MODO_DEBUG
    
    SysTrayUnload Me
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Unload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuRestaurar_Click()
    On Error GoTo ErrorHandle
    
    SystrayRestaurar Me

    Exit Sub
ErrorHandle:
    ReportErr "mnuRestaurar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub mnuSalir_Click()
    On Error GoTo ErrorHandle
    
    SysTrayExit Me
    
    Exit Sub
ErrorHandle:
    ReportErr "mnuSalir_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub tmrAYA_Timer()
    On Error GoTo ErrorHandle
    
    'Si se agotó el tiempo es porque falló el AYA
    GintRtaAYA = rtaEstaMuerto
    
    'Desactiva el timer
    tmrAYA.Interval = 0
    
    Exit Sub
ErrorHandle:
    ReportErr "tmrAYA_Timer", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub tmrTurno_Timer()
    On Error GoTo ErrorHandle
    
    Dim sngNuevoValorTimer As Single
    Dim sngTiempoRestante As Single
    
    'Actualiza el tiempo restante del turno
    sngNuevoValorTimer = Timer
    If sngNuevoValorTimer < GsngInicioTimerTurno Then
        GsngInicioTimerTurno = GsngInicioTimerTurno - (24# * 60# * 60#)
    End If
    
    sngTiempoRestante = GintValorTimerTotal - (sngNuevoValorTimer - GsngInicioTimerTurno)
    'Actualiza el valor actual del timer
    GintValorTimerActual = CInt(sngTiempoRestante)

    lblTimer.Caption = CLng(sngTiempoRestante)
    
    If sngTiempoRestante <= 0 Then
        sFinTurno True
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "tmrTurno_Timer", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub wskServer_Close(Index As Integer)
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    Dim ExisteAlguien As Boolean
    Dim intPosibleAdm As Integer
    
    'Al cerrarse la conexión no envia mensaje, entonces se fuerza
    'el origen del mensaje
    GintOrigenMensaje = Index
    
    ExisteAlguien = False
    intPosibleAdm = 0
    
    wskServer(Index).Close
    
    'Si no queda ninguna conexion abierta se baja el servidor
    For i = 1 To wskServer.Count - 1
        If wskServer(i).State <> 0 Then 'Cerrado
            ExisteAlguien = True
            'Busca un posible Adm Humano en caso que se cayera el Adm actual
            If i <> GintIndiceAdm Then
                '# Version 1.0.0
                If wskServer(i).Tag <> "" Then
                    If Mid$(wskServer(i).Tag, 1, 1) = enuInteligenciaJugador.hrHumano Then
                        intPosibleAdm = i
                    End If
                End If
            End If
        End If
    Next
            
    If Not ExisteAlguien Then
        sBajarServidor
    Else
        'Si el que se cayó es el Administrador...
        If Index = GintIndiceAdm Then
            If intPosibleAdm > 0 Then
                'Designa a otro jugador humano como Administrador
                sCambiarAdm IndiceAColor(intPosibleAdm), False
            Else
                'La Administración se deja disponible para el proximo Humano que se conecte
                GintIndiceAdm = -1
            End If
        End If
        
        sBajaJugador Index
        
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "wskServer_Close", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub wskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    Dim existeLibre As Integer
    existeLibre = 0
    'Los pedidos de conexion siempre llegan al winsock 0
    'Es el unico que escucha por un puerto
    If Index = 0 Then
        
        'Si hay algun socket libre
        For i = 1 To wskServer.Count - 1
            If wskServer(i).State = 0 Then 'Cerrado
                existeLibre = i
            End If
        Next
        
        'De acuerdo al estado en que se encuentre el Server...
        Select Case GEstadoServidor
        Case estEsperandoAdm, estConfigurandoServidor
            'Si el server todavía no se encuentra recibiendo conexiones...
            If GintCantConexiones > 0 Then
                'Si ya existe una conexión
                If wskServer(1).State = 0 Then
                    'Si la misma está cerrada le asigna la conexion entrante
                    AceptarPedido requestID, 1
                Else
                    'Si la conexión está abierta envia una AYA
                    'para determinar si el administrador está vivo
                    If Not estaVivo(1) Then
                        'Si no está vivo lo desconecta y le asigna la conexión
                        'al nuevo pedido
                        wskServer(1).Close
                        AceptarPedido requestID, 1
                    Else
                        '###Rechaza conexion
                    End If
                End If
            Else
                'Si no existe una conexion se acepta el pedido y se lo
                'crea como administrador
                AceptarPedido requestID
            End If
        Case estEsperandoJugadores
            'Si el servidor se encuentra recibiendo conexiones
            If existeLibre > 0 Then
                'Si existe alguna conexion libre se la asigna
                AceptarPedido requestID, existeLibre
            Else
                If GintCantConexiones >= CintCantMaxConexiones Then
                    'Si se superó la cantidad maxima de conexiones
                    'se rechaza el pedido
                    '###Rechaza conexion
                Else
                    'Si no se superó la cantidad maxima se le asigna
                    'el siguiente socket libre
                    AceptarPedido requestID
                End If
            End If
        Case estEjecutandoPartida
            'Si el servidor no se encuentra recibiendo conexiones
            'Acepta el pedido con un nuevo indice
            AceptarPedido requestID
        Case estPartidaDetenida
            'Si el Servidor está Pausado,
            'solo aceptará que se reconecte un jugador si no hay ADM.
            If GintIndiceAdm = -1 Then
                AceptarPedido requestID
            Else
                '###Rechaza conexion
            End If
                
        End Select
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "wskServer_ConnectionRequest", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub AceptarPedido(ByVal requestID As Long, Optional intPosicion As Integer = 0)

    On Error GoTo ErrorHandle
    Dim intNroSocket As Integer
    
    If intPosicion = 0 Then
        'Si se pide que se cree una nueva instancia
        'Se establece una nueva conexión
        GintCantConexiones = GintCantConexiones + 1
        
        ReDim Preserve GvecColoresSock(GintCantConexiones)
        
        Load wskServer(GintCantConexiones)
        wskServer(GintCantConexiones).LocalPort = 0
        wskServer(GintCantConexiones).Accept requestID
        
        intNroSocket = GintCantConexiones
        
    Else
        'Utiliza un socket ya existente
        wskServer(intPosicion).LocalPort = 0
        wskServer(intPosicion).Accept requestID
        
        intNroSocket = intPosicion
        
    End If
    
    '###Log
    GuardarLog wskServer(intNroSocket).RemoteHostIP & " se ha conectado fisicamente."
    
    'Limpia la versión del cliente
    wskServer(intNroSocket).Tag = ""
    'Envia la version del servidor al cliente
    sEnviarVersionServidor intNroSocket
    
    'Si no hay ningun Adm
    If GintIndiceAdm = -1 Then
        'Confirma al cliente como Administrador
        GintIndiceAdm = intNroSocket
        sConfirmarAdm
        
        'Envía las opciones por defecto al nuevo Administrador
        sEnviarOpcionesDefault
    End If
    
    If GintCantConexiones > 1 Then
        'Envia los colores disponibles y las conexiones actuales
        sConexionesActuales
    End If

    Exit Sub
ErrorHandle:
    ReportErr "AceptarPedido", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub wskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error GoTo ErrorHandle
    
    Dim strMensajes As String
    
    wskServer(Index).GetData strMensajes
    
    GintOrigenMensaje = Index
    SepararMensajes strMensajes, Index

    Exit Sub
ErrorHandle:
    ReportErr "wskServer_DataArrival", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

