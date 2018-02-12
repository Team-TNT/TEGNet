Attribute VB_Name = "mdlVP"
Option Explicit

#Const MODO_DEBUG = 0

'Constantes VP
Public Const CsngMaxValorPais = 45

'JUGADOR VIRTUAL
'-----------------------------------------------
Public Type typObjetivo
    esDestruir As Boolean
    intColorADestruir As Integer
    intContinente As enuContinentes
    intCantidadPaises As Integer
    sonLimitrofes As Boolean
    sngPorcentajeTerminado As Single
End Type

Public Type typContinente
    intBonus As Integer
    intCantidadPaises As Integer
    intCantidadPaisesMios As Integer
End Type

Public Enum enuContinentes
    coAfrica = 1
    coANorte
    coASur
    coAsia
    coEuropa
    coOceania
End Enum

Public GvecObjetivos() As typObjetivo
Public GvecContinentes(0 To 6) As typContinente

'Tipos globales
'-----------------------------------------------
Public Type typPais
    strNombre As String
    intColor As Integer
    intEstado As Integer
    intCantidadFichas As Integer
    intContinente As enuContinentes
    sngValor As Single
    intTropasFijas As Integer   'Almacena la cantidad de tropas que ya fueron movidas en el turno
End Type

Public Type typTarjeta
    byPais As Byte
    byFigura As Byte
    blCobrada As Boolean 'Si fue o no usada
End Type

Public Type typJugador
    strNombre As String
    intOrdenRonda As Integer    'Es 0 cuando no juega
    intCanje As Integer
    intCantidadTarjetas As Integer
    intTropasDisponibles As Integer
    intEstado As enuEstadoConexion
    vecDetalleTropasDisponibles(0 To 6) As Integer
End Type

Private Type typTercetos
    intTarjeta1 As Integer
    intTarjeta2 As Integer
    intTarjeta3 As Integer
End Type

Private Type typLimite
    byPaisDesde As Byte
    byPaisHasta As Byte
End Type
    
'Enumeraciones
'---------------------------------------------------

'Eventos que pueden modificar el estado del cliente
Public Enum enuEventosCli
    eveConfirmacionAdm
    eveJugadoresConectados
    eveConfirmacionAlta
    eveInicioPartida
    eveInicioTurnoRecuento
    eveInicioTurnoAccion
    eveAckFinTurno
    eveAckMoverTropa
    eveTarjetaTomada
    eveTarjetaCobrada
    eveMisionCumplida
    eveCierreConexion
    evePartidaPausada
    evePartidaContinuada
End Enum

'Matriz de estados
Public GMatrizEstados(0 To 12, 0 To 14) As enuEstadoCli

'Variables globales
Public GEstadoCliente As enuEstadoCli
Public GvecEstadoCliente(12) As String

Public GvecPaises() As typPais

Public GintCantJugadores As Integer

Public GsoyAdministrador As Boolean

'Color Jugador Local
Public GintMiColor As Integer

'Nombre Jugador Local
Public GstrMiNombre As String

'Este vector contiene información de todos los jugadores
'Su indice representa el color asignado
Public GvecJugadores(6) As typJugador

Public GvecTarjetas(5) As typTarjeta

Public GbyPaisActual As Byte
Public GbyPaisAnterior As Byte

'Color que está jugando
Public GintColorActual As Integer

'Indica el tipo de ronda
Public GintTipoRonda As enuTipoRonda

'Mision
Public GstrMision As String

'###
Public GvecColores(6) As Long
Public GvecColoresInv(6) As Long 'Guarda los colores inversos a los colores


'Propias del VP
Public GintPuerto As Integer
Public GstrServer As String
Public GstrJvNombre As String
Public GintJvColor As Integer
Public GintSegTimeOut As Integer
Public GintSegRestantesTimeOut As Integer

'Variables que tienen que ver con el cálculo del valor de cada pais
Public GsngFactorBonusContinente As Single  'Importancia que tiene la conquista de un continente (se lo multiplicara por el bonus del continente)
Public GsngFactorObjetivo As Single
Public GsngFactorMision As Single
Public GsngBonusPorTarjeta As Single        'Valor que tiene un pais por tener una tarjeta no cobrada
Public GsngUmbralAtaque As Single
Public GsngUmbralAtaqueOriginal As Single   'Almacena el umbral seleccionado por el usuario
Public GsngActitud As Single            'De 0 a 1 indica que tan importante es el ataque en relación a la defensa
Public GsngFactorReduccionUmbral As Single  'Factor en que se reduce el umbral de ataque hasta la primera conquista

Public GblnConquistoAlgunPais As Boolean

Public GintMsPausaMovimiento As Integer
Public GintMsPausaAtaque As Integer
Public GintMsPausaAgregar As Integer


Public GvecLimitrofes() As typLimite

'### Hacerla desaparecer
Public Sub Hardcodear()
    GintSegTimeOut = 20
    
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
    
    'Pausas
    GintMsPausaAtaque = 4000
    GintMsPausaAgregar = 2000
    GintMsPausaMovimiento = 2000
    
    'Setea los valores para el cálculo del valor de cada pais
    GsngFactorBonusContinente = 5
    GsngBonusPorTarjeta = 5
    GsngFactorObjetivo = 10
    GsngFactorMision = 50
    GsngFactorReduccionUmbral = 2
    
End Sub

Public Sub EnviarMensaje(strMensaje As String)
    On Error GoTo ErrorHandle
    
    '###
    frmPrincipal.txtEnviado.Text = strMensaje
    
    frmPrincipal.wskVP.SendData strMensaje
    
    '###Borrar Mensaje
    #If MODO_DEBUG = 1 Then
        Print #2, strMensaje
    #End If
    
    Exit Sub
ErrorHandle:
    ReportErr "EnviarMensaje", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub DistribuirMensaje(TipoMensaje As enuTipoMensaje, vecParametros() As String, Optional intIndiceOrigenMensaje As Integer)
    On Error GoTo ErrorHandle

    'De acuerdo al tipo de mensaje
    Select Case TipoMensaje
        Case msgConfirmarAdm
            '###E
'            ActualizarEstadoCliente eveConfirmacionAdm
'            cConfirmarAdm
        Case msgJugadoresConectados
            '###E
            ActualizarEstadoCliente eveJugadoresConectados
            cConexionesActuales vecParametros
        Case msgPais
            cActualizarPais CByte(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2)), CInt(vecParametros(3)), CInt(vecParametros(4))
        Case msgMisionAsignada
            cMostrarMision vecParametros
        Case msgAckAltaJugador
            cConfirmarAlta CInt(vecParametros(1)), CStr(vecParametros(2)), CInt(vecParametros(0))
        Case msgOrdenRonda
            cActualizarRonda vecParametros
'        Case msgComienzoTurno
'            cInicioTurno CInt(vecParametros(0)), CInt(vecParametros(1))
        Case msgChatEntrante
'            cMensajeChatEntrante CInt(vecParametros(0)), vecParametros(1)
        Case msgPartidasGuardadas
'            cPartidasGuardadas vecParametros
        Case msgAYA
            cEnviarIAA
        Case msgOpciones
'            cRecibirOpciones vecParametros
        Case msgOpcionesDefault
'            cRecibirOpcionesDefault vecParametros
        Case msgBajaAdm
'            cConfirmarBajaAdm
        Case msgInicioTurno
            cInicioTurno CInt(vecParametros(0)), CInt(vecParametros(1)), CBool(vecParametros(2))
        Case msgAckInicioPartida
            cConfirmarInicioPartida
        Case msgTropasDisponibles
            cTropasDisponibles vecParametros
            'Es una respuesta a un mensaje de poner
            'EfectuarAccion
        Case msgAckFinTurno
            cConfirmarFinTurno CBool(vecParametros(0))
        Case msgTipoRonda
            cActualizarTipoRonda CInt(vecParametros(0))
        Case msgAckAgregarTropas
            EfectuarAccion
        Case msgAckAtaque
            cAckAtacar CInt(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2)), _
                       CInt(vecParametros(3)), CInt(vecParametros(4)), CInt(vecParametros(5)), _
                       CByte(vecParametros(6)), CInt(vecParametros(7)), CInt(vecParametros(8)), _
                       CByte(vecParametros(9)), CInt(vecParametros(10)), CInt(vecParametros(11))
        Case msgAckMovimiento
            cAckMover CByte(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2)), _
                      CByte(vecParametros(3)), CInt(vecParametros(4)), CInt(vecParametros(5)), _
                      CInt(vecParametros(6)), CInt(vecParametros(7))
        Case msgTarjeta
            cMostrarTarjeta CByte(vecParametros(0)), CByte(vecParametros(1)), _
                            IIf(Trim(UCase(vecParametros(2))) = "S", True, False)
        Case msgTarjetasJugador
            cActualizarTarjetasJugador CInt(vecParametros(0)), CInt(vecParametros(1))
        Case msgAckCobroTarjeta
            cAckCobrarTarjeta CByte(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2))
        Case msgAckCanjeTarjeta
            cAckCanjearTarjeta CByte(vecParametros(0)), CByte(vecParametros(1)), CByte(vecParametros(2))
        Case msgEstadoTurnoCliente
            cActualizarEstadoTurno CInt(vecParametros(0))
        Case msgMisionCumplida
            cMisionCumplida
        Case msgVersionServidor
            cRecibirVersionServidor CInt(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2))
        Case msgLimitrofes
            CargarLimitrofesServidor vecParametros
        Case msgPaisContinente
            CargarPaisContinente vecParametros
        Case msgPartidaPausada
            cPartidaPausada
        Case msgPartidaContinuada
            cPartidaContinuada
        Case enuTipoMensaje.msgError
            'cResincronizar
            cMostrarError CStr(vecParametros(0)), CInt(vecParametros(1))
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "DistribuirMensaje", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

'-------------------------------------------------------
'-         INTERFASE                                   -
'-------------------------------------------------------

Public Sub cConectar(strServer As String, intPort As Integer)
    On Error GoTo ErrorHandle
    
    'Conecta al cliente con el servidor
    frmPrincipal.wskVP.RemoteHost = strServer
    frmPrincipal.wskVP.RemotePort = intPort
    ' Invoca el método Connect para iniciar
    ' una conexión.
    frmPrincipal.wskVP.Connect
    
    Exit Sub
ErrorHandle:
    ReportErr "cConectar", "mdlVP", Err.Description, Err.Number, Err.Source
    'Si no logra conectarse al servidor muere
    End
End Sub

Public Sub cConexionesActuales(pVecJugadores() As String)
    'Busca un color disponible y se conecta
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim intColorSeleccionado As Integer
    Dim intTipoJugadoresConectados As enuTipoJugadoresConectados
    
    'Toma del mensaje el Tipo que se encuentra en la última posición
    'del vector Jugadores para saber si se está reconectando
    intTipoJugadoresConectados = pVecJugadores(UBound(pVecJugadores))
    
    'Elige un color
'    For i = 1 To UBound(GvecJugadores)
'        If pVecJugadores(i - 1) <> "" Then
'            'Si el jugador juega
'            GvecJugadores(i).intEstado = conConectado
'        Else
'            'Si el jugador no juega
'            intColorSeleccionado = i
'            GvecJugadores(i).intEstado = conNoJuega
'        End If
'        GvecJugadores(i).strNombre = pVecJugadores(i - 1)
'    Next
    
    'Si todavía no está validado da de alta al jugador
    If GEstadoCliente < estValidado Then
        If intTipoJugadoresConectados = tjcIngreso Then
            cAltaJugador GintJvColor, GstrJvNombre
        Else
            cReconectarJugador GintJvColor
        End If
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cRefrescarConexiones", "mdlVP", Err.Description, Err.Number, Err.Source
    'Si no logra conectarse al servidor muere
    End
End Sub

Public Sub cReconectarJugador(intColor As Integer)
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgReconexion, intColor)
    
    Exit Sub
ErrorHandle:
    ReportErr "cReconectarJugador", "mdlVP", Err.Description, Err.Number, Err.Source
    'Si no logra conectarse al servidor muere
    End
End Sub

Public Sub cVolverAConectarse(strNickName As String)
    'Selecciona el primer color libre que encuentra y se conecta con
    'el nombre pasado por parametro
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim intColorSeleccionado As Integer
    
    For i = 1 To UBound(GvecJugadores)
        If GvecJugadores(i).intEstado = conNoJuega Or GvecJugadores(i).intEstado = conDesconectado Then
            intColorSeleccionado = i
            i = UBound(GvecJugadores) 'exit for
        End If
    Next i
    
    'Si todavía no está validado da de alta al jugador
    If GEstadoCliente < estValidado Then
        cAltaJugador intColorSeleccionado, strNickName
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cVolverAConectarse", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub
Public Sub cAltaJugador(intColorSeleccionado As Integer, strNickNameSeleccionado As String)
    'Envia al servidor el color y NickName seleccionado por el jugador
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgAltaJugador, CStr(intColorSeleccionado), strNickNameSeleccionado)
        
    Exit Sub
ErrorHandle:
    ReportErr "cAltaJugador", "mdlVP", Err.Description, Err.Number, Err.Source
    'Si no logra conectarse al servidor muere
    End
End Sub

Public Sub cConfirmarAlta(intColorAsignado As Integer, strNombreAsignado As String, intCodAck As enuAckAltaJugador)
    On Error GoTo ErrorHandle
    Dim strMensaje As String
    
    Select Case intCodAck
        Case enuAckAltaJugador.ackOk
            '###E
            ActualizarEstadoCliente eveConfirmacionAlta
            
            GintMiColor = intColorAsignado
            GstrMiNombre = strNombreAsignado
            
'        Case enuAckAltaJugador.ackColorUsado
            'Intenta volver a conectarse con el mismo nombre
'            cVolverAConectarse GstrJvNombre
'        Case enuAckAltaJugador.ackNombreUsado
            'Cambia el NickName e intenta volver a conectarse
'            GstrJvNombre = GstrJvNombre & "1"
'            cVolverAConectarse GstrJvNombre
'        Case enuAckAltaJugador.ackNombreYColorUsados
            'Cambia el NickName e intenta volver a conectarse
'            GstrJvNombre = GstrJvNombre & "1"
'            cVolverAConectarse GstrJvNombre
        Case Else
            End
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "cConfirmarAlta", "mdlVP", Err.Description, Err.Number, Err.Source
    'Si no logra conectarse al servidor muere
    End
End Sub

Public Sub cConfirmarInicioPartida()
    'Cambia el estado a jugando
    On Error GoTo ErrorHandle
    
    'Calcula la cantidad de paises que tiene cada continente
    GvecContinentes(enuContinentes.coAfrica).intCantidadPaises = CantidadPaises(coAfrica)
    GvecContinentes(enuContinentes.coANorte).intCantidadPaises = CantidadPaises(coANorte)
    GvecContinentes(enuContinentes.coAsia).intCantidadPaises = CantidadPaises(coAsia)
    GvecContinentes(enuContinentes.coASur).intCantidadPaises = CantidadPaises(coASur)
    GvecContinentes(enuContinentes.coEuropa).intCantidadPaises = CantidadPaises(coEuropa)
    GvecContinentes(enuContinentes.coOceania).intCantidadPaises = CantidadPaises(coOceania)
    
    ActualizarEstadoCliente eveInicioPartida
    
    Exit Sub
ErrorHandle:
    ReportErr "cConfirmarInicioPartida", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cActualizarPais(byPais As Byte, intColor As Integer, intCantidad As Integer, intOrigen As enuOrigenMsgPais, intTropasFijas As Integer)
    On Error GoTo ErrorHandle
    
    GvecPaises(byPais).intColor = intColor
    GvecPaises(byPais).intCantidadFichas = intCantidad
    GvecPaises(byPais).intTropasFijas = intTropasFijas
    
'''    '###
'''    frmPrincipal.shpFicha(byPais).FillColor = GvecColores(intColor)
'''    frmPrincipal.lblFicha(byPais).ForeColor = GvecColoresInv(intColor)
'''    frmPrincipal.shpFicha(byPais).BorderColor = GvecColoresInv(intColor)
'''    frmPrincipal.lblFicha(byPais).Caption = CStr(intCantidad)
    
    '###
    CalcularMisPaisesPorContinente
    MostrarPorcentajeCompletitudContinentes

    If intOrigen <> orRepartoInicial Then
        CalcularValorPais
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cActualizarPais", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cActualizarRonda(vecRonda() As String)
'Public Sub cActualizarRonda(intColor1 As Integer, intColor2 As Integer, _
                            intColor3 As Integer, intColor4 As Integer, _
                            intColor5 As Integer, intColor6 As Integer)

    On Error GoTo ErrorHandle
    Dim i As Integer
    
    'Limpia el orden de la ronda
    For i = LBound(GvecJugadores) To UBound(GvecJugadores)
        GvecJugadores(i).intOrdenRonda = 0
    Next i
    
    For i = LBound(vecRonda) To UBound(vecRonda)
        GvecJugadores(CInt(vecRonda(i))).intOrdenRonda = i + 1
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "cActualizarRonda", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMostrarMision(vecNuevaMision() As String)
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim strObjetivo As String
    Dim indice As Integer
    
    '###Borrar
    frmPrincipal.txtMision.Text = vecNuevaMision(0)
    
    ReDim GvecObjetivos(0)
    'Solo toma del mensaje los objetivos (a partir de la tercer posicion
    For i = 2 To UBound(vecNuevaMision)
        indice = Int((i - 2) / 4)
        Select Case (i - 2) Mod 4
            Case 0
                'Redimensiona el vector de objetivos
                ReDim Preserve GvecObjetivos(indice)
                'Continente
                If IsNull(vecNuevaMision(i)) Then
                    GvecObjetivos(indice).intContinente = 0
                Else
                    GvecObjetivos(indice).intContinente = CInt(vecNuevaMision(i))
                End If
            Case 1
                'Color a destruir
                If IsNull(vecNuevaMision(i)) Then
                    GvecObjetivos(indice).intColorADestruir = 0
                Else
                    GvecObjetivos(indice).intColorADestruir = CInt(vecNuevaMision(i))
                End If
                
                If GvecObjetivos(indice).intColorADestruir > 0 Then
                    GvecObjetivos(indice).esDestruir = True
                Else
                    GvecObjetivos(indice).esDestruir = False
                End If
            Case 2
                'Cantidad de paises
                If IsNull(vecNuevaMision(i)) Then
                    GvecObjetivos(indice).intCantidadPaises = 0
                Else
                    GvecObjetivos(indice).intCantidadPaises = CInt(vecNuevaMision(i))
                End If
            Case 3
                'Limitrofes
                If IsNull(vecNuevaMision(i)) Then
                    GvecObjetivos(indice).sonLimitrofes = False
                Else
                    If Trim(UCase(vecNuevaMision(i))) = "S" Then
                        GvecObjetivos(indice).sonLimitrofes = True
                    Else
                        GvecObjetivos(indice).sonLimitrofes = False
                    End If
                End If
        End Select
            
    Next i
    
    GstrMision = vecNuevaMision(0)
    
    CalcularValorPais

    Exit Sub
ErrorHandle:
    ReportErr "cMostrarMision", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cInicioTurno(intColorTurno As Integer, intTimerTurno As Integer, blnResincronizacion As Boolean)
    On Error GoTo ErrorHandle
    Dim byPais As Byte
    
    GintColorActual = intColorTurno
    
    'Si el turno es mio
    If intColorTurno = GintMiColor Then
    
        'Marca como que no conquisto ningún pais
        GblnConquistoAlgunPais = False
        
        'Reduce el umbral a la mitad
        GsngUmbralAtaque = GsngUmbralAtaqueOriginal / GsngFactorReduccionUmbral
    
        'Muestra el icono del systray con lucecita prendida (Activo)
        frmPrincipal.imgSysTray.Picture = frmPrincipal.imgLstSysTrayActivo.ListImages(GintJvColor).Picture
        SysTrayChangeIcon frmPrincipal.hwnd, frmPrincipal.imgSysTray
        
        
        'Solo limpia las tropas fijas y cambia el estado si no se trata
        'de una resincronizacion
        If Not blnResincronizacion Then
            'Limpia las cantidades de tropas fijas de cada pais
            For byPais = 1 To UBound(GvecPaises)
                GvecPaises(byPais).intTropasFijas = 0
            Next byPais
            
            'Cambia el estado
            If GintTipoRonda = trInicio Or GintTipoRonda = trRecuento Then
                ActualizarEstadoCliente eveInicioTurnoRecuento
            Else
                ActualizarEstadoCliente eveInicioTurnoAccion
            End If
        End If
        
        'Si el turno es de recuento ejecuta CanjearTarjetas
        If GintTipoRonda = trRecuento Then
            vCanjearTarjetas
        Else
            EfectuarAccion
        End If
        
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cInicioTurno", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cActualizarTipoRonda(intTipoRonda As enuTipoRonda)
    On Error GoTo ErrorHandle
    
    GintTipoRonda = intTipoRonda
    
    Exit Sub
ErrorHandle:
    ReportErr "cActualizarTipoRonda", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cAgregarTropas(intPais As Byte, intCantidad As Integer)
    'Agrega tropas a un pais determinado
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgAgregarTropas, intPais, intCantidad)
        
    Exit Sub
ErrorHandle:
    ReportErr "cAgregarTropas", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cFinTurno()
    'Informa al servidor el fin voluntario del turno
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgFinTurno)
    
    Exit Sub
ErrorHandle:
    ReportErr "cFinTurno", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cConfirmarFinTurno(blnExpiroTimer As Boolean)
    On Error GoTo ErrorHandle
    
    'Desactiva el timer del timeout
    frmPrincipal.tmr_TimeOut.Interval = 0
    
    'Cambia el estado a esperando turno
    ActualizarEstadoCliente eveAckFinTurno
    
    'Muestra el icono del systray normal (lucecita apagada)
    frmPrincipal.imgSysTray.Picture = frmPrincipal.imgLstSysTrayOK.ListImages(GintJvColor).Picture
    SysTrayChangeIcon frmPrincipal.hwnd, frmPrincipal.imgSysTray
    
    If blnExpiroTimer Then
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cConfirmarFinTurno", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cTropasDisponibles(vecMensaje() As String)
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim intTropasDisponibles As Integer
    Dim j As Integer
    Dim blnExisteNoLibre As Boolean
    
    blnExisteNoLibre = False
    intTropasDisponibles = 0
    For i = 1 To UBound(vecMensaje)
        GvecJugadores(CInt(vecMensaje(0))).vecDetalleTropasDisponibles(i - 1) = CInt(vecMensaje(i))
        intTropasDisponibles = intTropasDisponibles + CInt(vecMensaje(i))
        If i > 1 And CInt(vecMensaje(i)) > 0 Then
            blnExisteNoLibre = True
        End If
    Next i
    
    GvecJugadores(CInt(vecMensaje(0))).intTropasDisponibles = intTropasDisponibles
    
    Exit Sub
ErrorHandle:
    ReportErr "cTropasDisponibles", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

'### Hacer enumeracion con los distintos resultados
Public Sub cAckAgregarTropas(intResultado As Integer)
    
    On Error GoTo ErrorHandle
    
    EfectuarAccion
    
    Exit Sub
ErrorHandle:
    ReportErr "cAckAgregarTropas", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cAckAtacar(intDadoDesde1 As Integer, intDadoDesde2 As Integer, intDadoDesde3 As Integer, _
                      intDadoHasta1 As Integer, intDadoHasta2 As Integer, intDadoHasta3 As Integer, _
                      byPaisDesde As Byte, intColorDesde As Integer, intCantDesde As Integer, _
                      byPaisHasta As Byte, intColorHasta As Integer, intCantHasta As Integer)

    On Error GoTo ErrorHandle
    Dim vecCandidatos(0 To 2) As Byte
    Dim byPaisDestino As Byte
    Dim intCantTropasMover As Integer
    Dim blnEfectuarAccion As Boolean
    
    blnEfectuarAccion = False
    
    'Actualiza los dos paises
    cActualizarPais byPaisDesde, intColorDesde, intCantDesde, orAtaque, 0
    cActualizarPais byPaisHasta, intColorHasta, intCantHasta, orAtaque, 0
    
    'Si hubo conquista
    If intColorDesde = GintMiColor And intColorHasta = GintMiColor Then
        
        'Devuelve el umbral a su valor original
        GblnConquistoAlgunPais = True
        GsngUmbralAtaque = GsngUmbralAtaqueOriginal
        
        'Determina la cantidad de tropas a pasar
        
        'Si hay tropas para mover
        If intCantDesde > 1 Then
            vecCandidatos(1) = byPaisDesde
            vecCandidatos(2) = byPaisHasta
            
            'Determina si puede mover 1 o 2
            If intCantDesde = 2 Then
                'Hay una tropa para mover
                GvecPaises(byPaisDesde).intCantidadFichas = GvecPaises(byPaisDesde).intCantidadFichas - 1
                byPaisDestino = vObtenerPaisObjetivoAgregar(vecCandidatos)
                If byPaisDestino = byPaisDesde Then
                    'No hay que mover nada
                    blnEfectuarAccion = True
                Else
                    'Mueve la tropa
                    cMover byPaisDesde, byPaisDestino, 1, tmConquista
                End If
            Else
                intCantTropasMover = 0
                'Hay dos tropas para mover
                GvecPaises(byPaisDesde).intCantidadFichas = GvecPaises(byPaisDesde).intCantidadFichas - 2
                
                'Primer tropa
                byPaisDestino = vObtenerPaisObjetivoAgregar(vecCandidatos)
                If byPaisDestino = byPaisDesde Then
                    'No hay que mover nada
                Else
                    'Mueve la tropa
                    intCantTropasMover = intCantTropasMover + 1
                End If
                
                'Segunda tropa
                byPaisDestino = vObtenerPaisObjetivoAgregar(vecCandidatos)
                If byPaisDestino = byPaisDesde Then
                    'No hay que mover nada
                Else
                    'Mueve la tropa
                    intCantTropasMover = intCantTropasMover + 1
                End If
                
                If intCantTropasMover > 0 Then
                    cMover byPaisDesde, byPaisDestino, intCantTropasMover, tmConquista
                Else
                    blnEfectuarAccion = True
                End If
            End If
        Else
            blnEfectuarAccion = True
        End If
        
    Else
        blnEfectuarAccion = True
    End If
    
    If blnEfectuarAccion Then
        EfectuarAccion
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cAckAtacar", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cAckMover(byPaisDesde As Byte, intColorDesde As Integer, intCantDesde As Integer, _
                     byPaisHasta As Byte, intColorHasta As Integer, intCantHasta As Integer, _
                     intTipoMovimiento As enuTipoMovimiento, intCantidadTropas As Integer)
    On Error GoTo ErrorHandle
    
    'Cambia el estado del jugador que hizo el movimiento
    'Solo si el tipo de movimiento fue Movimiento
    If intTipoMovimiento = tmMovimiento Then
        ActualizarEstadoCliente eveAckMoverTropa
    
        'Actualiza los dos paises
        cActualizarPais byPaisDesde, intColorDesde, intCantDesde, orMovimiento, GvecPaises(byPaisDesde).intTropasFijas
        cActualizarPais byPaisHasta, intColorHasta, intCantHasta, orMovimiento, GvecPaises(byPaisHasta).intTropasFijas + intCantidadTropas
    
    Else
        'Actualiza los dos paises
        cActualizarPais byPaisDesde, intColorDesde, intCantDesde, orMovimiento, GvecPaises(byPaisDesde).intTropasFijas
        cActualizarPais byPaisHasta, intColorHasta, intCantHasta, orMovimiento, GvecPaises(byPaisHasta).intTropasFijas
    End If
    
    
    'Vuelve a atacar o a mover o pasa a tomar tarjeta
    EfectuarAccion
    
    Exit Sub
ErrorHandle:
    ReportErr "cAckMover", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMostrarTarjeta(byPais As Byte, byFigura As Byte, blnCobrada As Boolean)
    'Actualiza mis tarjetas y llama a Efectuar Accion
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    'Cambia el estado del cliente
    ActualizarEstadoCliente eveTarjetaTomada
    
    'Guarda la nueva tarjeta en la primer posición libre
    For i = 1 To UBound(GvecTarjetas)
        If GvecTarjetas(i).byPais = 0 Then
            GvecTarjetas(i).byPais = byPais
            GvecTarjetas(i).byFigura = byFigura
            GvecTarjetas(i).blCobrada = blnCobrada
            Exit For
        End If
    Next i
    
    EfectuarAccion
    
    Exit Sub
ErrorHandle:
    ReportErr "cMostrarTarjeta", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cActualizarTarjetasJugador(intColor As Integer, intCantTarjetas As Integer)
    'Actualiza la cantidad de tarjetas del jugador especificado
    On Error GoTo ErrorHandle
    
    GvecJugadores(intColor).intCantidadTarjetas = intCantTarjetas
    
    Exit Sub
ErrorHandle:
    ReportErr "cActualizarTarjetasJugadores", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cAckCobrarTarjeta(byPais As Byte, intColor As Integer, intCantidad As Integer)
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    'Si quien cobra la tarjeta soy yo
    If intColor = GintMiColor Then
        'Cambia el estado
        ActualizarEstadoCliente eveTarjetaCobrada
        
        'Busca la tarjeta que corresponda con el pais y la marca como cobrada
        For i = 1 To UBound(GvecTarjetas)
            If GvecTarjetas(i).byPais = byPais Then
                GvecTarjetas(i).blCobrada = True
                Exit For
            End If
        Next i
        
        EfectuarAccion
        
    End If
    
    'Actualiza el mapa
    cActualizarPais byPais, intColor, intCantidad, orCobroTarjeta, (intCantidad - GvecPaises(byPais).intCantidadFichas) + GvecPaises(byPais).intTropasFijas
    
    Exit Sub
ErrorHandle:
    ReportErr "cCobrarTarjeta", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cCanjearTarjetas(intTarjeta1 As Integer, intTarjeta2 As Integer, intTarjeta3 As Integer)
    On Error GoTo ErrorHandle

    EnviarMensaje ArmarMensajeParam(msgCanjeTarjeta, GvecTarjetas(intTarjeta1).byPais, _
                                                     GvecTarjetas(intTarjeta2).byPais, _
                                                     GvecTarjetas(intTarjeta3).byPais)
    Exit Sub
ErrorHandle:
    ReportErr "cCanjearTarjetas", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub


Public Sub cEnviarIAA()
    'Envia al servidor el mensaje que indica que el cliente está vivo
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgIAA)
        
    Exit Sub
ErrorHandle:
    ReportErr "cEnviarIAA", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cDesconectar()
    'Desconecta al cliente sin bajar el servidor
    On Error GoTo ErrorHandle
                
    frmPrincipal.wskVP.Close
    
    'mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panEstado).Text = ""
    GintMiColor = 0
    ActualizarEstadoCliente eveCierreConexion
    
    Exit Sub
ErrorHandle:
    ReportErr "cDesconectar", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cResincronizar()
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    '###Limpia las tarjetas
    For i = LBound(GvecTarjetas) To UBound(GvecTarjetas)
        GvecTarjetas(i).byPais = 0
    Next i
    
    EnviarMensaje ArmarMensajeParam(msgResincronizacion)
    
    Exit Sub
ErrorHandle:
    ReportErr "cResincronizar", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cAtacar(byPaisDesde As Byte, byPaisHasta As Byte)
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgAtaque, byPaisDesde, byPaisHasta)
    
    Exit Sub
ErrorHandle:
    ReportErr "cAtacar", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMostrarError(strMensajeError As String, intCodigoError As enuErrores)
    'Muestra en pantalla el mensaje de error recibido del servidor
    On Error GoTo ErrorHandle
    
    frmPrincipal.txtErrores.Text = strMensajeError
    
    Select Case intCodigoError
        Case enuErrores.errNoTomarNoConquista1, enuErrores.errNoTomarNoConquista2
            vCobrarTarjeta
        Case enuErrores.errNoTomarNoMasTarjetas
            vCobrarTarjeta
        Case Else
            cResincronizar
    End Select
        
    Exit Sub
ErrorHandle:
    ReportErr "cMostrarError", "mdlInterfase", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cTomarTarjeta()
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgPedidoTarjeta)
    
    Exit Sub
ErrorHandle:
    ReportErr "cTomarTarjeta", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cCobrarTarjeta(intTarjetaSeleccionada As Integer)
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    'Envia el mensaje
    EnviarMensaje ArmarMensajeParam(msgCobroTarjeta, GvecTarjetas(intTarjetaSeleccionada).byPais)
    
    Exit Sub
ErrorHandle:
    ReportErr "cCobrarTarjeta", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cAckCanjearTarjeta(byPais1 As Byte, byPais2 As Byte, byPais3 As Byte)
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    'Elimina las tarjetas canjeadas
    For i = 1 To UBound(GvecTarjetas)
        If GvecTarjetas(i).byPais = byPais1 Or GvecTarjetas(i).byPais = byPais2 Or GvecTarjetas(i).byPais = byPais3 Then
            GvecTarjetas(i).byPais = 0
        End If
    Next i
    
    EfectuarAccion
    
    Exit Sub
ErrorHandle:
    ReportErr "cAckCanjearTarjeta", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMover(byPaisDesde As Byte, byPaisHasta As Byte, intCantidadFichas As Integer, intTipoMovimiento As enuTipoMovimiento)
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgMovimiento, byPaisDesde, byPaisHasta, intCantidadFichas, intTipoMovimiento)
    
    Exit Sub
ErrorHandle:
    ReportErr "cMover", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cActualizarEstadoTurno(intEstadoTurno As enuEstadoCli)
    'Actualiza el estado del cliente (forzado por la resincronizacion)
    On Error GoTo ErrorHandle
    
    GEstadoCliente = intEstadoTurno
    frmPrincipal.lblEstado.Caption = GvecEstadoCliente(GEstadoCliente)
    'mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panEstado).Text = GvecEstadoCliente(GEstadoCliente)

    '###E Agregar a las desactivaciones de opciones de menu
''    If GEstadoCliente > estEsperandoTurno And GEstadoCliente <> estInconsistente Then
''        GblnMapaHabilitado = True
''    Else
''        GblnMapaHabilitado = False
''    End If
    
    '### Matriz de controles (habilitar/deshabilitar opciones)
    'ActualizarControles
    
    Exit Sub
ErrorHandle:
    ReportErr "cActualizarEstadoTurno", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMisionCumplida()
    On Error GoTo ErrorHandle
    Dim strMensaje As String
    
    '###E
    'Actualiza el estado del cliente
    ActualizarEstadoCliente eveMisionCumplida
    
    Exit Sub
ErrorHandle:
    ReportErr "cMisionCumplida", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub









'-------------------------------------------------------
'-         GENERALES                                   -
'-------------------------------------------------------

Private Sub CargarRegistroEstado(estadoOrigen As enuEstadoCli, eventoOrigen As enuEventosCli, estadoDestino As enuEstadoCli)
    'Subrutina utilizada para facilitar la carga de la matriz de estados
    On Error GoTo ErrorHandle
    
    GMatrizEstados(estadoOrigen, eventoOrigen) = estadoDestino
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarRegistroEstado", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CargarMatrizEstados()
    'Carga la matriz de estados del cliente
    On Error GoTo ErrorHandle
    
    CargarRegistroEstado estDesconectado, eveConfirmacionAdm, estConectado
    CargarRegistroEstado estDesconectado, eveJugadoresConectados, estConectado
    CargarRegistroEstado estDesconectado, eveConfirmacionAlta, estValidado
    CargarRegistroEstado estDesconectado, eveInicioPartida, estEsperandoTurno
    CargarRegistroEstado estDesconectado, eveInicioTurnoRecuento, estInconsistente
    CargarRegistroEstado estDesconectado, eveInicioTurnoAccion, estInconsistente
    CargarRegistroEstado estDesconectado, eveAckFinTurno, estInconsistente
    CargarRegistroEstado estDesconectado, eveAckMoverTropa, estInconsistente
    CargarRegistroEstado estDesconectado, eveTarjetaTomada, estInconsistente
    CargarRegistroEstado estDesconectado, eveTarjetaCobrada, estInconsistente
    CargarRegistroEstado estDesconectado, eveMisionCumplida, estInconsistente
    CargarRegistroEstado estDesconectado, eveCierreConexion, estDesconectado
    CargarRegistroEstado estDesconectado, evePartidaPausada, estPartidaPausada
    CargarRegistroEstado estDesconectado, evePartidaContinuada, estInconsistente
    
    CargarRegistroEstado estConectado, eveConfirmacionAdm, estConectado
    CargarRegistroEstado estConectado, eveJugadoresConectados, estConectado
    CargarRegistroEstado estConectado, eveConfirmacionAlta, estValidado
    CargarRegistroEstado estConectado, eveInicioPartida, estEsperandoTurno
    CargarRegistroEstado estConectado, eveInicioTurnoRecuento, estInconsistente
    CargarRegistroEstado estConectado, eveInicioTurnoAccion, estInconsistente
    CargarRegistroEstado estConectado, eveAckFinTurno, estInconsistente
    CargarRegistroEstado estConectado, eveAckMoverTropa, estInconsistente
    CargarRegistroEstado estConectado, eveTarjetaTomada, estInconsistente
    CargarRegistroEstado estConectado, eveTarjetaCobrada, estInconsistente
    CargarRegistroEstado estConectado, eveMisionCumplida, estInconsistente
    CargarRegistroEstado estConectado, eveCierreConexion, estDesconectado
    CargarRegistroEstado estConectado, evePartidaPausada, estPartidaPausada
    CargarRegistroEstado estConectado, evePartidaContinuada, estInconsistente

    CargarRegistroEstado estValidado, eveConfirmacionAdm, estValidado
    CargarRegistroEstado estValidado, eveJugadoresConectados, estValidado
    CargarRegistroEstado estValidado, eveConfirmacionAlta, estValidado
    CargarRegistroEstado estValidado, eveInicioPartida, estEsperandoTurno
    CargarRegistroEstado estValidado, eveInicioTurnoRecuento, estInconsistente
    CargarRegistroEstado estValidado, eveInicioTurnoAccion, estInconsistente
    CargarRegistroEstado estValidado, eveAckFinTurno, estInconsistente
    CargarRegistroEstado estValidado, eveAckMoverTropa, estInconsistente
    CargarRegistroEstado estValidado, eveTarjetaTomada, estInconsistente
    CargarRegistroEstado estValidado, eveTarjetaCobrada, estInconsistente
    CargarRegistroEstado estValidado, eveMisionCumplida, estInconsistente
    CargarRegistroEstado estValidado, eveCierreConexion, estDesconectado
    CargarRegistroEstado estValidado, evePartidaPausada, estPartidaPausada
    CargarRegistroEstado estValidado, evePartidaContinuada, estInconsistente
    
    CargarRegistroEstado estEsperandoTurno, eveConfirmacionAdm, estEsperandoTurno
    CargarRegistroEstado estEsperandoTurno, eveJugadoresConectados, estEsperandoTurno
    CargarRegistroEstado estEsperandoTurno, eveConfirmacionAlta, estValidado
    CargarRegistroEstado estEsperandoTurno, eveInicioPartida, estEsperandoTurno
    CargarRegistroEstado estEsperandoTurno, eveInicioTurnoRecuento, estAgregando
    CargarRegistroEstado estEsperandoTurno, eveInicioTurnoAccion, estAtacando
    CargarRegistroEstado estEsperandoTurno, eveAckFinTurno, estInconsistente
    CargarRegistroEstado estEsperandoTurno, eveAckMoverTropa, estEsperandoTurno
    CargarRegistroEstado estEsperandoTurno, eveTarjetaTomada, estInconsistente
    CargarRegistroEstado estEsperandoTurno, eveTarjetaCobrada, estInconsistente
    CargarRegistroEstado estEsperandoTurno, eveMisionCumplida, estPartidaFinalizada
    CargarRegistroEstado estEsperandoTurno, eveCierreConexion, estDesconectado
    CargarRegistroEstado estEsperandoTurno, evePartidaPausada, estPartidaPausada
    CargarRegistroEstado estEsperandoTurno, evePartidaContinuada, estInconsistente
    
    CargarRegistroEstado estAgregando, eveConfirmacionAdm, estAgregando
    CargarRegistroEstado estAgregando, eveJugadoresConectados, estAgregando
    CargarRegistroEstado estAgregando, eveConfirmacionAlta, estInconsistente
    CargarRegistroEstado estAgregando, eveInicioPartida, estInconsistente
    CargarRegistroEstado estAgregando, eveInicioTurnoRecuento, estInconsistente
    CargarRegistroEstado estAgregando, eveInicioTurnoAccion, estInconsistente
    CargarRegistroEstado estAgregando, eveAckFinTurno, estEsperandoTurno
    CargarRegistroEstado estAgregando, eveAckMoverTropa, estInconsistente
    CargarRegistroEstado estAgregando, eveTarjetaTomada, estInconsistente
    CargarRegistroEstado estAgregando, eveTarjetaCobrada, estInconsistente
    CargarRegistroEstado estAgregando, eveMisionCumplida, estPartidaFinalizada
    CargarRegistroEstado estAgregando, eveCierreConexion, estDesconectado
    CargarRegistroEstado estAgregando, evePartidaPausada, estPartidaPausada
    CargarRegistroEstado estAgregando, evePartidaContinuada, estInconsistente
    
    CargarRegistroEstado estAtacando, eveConfirmacionAdm, estAtacando
    CargarRegistroEstado estAtacando, eveJugadoresConectados, estAtacando
    CargarRegistroEstado estAtacando, eveConfirmacionAlta, estInconsistente
    CargarRegistroEstado estAtacando, eveInicioPartida, estInconsistente
    CargarRegistroEstado estAtacando, eveInicioTurnoRecuento, estInconsistente
    CargarRegistroEstado estAtacando, eveInicioTurnoAccion, estInconsistente
    CargarRegistroEstado estAtacando, eveAckFinTurno, estEsperandoTurno
    CargarRegistroEstado estAtacando, eveAckMoverTropa, estMoviendo
    CargarRegistroEstado estAtacando, eveTarjetaTomada, estTarjetaTomada
    CargarRegistroEstado estAtacando, eveTarjetaCobrada, estTarjetaCobrada
    CargarRegistroEstado estAtacando, eveMisionCumplida, estPartidaFinalizada
    CargarRegistroEstado estAtacando, eveCierreConexion, estDesconectado
    CargarRegistroEstado estAtacando, evePartidaPausada, estPartidaPausada
    CargarRegistroEstado estAtacando, evePartidaContinuada, estInconsistente

    CargarRegistroEstado estMoviendo, eveConfirmacionAdm, estMoviendo
    CargarRegistroEstado estMoviendo, eveJugadoresConectados, estMoviendo
    CargarRegistroEstado estMoviendo, eveConfirmacionAlta, estInconsistente
    CargarRegistroEstado estMoviendo, eveInicioPartida, estInconsistente
    CargarRegistroEstado estMoviendo, eveInicioTurnoRecuento, estInconsistente
    CargarRegistroEstado estMoviendo, eveInicioTurnoAccion, estInconsistente
    CargarRegistroEstado estMoviendo, eveAckFinTurno, estEsperandoTurno
    CargarRegistroEstado estMoviendo, eveAckMoverTropa, estMoviendo
    CargarRegistroEstado estMoviendo, eveTarjetaTomada, estTarjetaTomada
    CargarRegistroEstado estMoviendo, eveTarjetaCobrada, estTarjetaCobrada
    CargarRegistroEstado estMoviendo, eveMisionCumplida, estPartidaFinalizada
    CargarRegistroEstado estMoviendo, eveCierreConexion, estDesconectado
    CargarRegistroEstado estMoviendo, evePartidaPausada, estPartidaPausada
    CargarRegistroEstado estMoviendo, evePartidaContinuada, estInconsistente

    CargarRegistroEstado estTarjetaCobrada, eveConfirmacionAdm, estTarjetaCobrada
    CargarRegistroEstado estTarjetaCobrada, eveJugadoresConectados, estTarjetaCobrada
    CargarRegistroEstado estTarjetaCobrada, eveConfirmacionAlta, estInconsistente
    CargarRegistroEstado estTarjetaCobrada, eveInicioPartida, estInconsistente
    CargarRegistroEstado estTarjetaCobrada, eveInicioTurnoRecuento, estInconsistente
    CargarRegistroEstado estTarjetaCobrada, eveInicioTurnoAccion, estInconsistente
    CargarRegistroEstado estTarjetaCobrada, eveAckFinTurno, estEsperandoTurno
    CargarRegistroEstado estTarjetaCobrada, eveAckMoverTropa, estInconsistente
    CargarRegistroEstado estTarjetaCobrada, eveTarjetaTomada, estTarjetaCobradaTomada
    CargarRegistroEstado estTarjetaCobrada, eveTarjetaCobrada, estTarjetaCobrada
    CargarRegistroEstado estTarjetaCobrada, eveMisionCumplida, estPartidaFinalizada
    CargarRegistroEstado estTarjetaCobrada, eveCierreConexion, estDesconectado
    CargarRegistroEstado estTarjetaCobrada, evePartidaPausada, estPartidaPausada
    CargarRegistroEstado estTarjetaCobrada, evePartidaContinuada, estInconsistente

    CargarRegistroEstado estTarjetaTomada, eveConfirmacionAdm, estTarjetaTomada
    CargarRegistroEstado estTarjetaTomada, eveJugadoresConectados, estTarjetaTomada
    CargarRegistroEstado estTarjetaTomada, eveConfirmacionAlta, estInconsistente
    CargarRegistroEstado estTarjetaTomada, eveInicioPartida, estInconsistente
    CargarRegistroEstado estTarjetaTomada, eveInicioTurnoRecuento, estInconsistente
    CargarRegistroEstado estTarjetaTomada, eveInicioTurnoAccion, estInconsistente
    CargarRegistroEstado estTarjetaTomada, eveAckFinTurno, estEsperandoTurno
    CargarRegistroEstado estTarjetaTomada, eveAckMoverTropa, estInconsistente
    CargarRegistroEstado estTarjetaTomada, eveTarjetaTomada, estInconsistente
    CargarRegistroEstado estTarjetaTomada, eveTarjetaCobrada, estTarjetaCobradaTomada
    CargarRegistroEstado estTarjetaTomada, eveMisionCumplida, estPartidaFinalizada
    CargarRegistroEstado estTarjetaTomada, eveCierreConexion, estDesconectado
    CargarRegistroEstado estTarjetaTomada, evePartidaPausada, estPartidaPausada
    CargarRegistroEstado estTarjetaTomada, evePartidaContinuada, estInconsistente

    CargarRegistroEstado estTarjetaCobradaTomada, eveConfirmacionAdm, estTarjetaTomada
    CargarRegistroEstado estTarjetaCobradaTomada, eveJugadoresConectados, estTarjetaTomada
    CargarRegistroEstado estTarjetaCobradaTomada, eveConfirmacionAlta, estInconsistente
    CargarRegistroEstado estTarjetaCobradaTomada, eveInicioPartida, estInconsistente
    CargarRegistroEstado estTarjetaCobradaTomada, eveInicioTurnoRecuento, estInconsistente
    CargarRegistroEstado estTarjetaCobradaTomada, eveInicioTurnoAccion, estInconsistente
    CargarRegistroEstado estTarjetaCobradaTomada, eveAckFinTurno, estEsperandoTurno
    CargarRegistroEstado estTarjetaCobradaTomada, eveAckMoverTropa, estInconsistente
    CargarRegistroEstado estTarjetaCobradaTomada, eveTarjetaTomada, estInconsistente
    CargarRegistroEstado estTarjetaCobradaTomada, eveTarjetaCobrada, estTarjetaCobradaTomada
    CargarRegistroEstado estTarjetaCobradaTomada, eveMisionCumplida, estPartidaFinalizada
    CargarRegistroEstado estTarjetaCobradaTomada, eveCierreConexion, estDesconectado
    CargarRegistroEstado estTarjetaCobradaTomada, evePartidaPausada, estPartidaPausada
    CargarRegistroEstado estTarjetaCobradaTomada, evePartidaContinuada, estInconsistente

    CargarRegistroEstado estPartidaPausada, eveConfirmacionAdm, estPartidaPausada
    CargarRegistroEstado estPartidaPausada, eveJugadoresConectados, estPartidaPausada
    CargarRegistroEstado estPartidaPausada, eveConfirmacionAlta, estPartidaPausada
    CargarRegistroEstado estPartidaPausada, eveInicioPartida, estInconsistente
    CargarRegistroEstado estPartidaPausada, eveInicioTurnoRecuento, estInconsistente
    CargarRegistroEstado estPartidaPausada, eveInicioTurnoAccion, estInconsistente
    CargarRegistroEstado estPartidaPausada, eveAckFinTurno, estInconsistente
    CargarRegistroEstado estPartidaPausada, eveAckMoverTropa, estInconsistente
    CargarRegistroEstado estPartidaPausada, eveTarjetaTomada, estInconsistente
    CargarRegistroEstado estPartidaPausada, eveTarjetaCobrada, estInconsistente
    CargarRegistroEstado estPartidaPausada, eveMisionCumplida, estInconsistente
    CargarRegistroEstado estPartidaPausada, eveCierreConexion, estDesconectado
    CargarRegistroEstado estPartidaPausada, evePartidaPausada, estInconsistente
    CargarRegistroEstado estPartidaPausada, evePartidaContinuada, estInconsistente 'Será dinamico

    CargarRegistroEstado estPartidaFinalizada, eveConfirmacionAdm, estPartidaFinalizada
    CargarRegistroEstado estPartidaFinalizada, eveJugadoresConectados, estPartidaFinalizada
    CargarRegistroEstado estPartidaFinalizada, eveConfirmacionAlta, estInconsistente
    CargarRegistroEstado estPartidaFinalizada, eveInicioPartida, estInconsistente
    CargarRegistroEstado estPartidaFinalizada, eveInicioTurnoRecuento, estInconsistente
    CargarRegistroEstado estPartidaFinalizada, eveInicioTurnoAccion, estInconsistente
    CargarRegistroEstado estPartidaFinalizada, eveAckFinTurno, estInconsistente
    CargarRegistroEstado estPartidaFinalizada, eveAckMoverTropa, estInconsistente
    CargarRegistroEstado estPartidaFinalizada, eveTarjetaTomada, estInconsistente
    CargarRegistroEstado estPartidaFinalizada, eveTarjetaCobrada, estInconsistente
    CargarRegistroEstado estPartidaFinalizada, eveMisionCumplida, estPartidaFinalizada
    CargarRegistroEstado estPartidaFinalizada, eveCierreConexion, estDesconectado
    CargarRegistroEstado estPartidaFinalizada, evePartidaPausada, estPartidaPausada
    CargarRegistroEstado estPartidaFinalizada, evePartidaContinuada, estInconsistente
    
    CargarRegistroEstado estInconsistente, eveConfirmacionAdm, estInconsistente
    CargarRegistroEstado estInconsistente, eveJugadoresConectados, estInconsistente
    CargarRegistroEstado estInconsistente, eveConfirmacionAlta, estInconsistente
    CargarRegistroEstado estInconsistente, eveInicioPartida, estInconsistente
    CargarRegistroEstado estInconsistente, eveInicioTurnoRecuento, estInconsistente
    CargarRegistroEstado estInconsistente, eveInicioTurnoAccion, estInconsistente
    CargarRegistroEstado estInconsistente, eveAckFinTurno, estInconsistente
    CargarRegistroEstado estInconsistente, eveAckMoverTropa, estInconsistente
    CargarRegistroEstado estInconsistente, eveTarjetaTomada, estInconsistente
    CargarRegistroEstado estInconsistente, eveTarjetaCobrada, estInconsistente
    CargarRegistroEstado estInconsistente, eveMisionCumplida, estInconsistente
    CargarRegistroEstado estInconsistente, eveCierreConexion, estDesconectado
    CargarRegistroEstado estInconsistente, evePartidaPausada, estInconsistente
    CargarRegistroEstado estInconsistente, evePartidaContinuada, estInconsistente

    Exit Sub
ErrorHandle:
    ReportErr "CargarMatrizEstados", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub DescripcionEstadosCliente()
    'Carga un vector con las descripciones de los estados del cliente
    On Error GoTo ErrorHandle
    
    GvecEstadoCliente(enuEstadoCli.estDesconectado) = "Desconectado"
    GvecEstadoCliente(enuEstadoCli.estConectado) = "Conectado"
    GvecEstadoCliente(enuEstadoCli.estValidado) = "Validado"
    GvecEstadoCliente(enuEstadoCli.estEsperandoTurno) = "Esperando Turno"
    GvecEstadoCliente(enuEstadoCli.estAgregando) = "Agregando Tropas"
    GvecEstadoCliente(enuEstadoCli.estAtacando) = "Atacando"
    GvecEstadoCliente(enuEstadoCli.estMoviendo) = "Moviendo Tropas"
    GvecEstadoCliente(enuEstadoCli.estTarjetaTomada) = "Tarjeta Tomada"
    GvecEstadoCliente(enuEstadoCli.estPartidaPausada) = "Partida Pausada"
    GvecEstadoCliente(enuEstadoCli.estPartidaFinalizada) = "Partida Finalizada"
    GvecEstadoCliente(enuEstadoCli.estInconsistente) = "Estado Inconsistente"
    
    Exit Sub
ErrorHandle:
        ReportErr "DescripcionEstadosCliente", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub ActualizarEstadoCliente(NuevoEvento As enuEventosCli)
    'Cambia el estado del cliente de acuerdo a un evento pasado por parametro
    On Error GoTo ErrorHandle
    
    GEstadoCliente = GMatrizEstados(GEstadoCliente, NuevoEvento)
    frmPrincipal.lblEstado.Caption = GvecEstadoCliente(GEstadoCliente)
        
    Exit Sub
ErrorHandle:
    ReportErr "ActualizarEstadoCliente", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CargarPaisContinente(vecPaises() As String)
' Toma el mensaje con los paises y su continente enviado por el servidor,
' y lo guarda en el vector global de paises.
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    Dim vecPais As Variant
    
    ReDim GvecPaises(0)
    
    For i = LBound(vecPaises) To UBound(vecPaises)
        vecPais = Split(vecPaises(i), ",", 2)
        'El país que viene en el mensaje debe coincidir con el indice del vector.
        GvecPaises(UBound(GvecPaises)).intContinente = IIf(vecPais(1) = "", 0, vecPais(1))
        ReDim Preserve GvecPaises(UBound(GvecPaises) + 1)
    Next
    
    ReDim Preserve GvecPaises(UBound(GvecPaises) - 1)
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarPaisContinente", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Function CantidadPaises(intContinente As enuContinentes) As Integer
'Devuelve la cantidad de paises del continente especificado
    On Error GoTo ErrorHandle
    
    Dim intContador As Integer
    Dim i As Integer
    
    intContador = 0
    For i = LBound(GvecPaises) To UBound(GvecPaises)
        If GvecPaises(i).intContinente = intContinente Then
            intContador = intContador + 1
        End If
    Next
    
    CantidadPaises = intContador
    
    Exit Function
ErrorHandle:
    ReportErr "CantidadPaises", "mdlVP", Err.Description, Err.Number, Err.Source
End Function


Function ObtenerLineaComando(Optional MaxArgs)
   'Declara las variables.
   Dim C, LineaComando, LonLínComando, ArgIn, i, NúmArgs
   'Ver si MaxArgs está.
   If IsMissing(MaxArgs) Then MaxArgs = 10
   'Crea una matriz del tamaño correcto.
   ReDim ArgArray(MaxArgs)
   NúmArgs = 0: ArgIn = False
   'Obtiene los argumentos de la línea de comandos.
   LineaComando = Command()
   LonLínComando = Len(LineaComando)
   'Recorre la línea de comando carácter a carácter
   'a la vez.
   For i = 1 To LonLínComando
      C = Mid(LineaComando, i, 1)
      'Comprueba espacio o tabulación.
      If (C <> "|") Then
         'Ningún espacio o tabulación.
         'Comprueba si está en el argumento.
         If Not ArgIn Then
         'Empieza el nuevo argumento.
         'Comprueba para más argumentos.
            If NúmArgs = MaxArgs Then Exit For
            NúmArgs = NúmArgs + 1
   ArgIn = True
            End If
         'Concatenar el carácter al argumento actual.
         ArgArray(NúmArgs) = ArgArray(NúmArgs) & C
      Else
         'Encontró un espacio o tabulador.
         'Establece ArgIn a False.
         ArgIn = False
      End If
   Next i
   'Redimensiona la matriz lo suficiente para contener los argumentos.
   ReDim Preserve ArgArray(NúmArgs)
   'Devuelve la matriz en nombre de la función.
   ObtenerLineaComando = ArgArray()
End Function






'-------------------------------------------------------
'-         JUGADOR VIRTUAL                             -
'-------------------------------------------------------
Public Sub CalcularMisPaisesPorContinente()
    On Error GoTo ErrorHandle
    Dim intPais As Integer
    Dim intContinente As Integer
    
    'Cuenta la cantidad de paises mios que corresponden a ese continente
    For intContinente = 1 To UBound(GvecContinentes)
        GvecContinentes(intContinente).intCantidadPaisesMios = 0
    Next intContinente
    
    For intPais = 1 To UBound(GvecPaises)
        'Si el pais es mio
        If GvecPaises(intPais).intColor = GintMiColor Then
            GvecContinentes(GvecPaises(intPais).intContinente).intCantidadPaisesMios = GvecContinentes(GvecPaises(intPais).intContinente).intCantidadPaisesMios + 1
        End If
    Next intPais
    
    '### Calcula el bonue en base a la cantidad de paises
    For intContinente = 1 To UBound(GvecContinentes)
        GvecContinentes(intContinente).intBonus = Int(GvecContinentes(intContinente).intCantidadPaises / 2)
    Next intContinente
    
    Exit Sub
ErrorHandle:
    ReportErr "CalcularMisPaisesPorContinente", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub MostrarPorcentajeCompletitudContinentes()
    On Error GoTo ErrorHandle
    Dim intContinente
    
    For intContinente = 1 To UBound(GvecContinentes)
        frmPrincipal.lblPorcentajeContinente(intContinente).Caption = (GvecContinentes(intContinente).intCantidadPaisesMios / GvecContinentes(intContinente).intCantidadPaises) * 100
    Next intContinente
    
    Exit Sub
ErrorHandle:
    ReportErr "MostrarPorcentajeCompletitudContinentes", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Function CantidadPaisesJugador(intColor As Integer) As Integer
    'Calcula la cantidad de paises de un color dado
    On Error GoTo ErrorHandle
    
    Dim tmpCantPaises As Integer
    Dim byPais As Byte
    
    tmpCantPaises = 0
    
    For byPais = 1 To UBound(GvecPaises)
        If GvecPaises(byPais).intColor = intColor Then
            tmpCantPaises = tmpCantPaises + 1
        End If
    Next byPais
    
    CantidadPaisesJugador = tmpCantPaises
    
    Exit Function
ErrorHandle:
    ReportErr "CantidadPaisesJugador", "mdlVP", Err.Description, Err.Number, Err.Source
End Function

Public Function CalcularPorcentajeCompletitudMision() As Single
    'Devuelve el porcentaje de completitud de la mision y actualiza el porcentaje
    'de completidud de cada objetivo
    On Error GoTo ErrorHandle
    Dim intObjetivo As Integer
    Dim intCantPaisesFinMision As Integer 'Cuantos paises faltan para lograr el objetivo
    Dim intCantPaisesFinObjetivo As Integer 'Cuantos paises faltan para lograr la mision
    Dim intCantTotalPaises As Integer
    
    'Obtiene la cantidad maxima de paises
    intCantTotalPaises = UBound(GvecPaises)
    
    intCantPaisesFinMision = 0
    For intObjetivo = LBound(GvecObjetivos) To UBound(GvecObjetivos)
        If GvecObjetivos(intObjetivo).esDestruir Then
            'Objetivo destruir
            intCantPaisesFinObjetivo = CantidadPaisesJugador(GvecObjetivos(intObjetivo).intColorADestruir)
            GvecObjetivos(intObjetivo).sngPorcentajeTerminado = 1 - (intCantPaisesFinObjetivo / intCantTotalPaises)
            intCantPaisesFinMision = intCantPaisesFinMision + intCantPaisesFinObjetivo
        Else
            If GvecObjetivos(intObjetivo).intContinente <> 0 Then
                'Si el objetivo trata de conquistar paises en un continente
                intCantPaisesFinObjetivo = GvecObjetivos(intObjetivo).intCantidadPaises - GvecContinentes(GvecObjetivos(intObjetivo).intContinente).intCantidadPaisesMios
                GvecObjetivos(intObjetivo).sngPorcentajeTerminado = GvecContinentes(GvecObjetivos(intObjetivo).intContinente).intCantidadPaisesMios / GvecObjetivos(intObjetivo).intCantidadPaises
                intCantPaisesFinMision = intCantPaisesFinMision + intCantPaisesFinObjetivo
            Else
                'Si el objetivo trata de conquistar paises en cualquier parte
                intCantPaisesFinObjetivo = GvecObjetivos(intObjetivo).intCantidadPaises - CantidadPaisesJugador(GintMiColor)
                GvecObjetivos(intObjetivo).sngPorcentajeTerminado = CantidadPaisesJugador(GintMiColor) / GvecObjetivos(intObjetivo).intCantidadPaises
                intCantPaisesFinMision = intCantPaisesFinMision + intCantPaisesFinObjetivo
            End If
        End If
    Next
    
    CalcularPorcentajeCompletitudMision = 1 - (intCantPaisesFinMision / intCantTotalPaises)
    
    Exit Function
ErrorHandle:
    ReportErr "CalcularPorcentajeCompletitudMision", "mdlVP", Err.Description, Err.Number, Err.Source
End Function

Public Sub CalcularValorPais()
    'Calcula en el valor de cada pais de acuerdo a los algoritmos definidos:
    '   -Segun porcentaje de cumplido de la mision
    '   -Segun si pertenece o no a la mision
    '   -Segun el porcentaje conquistado de un continente
    '   -Si se tiene la tarjeta de ese pais (y no está cobrada)
    '   -Según si el pais cierra o no una frontera (proximamente)
    '
    On Error GoTo ErrorHandle
    Dim intPais As Integer
    Dim intContinente As enuContinentes
    Dim auxValor As Single
    Dim intObjetivo As Integer
    Dim intTarjeta As Integer
    Dim blnPerteneceAMision As Boolean
    Dim sngPorcentajeMision As Single
'    Dim intCantPaisesFinMision As Integer
    Dim intCantObjetivosMision As Integer
    
    For intPais = 1 To UBound(GvecPaises)
        auxValor = 0
        
        intContinente = GvecPaises(intPais).intContinente
        
        'Calculo del valor por pertenecer a un continente
        '###Por ahora cuadrática
        With GvecContinentes(intContinente)
            If .intCantidadPaises > 0 Then
                auxValor = auxValor + (.intBonus * GsngFactorBonusContinente * (.intCantidadPaisesMios / .intCantidadPaises) ^ 2)
            End If
        End With
        
        'Calculo del valor por tener tarjeta de ese pais
        'Si el pais a evaluar corresponde a una tarjeta no cobrada que no es mia
        For intTarjeta = 1 To UBound(GvecTarjetas)
            If Not GvecTarjetas(intTarjeta).blCobrada Then
                If GvecTarjetas(intTarjeta).byPais = intPais Then
                    If GvecPaises(intPais).intColor <> GintMiColor Then
                        auxValor = auxValor + GsngBonusPorTarjeta
                    End If
                End If
            End If
        Next
        
        'Calculo del valor por objetivos
        'Se calcula el porcentaje completado de la misión y se obtiene un valor
        sngPorcentajeMision = CalcularPorcentajeCompletitudMision
        blnPerteneceAMision = False
        'Por cada objetivo de la misión...
        intCantObjetivosMision = UBound(GvecObjetivos) - LBound(GvecObjetivos) + 1
        For intObjetivo = LBound(GvecObjetivos) To UBound(GvecObjetivos)
            'Objetivo destruir
            If GvecObjetivos(intObjetivo).esDestruir Then
                'Si el pais a evaluar pertenece al jugador a destruir...
                If GvecPaises(intPais).intColor = GvecObjetivos(intObjetivo).intColorADestruir Then
                    blnPerteneceAMision = True
                End If
            Else
                'Si el objetivo trata de conquistar un continente
                If GvecObjetivos(intObjetivo).intContinente <> 0 Then
                    'Si el pais pertenece a ese continente
                    If GvecPaises(intPais).intContinente = GvecObjetivos(intObjetivo).intContinente Then
                        blnPerteneceAMision = True
                        'intCantPaisesFinMision = intCantPaisesFinMision + GvecObjetivos(intObjetivo).intCantidadPaises - GvecContinentes(GvecPaises(intPais).intContinente).intCantidadPaisesMios
                    End If
                Else
                    blnPerteneceAMision = True
 '                   intCantPaisesFinMision = intCantPaisesFinMision + GvecObjetivos(intObjetivo).intCantidadPaises - CantidadPaisesJugador(GintMiColor)
                End If
            End If
            'Bonus por objetivo
            '###Si el objetivo ya se completó lo toma como completado el 100%
            'If blnPerteneceAMision Then
            '    auxValor = auxValor + ((GsngFactorObjetivo) * IIf(GvecObjetivos(intObjetivo).sngPorcentajeTerminado > 1, 1, GvecObjetivos(intObjetivo).sngPorcentajeTerminado) ^ 2)
            'End If
        Next
        
        'Bonus por mision
        If blnPerteneceAMision Then
            auxValor = auxValor + (GsngFactorMision * sngPorcentajeMision ^ 2)
        End If
        
        GvecPaises(intPais).sngValor = auxValor
        
        '### I
        '''frmPrincipal.lblValor(intPais).Caption = GvecPaises(intPais).sngValor
    
    Next intPais
    Exit Sub
ErrorHandle:
    ReportErr "CalcularValorPais", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CargarVectorConMisPaises(ByRef vecMisPaises() As Byte, Optional intContinente As enuContinentes = 0)
    'Carga el vector pasado por parametro con los paises pertenecientes al cliente
    'de un continente determinado (si no se pone devuelve todos)
    On Error GoTo ErrorHandle
    Dim intPais As Byte
    
    ReDim vecMisPaises(0 To 0)
    
    For intPais = 1 To UBound(GvecPaises)
        If GvecPaises(intPais).intColor = GintMiColor Then
            If intContinente = 0 Or GvecPaises(intPais).intContinente = intContinente Then
                'No redimensiona la matriz para el primer elemento
                If vecMisPaises(0) > 0 Then
                    ReDim Preserve vecMisPaises(UBound(vecMisPaises) + 1)
                End If
                vecMisPaises(UBound(vecMisPaises)) = intPais
            End If
        End If
    Next intPais
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarVectorConMisPaises", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub IniciarTimer()
    'Inicializa el timer
    On Error GoTo ErrorHandle
    
    GintSegRestantesTimeOut = GintSegTimeOut
    frmPrincipal.tmr_TimeOut.Interval = 1000
    
    Exit Sub
ErrorHandle:
    ReportErr "IniciarTimer", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub EfectuarAccion()
    'De acuerdo al estado del jugador efectua una accion (ataca, mueve, pone)
    On Error GoTo ErrorHandle
    
    '###Pausa
    Pausa 1000, False
    
    Select Case GEstadoCliente
        Case enuEstadoCli.estAgregando
            vAgregar
        Case enuEstadoCli.estAtacando
            vAtacar
        Case enuEstadoCli.estMoviendo
            vMover
        Case enuEstadoCli.estTarjetaTomada
            vCobrarTarjeta
        Case enuEstadoCli.estTarjetaCobradaTomada
            vCobrarTarjeta
        Case enuEstadoCli.estTarjetaCobrada
            vCobrarTarjeta
        Case enuEstadoCli.estInconsistente
            'cResincronizar
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "EfectuarAccion", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub vAgregar()
    On Error GoTo ErrorHandle
    Dim vecMisPaises() As Byte
    Dim byPaisObjetivo As Byte
    Dim intContinente As enuContinentes
    Dim sngPausa As Single
    
    IniciarPausa GintMsPausaAgregar, sngPausa
    
    'Si no tengo mas tropas disponibles pasa el turno
    If GvecJugadores(GintMiColor).intTropasDisponibles <= 0 Then
        cFinTurno
        Exit Sub
    End If
    
    'Chequea si tiene tropas de un solo continente
    For intContinente = coAfrica To coOceania
        If GvecJugadores(GintMiColor).vecDetalleTropasDisponibles(intContinente) > 0 Then
            CargarVectorConMisPaises vecMisPaises, intContinente
            byPaisObjetivo = vObtenerPaisObjetivoAgregar(vecMisPaises)
            
            IniciarTimer
            cAgregarTropas byPaisObjetivo, 1
            Exit Sub
        End If
    Next intContinente
    
    'Carga todos mis paises del mapa
    CargarVectorConMisPaises vecMisPaises
    byPaisObjetivo = vObtenerPaisObjetivoAgregar(vecMisPaises)
    
    FinalizarPausa GintMsPausaAgregar, sngPausa
    
    'Activa el timeOut
    IniciarTimer
    cAgregarTropas byPaisObjetivo, 1
    
    Exit Sub
ErrorHandle:
    ReportErr "vAgregar", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Function vObtenerPaisObjetivoAgregar(vecCandidatos() As Byte) As Byte
    'Recibe un vector con los paises a evaluar y selecciona uno de acuerdo
    'a los criterios establecidos
    On Error GoTo ErrorHandle
    Dim i As Integer, j As Integer
    Dim byPais As Byte
    Dim byLimitrofe As Byte
    Dim vecLimitrofes() As Byte
    Dim sngRiesgo As Single
    Dim sngOportunidad As Single
    Dim sngMaxOportunidad As Single
    Dim intCantFichasEnemigas As Integer
    Dim sngProbabilidad As Single
    Dim sngMaxValor As Single
    Dim sngValor As Single
    Dim byPaisObjetivo As Byte
    
    'Por cada pais candidato se calculan dos valores: Riesgo y Oportunidad
    'El Riesgo se entiende por el riesgo de perder un pais en manos de un enemigo
    '  limitrofe. Se calcula como la probabilidad de perder el pais ( se calcula
    '  con la sumatoria de las fichas de los paises limitrofes) y el valor asignado
    '  al pais objetivo.
    'Por Oportunidad se entiende a la posibilidad de conquistar un enemigo en base al
    '  país objetivo y se calcula como una función entre la probabilidad
    '  de conquistar el pais enemigo y el valor de dicho país.
    'Una vez obtenidos los valores se realiza una suma ponderada de los mismos en base al
    '  parametro que designa la agresividad del jugador
    
    'Inicializacion de variables
    sngMaxValor = 0
    
    'Por cada pais
    For i = LBound(vecCandidatos) To UBound(vecCandidatos)
        
        byPais = vecCandidatos(i)
        
        'Limpia las variables
        sngRiesgo = 0
        sngOportunidad = 0
        sngMaxOportunidad = 0
        intCantFichasEnemigas = 0
        
        'Obtiene la lista de los paises limitrofes
        vPaisesLimitrofes byPais, vecLimitrofes, False, True
        
        'Calculo del Riesgo
        '------------------
            'Cuenta la cantidad de fichas enemigas
            For j = LBound(vecLimitrofes) To UBound(vecLimitrofes)
                byLimitrofe = vecLimitrofes(j)
                If GvecPaises(byLimitrofe).intColor <> GintMiColor Then
                    'Se resta una porque uno siempre ataca con todas sus fichas menos 1
                    intCantFichasEnemigas = intCantFichasEnemigas + GvecPaises(byLimitrofe).intCantidadFichas - 1
                End If
            Next j
            'Suma una mas a la cantidad de fichas para realizar bien el calculo (que bolonqui no?)
            intCantFichasEnemigas = intCantFichasEnemigas + 1
            
            'Cálculo de la probabilidad de perder (1-Probabilidad de ganar)
            sngProbabilidad = vCalcularProbabilidadGanar(intCantFichasEnemigas, GvecPaises(byPais).intCantidadFichas)
            
            'Cálculo del riesgo
            sngRiesgo = GvecPaises(byPais).sngValor * sngProbabilidad
            
        'Calculo de la Oportunidad
        '-------------------------
            'Por cada pais limitrofe...
            For j = LBound(vecLimitrofes) To UBound(vecLimitrofes)
                byLimitrofe = vecLimitrofes(j)
                If GvecPaises(byLimitrofe).intColor <> GintMiColor Then
                    'Calcula la probabilidad de ganarlo
                    sngProbabilidad = vCalcularProbabilidadGanar(GvecPaises(byPais).intCantidadFichas, GvecPaises(byLimitrofe).intCantidadFichas)
                    
                    'Acumula el valor de la oportunidad
                    sngOportunidad = GvecPaises(byPais).sngValor * (1 - sngProbabilidad)
                    
                    If sngOportunidad > sngMaxOportunidad Then
                        sngMaxOportunidad = sngOportunidad
                    End If
                    
                End If
            Next j
            
        'Calculo del valor de evaluación de Recuento
        sngValor = (sngOportunidad * GsngActitud) + (sngRiesgo * (1 - GsngActitud))
        
        'Si es mayor al máximo se lo toma como objetivo potencial
        If sngValor >= sngMaxValor Then
            sngMaxValor = sngValor
            byPaisObjetivo = byPais
        End If
            
    Next i
    
    vObtenerPaisObjetivoAgregar = byPaisObjetivo
    
    '### Obtiene uno al azar
    'Randomize
    'vObtenerPaisObjetivoAgregar = vecCandidatos(Aleatorio(1, UBound(vecCandidatos)))
    
    Exit Function
ErrorHandle:
    ReportErr "vObtenerPaisObjetivoAgregar", "mdlVP", Err.Description, Err.Number, Err.Source
End Function

Public Sub vAtacar()
    On Error GoTo ErrorHandle
    Dim mtzE(0 To 50, 0 To 50) As Single
    Dim vecMisPaises() As Byte
    Dim byPaisOrigen As Byte
    Dim byPaisDestino As Byte
    Dim sngPausa As Single
    
    IniciarPausa GintMsPausaAtaque, sngPausa
    
    CargarVectorConMisPaises vecMisPaises
    vCalcularMatrizE mtzE, vecMisPaises
    
    If vObtenerPaisObjetivoAtaque(byPaisOrigen, byPaisDestino, mtzE) = True Then
        'Si algun objetivo supero el umbral de ataque, efectua el ataque
        FinalizarPausa GintMsPausaAtaque, sngPausa
        IniciarTimer
        cAtacar byPaisOrigen, byPaisDestino
    Else
        vMover
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "vAtacar", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub vMover()
    On Error GoTo ErrorHandle
    Dim byPais As Byte
    Dim vecLimitrofes() As Byte
    Dim byPaisDestino As Byte
    Dim sngPausa As Single
    
    IniciarPausa GintMsPausaMovimiento, sngPausa
    
    'Recorre todos los paises del jugador, identifica los que están en
    'condiciones de mover, les extrae las tropas de mas y luego las 'agrega'
    'al grupo de limitrofes o al mismo pais.
    byPaisDestino = 0
    For byPais = 1 To UBound(GvecPaises)
        If GvecPaises(byPais).intColor = GintMiColor Then
            'Si es mi pais...
            If GvecPaises(byPais).intCantidadFichas - GvecPaises(byPais).intTropasFijas > 1 Then
                'Si tiene tropas para mover...
                vPaisesLimitrofes byPais, vecLimitrofes
                If UBound(vecLimitrofes) > 0 Then
                    'Obtiene el pais al cual se debe mover la tropa
                    byPaisDestino = vObtenerPaisObjetivoAgregar(vecLimitrofes)
                    'Si el pais elegido es el propio pais de origen
                    If byPaisDestino = byPais Then
                        'Marca la tropa como fija
                        byPaisDestino = 0
                        GvecPaises(byPais).intTropasFijas = GvecPaises(byPais).intTropasFijas + 1
                    Else
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    
    If byPaisDestino > 0 Then
        'Si encontró un destino donde mover una tropa, efectua el movimiento
        'La tropa se marca como fija al recibir el Ack del movimiento
        FinalizarPausa GintMsPausaMovimiento, sngPausa
        'Inicia el timeout
        IniciarTimer
        cMover byPais, byPaisDestino, 1, tmMovimiento
    Else
        'En el momento que detecta que no tiene mas movimientos
        'para hacer, toma tarjeta.
        vTomarTarjeta
    End If
    Exit Sub
ErrorHandle:
    ReportErr "vMover", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub vTomarTarjeta()
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim blnPuedeTomarTarjeta As Boolean

    'Verifica que ya no tenga 5 tarjetas
    blnPuedeTomarTarjeta = False
    For i = 1 To UBound(GvecTarjetas)
        If GvecTarjetas(i).byPais = 0 Then
            blnPuedeTomarTarjeta = True
        End If
    Next i
    
    If Not blnPuedeTomarTarjeta Then
        'Si no puede tomar
        vCobrarTarjeta
    Else
        'Si puede tomar
        IniciarTimer
        cTomarTarjeta
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "vTomarTarjeta", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub vCobrarTarjeta()
    On Error GoTo ErrorHandle
    Dim intTarjeta As Integer
    Dim intTarjetaACobrar As Integer
    
    intTarjetaACobrar = -1
    
    'Recorre el vector de tarjetas en busca de una mia que
    'no este cobrada
    For intTarjeta = 1 To UBound(GvecTarjetas)
        'Si el pais de la tarjeta es mio...
        If GvecPaises(GvecTarjetas(intTarjeta).byPais).intColor = GintMiColor Then
            'Si la tarjeta no fue cobrada
            If Not GvecTarjetas(intTarjeta).blCobrada Then
                intTarjetaACobrar = intTarjeta
            End If
        End If
    Next
    
    'Si hay alguna tarjeta para cobrar
    If intTarjetaACobrar <> -1 Then
        'La cobra
        IniciarTimer
        cCobrarTarjeta intTarjetaACobrar
    Else
        'Pasa el turno
        cFinTurno
    End If

    Exit Sub
ErrorHandle:
    ReportErr "vCobrarTarjeta", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub vCanjearTarjetas()
    'Arma todos los tercetos posibles y selecciona la mejor combinación
    'Si no encuentra ninguno empieza a poner
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim intCantMaxTarjetas As Integer
    Dim vecTercetos(0 To 9) As typTercetos
    Dim intIndice As Integer
    Dim intValor As Integer
    Dim intMejorValor As Integer
    Dim intMejorIndice As Integer
    Dim blnCanjeValido As Boolean
    
    Dim intValorCobrada As Integer
    Dim intValorNoCobrada As Integer
    Dim intValorComodin As Integer
    
    intValorCobrada = 1
    intValorNoCobrada = 3
    intValorComodin = 3
    
    intCantMaxTarjetas = 5
    intIndice = 0
    intMejorValor = 0
    
    'Arma los tercetos posibles (de 5 agrupados de a 3 sin repeticiones)
    For i = 1 To intCantMaxTarjetas - 2
        For j = i + 1 To intCantMaxTarjetas - 1
            For k = j + 1 To intCantMaxTarjetas
               vecTercetos(intIndice).intTarjeta1 = i
               vecTercetos(intIndice).intTarjeta2 = j
               vecTercetos(intIndice).intTarjeta3 = k
               intIndice = intIndice + 1
            Next k
        Next j
    Next i
    
    'Asigna un valor a cada terceto
    For i = LBound(vecTercetos) To UBound(vecTercetos)
        'Si el terceto está completo...
        blnCanjeValido = False
        If GvecTarjetas(vecTercetos(i).intTarjeta1).byPais <> 0 And _
           GvecTarjetas(vecTercetos(i).intTarjeta2).byPais <> 0 And _
           GvecTarjetas(vecTercetos(i).intTarjeta3).byPais <> 0 Then
            
            'Busca si hay algun comodin
            If GvecTarjetas(vecTercetos(i).intTarjeta1).byFigura = figComodin _
            Or GvecTarjetas(vecTercetos(i).intTarjeta2).byFigura = figComodin _
            Or GvecTarjetas(vecTercetos(i).intTarjeta3).byFigura = figComodin Then
                'Si hay algun comodin el canje es válido
                blnCanjeValido = True
            Else
                'Si no hay comodin se fija si son todas iguales o todas distintas
                If GvecTarjetas(vecTercetos(i).intTarjeta1).byFigura = GvecTarjetas(vecTercetos(i).intTarjeta2).byFigura _
                And GvecTarjetas(vecTercetos(i).intTarjeta2).byFigura = GvecTarjetas(vecTercetos(i).intTarjeta3).byFigura Then
                    'Son las tres iguales
                    blnCanjeValido = True
                ElseIf GvecTarjetas(vecTercetos(i).intTarjeta1).byFigura <> GvecTarjetas(vecTercetos(i).intTarjeta2).byFigura _
                   And GvecTarjetas(vecTercetos(i).intTarjeta2).byFigura <> GvecTarjetas(vecTercetos(i).intTarjeta3).byFigura _
                   And GvecTarjetas(vecTercetos(i).intTarjeta1).byFigura <> GvecTarjetas(vecTercetos(i).intTarjeta3).byFigura Then
                    'Son las tres distintas
                    blnCanjeValido = True
                End If
            End If
            
            If blnCanjeValido Then
                'Calcula el valor
                If GvecTarjetas(vecTercetos(i).intTarjeta1).blCobrada Then
                    intValor = intValor + intValorCobrada
                Else
                    intValor = intValor + intValorNoCobrada
                End If
                If GvecTarjetas(vecTercetos(i).intTarjeta1).byFigura = enuFigurasTarjetas.figComodin Then
                    intValor = intValor + intValorComodin
                End If
                
                If GvecTarjetas(vecTercetos(i).intTarjeta2).blCobrada Then
                    intValor = intValor + intValorCobrada
                Else
                    intValor = intValor + intValorNoCobrada
                End If
                If GvecTarjetas(vecTercetos(i).intTarjeta2).byFigura = enuFigurasTarjetas.figComodin Then
                    intValor = intValor + intValorComodin
                End If
                
                If GvecTarjetas(vecTercetos(i).intTarjeta3).blCobrada Then
                    intValor = intValor + intValorCobrada
                Else
                    intValor = intValor + intValorNoCobrada
                End If
                If GvecTarjetas(vecTercetos(i).intTarjeta3).byFigura = enuFigurasTarjetas.figComodin Then
                    intValor = intValor + intValorComodin
                End If
                
                If intValor > intMejorValor Then
                    intMejorValor = intValor
                    intMejorIndice = i
                End If
            End If
            
        End If
        
    Next i
    
    'Si hubo algun terceto valido
    If intMejorValor > 0 Then
        'Canjea las tarjetas del mejor terceto
        IniciarTimer
        cCanjearTarjetas vecTercetos(intMejorIndice).intTarjeta1, vecTercetos(intMejorIndice).intTarjeta2, vecTercetos(intMejorIndice).intTarjeta3
    Else
        'Pasa a agregar tropas
        EfectuarAccion
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "vCanjearTarjetas", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Function vObtenerPaisObjetivoAtaque(ByRef byPaisOrigen As Byte, ByRef byPaisDestino As Byte, mtzE() As Single) As Boolean
    'Devuelve true si existe algun pais que supere el umbral para el ataque
    On Error GoTo ErrorHandle
    Dim sngMayor As Single
    Dim byOrigen As Byte
    Dim bydestino As Byte
    
    sngMayor = 0
    
    'Busca dentro de la matriz E la tupla con mayor valor
    For byOrigen = LBound(mtzE, 1) To UBound(mtzE, 1)
        For bydestino = LBound(mtzE, 2) To UBound(mtzE, 2)
            If mtzE(byOrigen, bydestino) > sngMayor Then
                sngMayor = mtzE(byOrigen, bydestino)
                byPaisOrigen = byOrigen
                byPaisDestino = bydestino
            End If
        Next
    Next
    
    'Si no se encontró ninguno mayor que el umbral
    If sngMayor <= GsngUmbralAtaque Then
        vObtenerPaisObjetivoAtaque = False
    Else
        vObtenerPaisObjetivoAtaque = True
    End If
    
    Exit Function
ErrorHandle:
    ReportErr "vObtenerPaisObjetivoAtaque", "mdlVP", Err.Description, Err.Number, Err.Source
End Function

Public Function vCalcularProbabilidadGanar(intFichasOrigen As Integer, intFichasDestino As Integer) As Double
    'Calcula la probabilidad de ganar de acuerdo a la cantidad de fichas
    'que tenga el pais origen y las que tenga el pais destino
    On Error GoTo ErrorHandle
    
    vCalcularProbabilidadGanar = ProbabilidadDeGanarGuerra(intFichasOrigen - 1, intFichasDestino)
    
    Exit Function
ErrorHandle:
    ReportErr "vCalcularProbabilidadGanar", "mdlVP", Err.Description, Err.Number, Err.Source
End Function

Public Sub vCalcularMatrizE(ByRef mtzE() As Single, vecMisPaises() As Byte)
    On Error GoTo ErrorHandle
    Dim byOrigen As Byte
    Dim bydestino As Byte
    Dim i As Integer
    Dim j As Integer
    Dim vecLimitrofes() As Byte
    Dim sngP As Single
    
    'Inicializa en 0
    For byOrigen = LBound(mtzE, 1) To UBound(mtzE, 1)
        For bydestino = LBound(mtzE, 2) To UBound(mtzE, 2)
            mtzE(byOrigen, bydestino) = 0
        Next
    Next
    
    'Por cada uno de mis paises
    For i = LBound(vecMisPaises) To UBound(vecMisPaises)
        'Calcula el valor de la esperanza matemática
        'por cada uno de sus paises limitrofes
        byOrigen = vecMisPaises(i)
'        CargarPaisesLimitrofes vecLimitrofes, byOrigen
        vPaisesLimitrofes byOrigen, vecLimitrofes, False, False
        For j = LBound(vecLimitrofes) To UBound(vecLimitrofes)
            bydestino = vecLimitrofes(j)
            'Si no tiene paises limitrofes o el pais de destino es mio
            If bydestino = 0 Or GvecPaises(bydestino).intColor = GintMiColor Then
                mtzE(byOrigen, bydestino) = 0
            Else
                'Calcula la probabilidad de ganar
                sngP = vCalcularProbabilidadGanar(GvecPaises(byOrigen).intCantidadFichas, GvecPaises(bydestino).intCantidadFichas)
                
                'Guarda la Esperanza Matematica
                mtzE(byOrigen, bydestino) = sngP * GvecPaises(bydestino).sngValor
            End If
        Next
    Next
    
    
    Exit Sub
ErrorHandle:
    ReportErr "vCalcularMatrizE", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub vPaisesLimitrofes(byPais As Byte, ByRef vecLimitrofes() As Byte, Optional blnSoloMios As Boolean = True, Optional blnIncluirPaisDesde As Boolean = True)
    'Devuelve un vector con los paises limitrofes del pais pasado por parametro, incluyéndolo o no de acuerdo a blnIncluirPaisDesde
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    ReDim vecLimitrofes(0)
    
    If blnIncluirPaisDesde Then
        vecLimitrofes(0) = byPais
        ReDim Preserve vecLimitrofes(1)
    End If
    
    'Recorre el vector de paises limítrofes en busca de byPais
    For i = LBound(GvecLimitrofes) To UBound(GvecLimitrofes)
        If GvecLimitrofes(i).byPaisDesde = byPais Then
            If (blnSoloMios And GvecPaises(GvecLimitrofes(i).byPaisHasta).intColor = GintMiColor) _
            Or Not blnSoloMios Then
                'Lo inserta en el vector
                vecLimitrofes(UBound(vecLimitrofes)) = GvecLimitrofes(i).byPaisHasta
                ReDim Preserve vecLimitrofes(UBound(vecLimitrofes) + 1)
            End If
        End If
    Next
    
    ReDim Preserve vecLimitrofes(UBound(vecLimitrofes) - 1)
    
    Exit Sub
ErrorHandle:
    ReportErr "vPaisesLimitrofes", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CargarLimitrofesServidor(vecLimitrofes() As String)
    'Carga los paises limitrofes de acuerdo a lo que indica el servidor
    'Se cargan siempre
    On Error GoTo ErrorHandle
    Dim byPais As Byte
    Dim byPaisDesde As Byte
    Dim byPaisHasta As Byte
    Dim i As Integer
    Dim vecPais As Variant
    
    ReDim GvecLimitrofes(0)
    
    For i = LBound(vecLimitrofes) To UBound(vecLimitrofes)
        vecPais = Split(vecLimitrofes(i), ",", 2)
        GvecLimitrofes(UBound(GvecLimitrofes)).byPaisDesde = vecPais(0)
        GvecLimitrofes(UBound(GvecLimitrofes)).byPaisHasta = vecPais(1)
        ReDim Preserve GvecLimitrofes(UBound(GvecLimitrofes) + 1)
    Next
    
    ReDim Preserve GvecLimitrofes(UBound(GvecLimitrofes) - 1)
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarLimitrofesServidor", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub ReportErr(ByVal strFuncion As String, ByVal strModulo As String, ByVal strDesc As String, _
                    ByVal intErr As Long, ByVal strSource As String, _
                    Optional styIcono As VbMsgBoxStyle = vbCritical)
    
    'Cambia el icono
    frmPrincipal.imgSysTray.Picture = frmPrincipal.imgLstSysTrayError.ListImages(GintJvColor).Picture
    SysTrayChangeIcon frmPrincipal.hwnd, frmPrincipal.imgSysTray
    
    'Reportes de Errores
    On Error GoTo ErrorHandle
    Dim strMsg As String
    If intErr <> 0 Then
        strMsg = "Rutina: " & strFuncion & Chr(10) & "Módulo: " & strModulo & vbCrLf _
            & "" _
            & "Descripción: " & strDesc & Space(10) & vbCrLf _
            & "Origen: " & strSource
        MsgBox strMsg, styIcono, "Error en sistema #" & intErr
    End If
    strMsg = "Number #" & intErr & " - Rutina: " & strFuncion & " - Módulo: " & strModulo & " - Descripción: " & strDesc & " - Source: " & strSource
    
    Err.Clear

    Exit Sub

ErrorHandle:
    MsgBox "Error en rutina ReportErr" & Chr(10) & "Descripción: " & Err.Description, vbCritical, "Error #" & Err.Number
    Close

End Sub

Public Sub cRecibirVersionServidor(intMajor As Integer, intMinor As Integer, intRevision As Integer)
    'Valida que la versión del servidor sea compatible con la versión del cliente
    On Error GoTo ErrorHandle
    
    cEnviarVersionCliente
        
    Exit Sub
ErrorHandle:
    ReportErr "cRecibirVersionServidor", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cEnviarVersionCliente()
    'Informa al servidor de la version y tipo del cliente
    On Error GoTo ErrorHandle
    
    'Tipo de Cliente:
    EnviarMensaje ArmarMensajeParam(msgVersionCliente, enuInteligenciaJugador.hrRobot, App.Major, App.Minor, App.Revision)
    
    Exit Sub
ErrorHandle:
    ReportErr "cEnviarVersionCliente", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cPartidaPausada()
    'Detiene el funcionamiento del Robot,
    'mientras la partida se encuentra Pausada
    On Error GoTo ErrorHandle
    
    'Guarda el estado Anterior (al cual luego deberá volver)
    CargarRegistroEstado estPartidaPausada, evePartidaContinuada, GEstadoCliente
    '###E
    ActualizarEstadoCliente evePartidaPausada
    
    'Desactiva el timer del Timeout
    frmPrincipal.tmr_TimeOut.Interval = 0
    
    Exit Sub
ErrorHandle:
    ReportErr "cPartidaPausada", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cPartidaContinuada()
    'Continua el funcionamiento del Robot,
    'luego de haberse continuado la partida
    On Error GoTo ErrorHandle
    
    '###E
    ActualizarEstadoCliente evePartidaContinuada
    
    EfectuarAccion
    
    Exit Sub
ErrorHandle:
    ReportErr "cPartidaContinuada", "mdlVP", Err.Description, Err.Number, Err.Source
End Sub

