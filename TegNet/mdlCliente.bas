Attribute VB_Name = "mdlCliente"
Option Explicit

'Constantes
'-----------------------------------------------
Public Const CintPausaMsAgregado = 300
Public Const CintPausaMsMovimiento = 300
Public Const CintTiempoGiroMs = 2400

'Tipos globales
'-----------------------------------------------
Public Type typTarjeta
    byPais As Byte
    byFigura As enuFigurasTarjetas
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
    byTipoJugador As Byte    '1=Real, 2=Virtual
    strDirIP As String
    strVersion As String
End Type
    
'Tipo de Opcion
Public Type typOpcion
    Id As Integer
    Valor As String
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

'Paneles de la barra de estado
Public Enum enuPaneles
    panIcoIntercambio = 1
    panIcoColor = 2
'    panConexion = 3 (disponible)
    panTipoRonda = 4
    panEstado = 5
    panADM = 6
    panIcoTimer = 7
    panTimer = 8
End Enum

'Botones de la barra de herramientas
Public Enum enuToolBar
    tbConexion = 1
    tbResincronizacion = 2
    tbOpciones = 3
    tbGuardar = 5
    tbPausa = 6
    tbMision = 8
    tbVerTarjetas = 9
    tbAgregar = 11
    tbAtacar = 12
    tbMover = 13
    tbTomarTarjeta = 14
    tbFinTurno = 15
End Enum

'Controles
Public Enum enuControles
    coConexion
    coResincronizar
    coOpciones
    coAgregar
    coAtacar
    coMover
    coTomarTarjeta
    coFinTurno
    coMapa
    coCanjearTarjeta
    coCobrarTarjeta
    coVerMision
    coVerTarjetas
    coGuardar
    coCambiarAdm
    coAsignarJV
    coBajarServidor
    coEnviarChat
    coPausar
End Enum

'Formas de Bajar el Servidor
Public Enum enuServidorCerrado
    secEventual
    secVoluntaria
    secSalida
End Enum

'Matriz de estados
Public GMatrizEstados(0 To 12, 0 To 14) As enuEstadoCli

'Matriz de controles
Public GMatrizControles(0 To 12, 0 To 18) As Boolean

Public GEstadoCliente As enuEstadoCli
Public GvecEstadoCliente(12) As String

'Variables globales

Public GblnMapaHabilitado As Boolean

Public GintCantJugadores As Integer

Public GsoyAdministrador As Boolean
'Color Jugador Local
Public GintMiColor As Integer
'Nombre Jugador Local
Public GstrMiNombre As String

'Vector que contiene las opciones
Public GvecOpciones() As typOpcion
Public GvecOpcionesDefault() As typOpcion

'Este vector contiene información de todos los jugadores
'Su indice representa el color asignado
Public GvecJugadores(6) As typJugador
Public GvecColores(6) As Long
Public GvecColoresInv(6) As Long 'Guarda los colores inversos a los colores
Public GvecTarjetas(5) As typTarjeta

'Vector que contiene los nombres de las partidas guardadas
Public GvecNombresPartidas() As String
 
Public GbyCantidadPaises As Byte
Public GbyPaisActual As Byte
Public GbyPaisAnterior As Byte

'Color que está jugando
Public GintColorActual As Integer

'Indica a todas las pantallas que se cierren sin preguntar
Public GblnSeCierra As Boolean

'Indica el tipo de ronda
Public GintTipoRonda As enuTipoRonda

'Flag que indica si los jugadores se están reconectando
Public GintTipoJugadoresConectados As enuTipoJugadoresConectados

'Mision
Public GstrMision As String

'Paises seleccionados
Public GbyPaisSeleccionadoOrigen As Byte
Public GbyPaisSeleccionadoDestino As Byte

'Cantidad de conquistas del turno
Public GintCantConquistas As Integer

'Valor inicial del timer del turno
Public GsngInicioTimerTurno As Single
Public GsngInicioPausaTimerTurno As Single
Public GintTimerTurno As Integer

'Flag usado para el icono del systray
Public intFlagSysTray As Integer

'Guarda el puerto y el servidor de la partida
Public GintPuerto As Integer
Public GstrServidor As String
Public GstrIpServidor As String

'Versión del servidor
Public GlngVersionServidor As Long

'Indica si el Servidor fue bajado voluntariamente
Public GintServidorCerrado As enuServidorCerrado

Public Sub Main()
    On Error GoTo ErrorHandle
    
    Dim intBaseIdioma As Integer
    
    'Obtiene de la registry el idioma
    intBaseIdioma = CargarSeteo("BaseIdioma", -1)
    
    'Si no está configurado el idioma...
    If intBaseIdioma = -1 Then
        'Por defecto arranca en español
        GintBaseIdioma = CintBaseSpanish
    Else
        GintBaseIdioma = intBaseIdioma
    End If
    
    'Splash
    frmInicio.Show
    frmInicio.Refresh
    AlwaysOnTop frmInicio, True
    
    'Si no está configurado el idioma...
    If intBaseIdioma = -1 Then
        'Muestra el formulario de seleccion de idioma
        frmIdioma.IdiomaSeteado = False
        MostrarFormulario frmIdioma, vbModal
    End If
    
    MostrarFormulario mdifrmPrincipal, vbModeless

    Exit Sub
ErrorHandle:
    ReportErr "Main", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub Informar(strTitulo As String, strMensaje As String, intGravedad As VbMsgBoxStyle, strProcedimiento As String)
    'Informa por pantalla un mensaje del servidor al cliente
    On Error GoTo ErrorHandle
    
    MostrarFormulario frmMensaje, vbModal
    frmMensaje.Caption = strTitulo
    frmMensaje.lblMensaje.Caption = strMensaje

    Exit Sub
ErrorHandle:
    ReportErr "Informar", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Function CantidadTropas(intColor As Integer) As Integer
    'Calcula la cantidad de tropas de un país dado
    On Error GoTo ErrorHandle
    
    Dim tmpCantTropas As Integer
    Dim byPais As Byte
    
    tmpCantTropas = 0
    
    For byPais = 1 To frmMapa.objPais.Count - 1
        If frmMapa.objPais(byPais).Color = intColor Then
            tmpCantTropas = tmpCantTropas + frmMapa.objPais(byPais).CantTropas
        End If
    Next byPais
    
    CantidadTropas = tmpCantTropas
    
    Exit Function
ErrorHandle:
    ReportErr "CantidadTropas", "mdlCliente", Err.Description, Err.Number, Err.Source
End Function

Public Function CantidadPaises(intColor As Integer) As Integer
    'Calcula la cantidad de paises de un color dado
    On Error GoTo ErrorHandle
    
    Dim tmpCantPaises As Integer
    Dim byPais As Byte
    
    tmpCantPaises = 0
    
    For byPais = 1 To frmMapa.objPais.Count - 1
        If frmMapa.objPais(byPais).Color = intColor Then
            tmpCantPaises = tmpCantPaises + 1
        End If
    Next byPais
    
    CantidadPaises = tmpCantPaises
    
    Exit Function
ErrorHandle:
    ReportErr "CantidadPaises", "mdlCliente", Err.Description, Err.Number, Err.Source
End Function

Public Function MayorACero(Valor As Variant) As Variant
    'Devuelve el valor cuando es mayor que cero y devuelve 0 si es menor a cero
    On Error GoTo ErrorHandle
    
    If Valor < 0 Then MayorACero = 0 Else MayorACero = Valor

    Exit Function
ErrorHandle:
    ReportErr "MayorACero", "mdlCliente", Err.Description, Err.Number, Err.Source
End Function

Public Sub SetearAdm(blnValor As Boolean)
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim blnSoloRobots As Boolean
    
    'Habilita/Deshabilita las opciones de Administración
    mdifrmPrincipal.MnuAdministracion.Visible = blnValor
    mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbGuardar - 1).Visible = blnValor
    mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbGuardar).Visible = blnValor
    mdifrmPrincipal.mnuPartidaGuardar.Visible = blnValor
    EvaluarHabilitacionPausa
    
    'Habilita/Deshabilita las opciones de la partida
    For i = 0 To frmOpciones.fraSolapas.Count - 1
        frmOpciones.fraSolapas(i).Enabled = blnValor
    Next i
    frmOpciones.cmdAceptar.Visible = blnValor
    frmOpciones.cmdCancelar.Caption = IIf(blnValor, ObtenerTextoRecurso(CintOpcionesCancelar), ObtenerTextoRecurso(CintOpcionesAceptar))
    frmOpciones.cmdGuardarDefault.Enabled = blnValor
    frmOpciones.lblRecuperarDefecto.Enabled = blnValor
    frmOpciones.cmdRecuperarDefault.Enabled = blnValor
    frmOpciones.lblGuardarDefecto.Enabled = blnValor
    
    'Informa en la statusbar si es o no Administrador
    mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panADM).Text = IIf(blnValor = True, ObtenerTextoRecurso(CintGralTextoAdministrador), "")
    
    GsoyAdministrador = blnValor
    
    Exit Sub
ErrorHandle:
    ReportErr "SetearAdm", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub InicializarJugadores()
    'Esta rutina inicializa los valores del vector que contiene
    'la información de los distintos jugadores
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    For i = LBound(GvecJugadores) To UBound(GvecJugadores)
        GvecJugadores(i).intCanje = 0
        GvecJugadores(i).intCantidadTarjetas = 0
        GvecJugadores(i).intEstado = conNoJuega
        GvecJugadores(i).intOrdenRonda = 0
        GvecJugadores(i).intTropasDisponibles = 0
        GvecJugadores(i).strNombre = ""
        GvecJugadores(i).byTipoJugador = 0
        GvecJugadores(i).strDirIP = ""
        GvecJugadores(i).strVersion = ""
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "InicializarJugadores", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub VisibilidadMapa(blnVisible As Boolean)
    'Muestra u oculta los paises del mapa
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    For i = 1 To frmMapa.objPais.Count - 1
        frmMapa.objPais(i).MostrarFicha = blnVisible
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "VisibilidadMapa", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CentrarFormulario(Formulario As Form)
    On Error GoTo ErrorHandle
    
    Formulario.Left = mdifrmPrincipal.ScaleWidth / 2 - Formulario.ScaleWidth / 2
    Formulario.Top = mdifrmPrincipal.ScaleHeight / 2 - Formulario.ScaleHeight / 2
    
    Exit Sub
ErrorHandle:
    ReportErr "CentrarFormulario", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub Mensaje(strMensaje As String, strTitulo As String, Optional intEstilo As VbMsgBoxStyle)
    On Error GoTo ErrorHandle
    
    frmMensaje.Caption = strTitulo
    frmMensaje.lblMensaje.Caption = strMensaje
    CentrarFormulario frmMensaje
    MostrarFormulario frmMensaje
    frmMapa.Enabled = False
    
    Exit Sub
ErrorHandle:
    ReportErr "Mensaje", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub MostrarAtaque(intDadoDesde1 As Integer, intDadoDesde2 As Integer, intDadoDesde3 As Integer, _
                        intDadoHasta1 As Integer, intDadoHasta2 As Integer, intDadoHasta3 As Integer, _
                        byPaisDesde As Byte, byPaisHasta As Byte)
    On Error GoTo ErrorHandle

    Dim intTimer As Integer
    Dim sngPausa As Single
    Dim intNroVuelta As Integer
    Dim blnFlagEfectoPais As Boolean

    'Limpia todo
    Set frmDados.imgJugador1(1) = Nothing
    Set frmDados.imgJugador1(2) = Nothing
    Set frmDados.imgJugador1(0) = Nothing
    Set frmDados.imgJugador2(1) = Nothing
    Set frmDados.imgJugador2(2) = Nothing
    Set frmDados.imgJugador2(0) = Nothing
    frmMapa.LimpiarMapa

    'Efecto especial (hace girar los dados)
    intNroVuelta = 1
    intTimer = CintTiempoGiroMs
    While intTimer > 0

        IniciarPausa 200, sngPausa

        'Alterna el valor
        blnFlagEfectoPais = Not blnFlagEfectoPais
        If blnFlagEfectoPais = True Then
            frmMapa.objPais(byPaisDesde).IniciarDestello
            frmMapa.objPais(byPaisDesde).ZOrder 0
            frmMapa.objPais(byPaisHasta).FinalizarDestello
        Else
            frmMapa.objPais(byPaisHasta).IniciarDestello
            frmMapa.objPais(byPaisHasta).ZOrder 0
            frmMapa.objPais(byPaisDesde).FinalizarDestello
        End If
        frmMapa.Refresh

        'Muestra dados random (de a uno)
        'son 12 vueltas, asi que cada 2 vueltas muestra un dado.
        Select Case intNroVuelta
            Case 2
                frmDados.imgJugador1(0).Picture = frmDados.iLstDados.ListImages(intDadoDesde1).Picture
                frmDados.imgJugador1(0).Refresh
            Case 4
                If intDadoDesde2 > 0 Then frmDados.imgJugador1(1).Picture = frmDados.iLstDados.ListImages(intDadoDesde2).Picture
                frmDados.imgJugador1(1).Refresh
            Case 6
                If intDadoDesde3 > 0 Then frmDados.imgJugador1(2).Picture = frmDados.iLstDados.ListImages(intDadoDesde3).Picture
                frmDados.imgJugador1(2).Refresh
            Case 8
                frmDados.imgJugador2(0).Picture = frmDados.iLstDados.ListImages(intDadoHasta1).Picture
                frmDados.imgJugador2(0).Refresh
            Case 10
                If intDadoHasta2 > 0 Then frmDados.imgJugador2(1).Picture = frmDados.iLstDados.ListImages(intDadoHasta2).Picture
                frmDados.imgJugador2(1).Refresh
            Case 12
                If intDadoHasta3 > 0 Then frmDados.imgJugador2(2).Picture = frmDados.iLstDados.ListImages(intDadoHasta3).Picture
                frmDados.imgJugador2(2).Refresh
        End Select

        FinalizarPausa 200, sngPausa

        'Decrementa la variable con el tiempo de vida del timer
        intTimer = intTimer - 200
        intNroVuelta = intNroVuelta + 1
    Wend

    'Deja iluminados los paises involucrados
    frmMapa.objPais(byPaisDesde).Restaurar
    frmMapa.objPais(byPaisHasta).Restaurar
    frmMapa.objPais(byPaisHasta).ZOrder 0
    frmMapa.Refresh

    Exit Sub
ErrorHandle:
    ReportErr "MostrarAtaque", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub EfectoMover(byPaisDesde As Byte, byPaisHasta As Byte)
    On Error GoTo ErrorHandle
    
    'Deselecciona todo
    frmMapa.LimpiarMapa
    
    'Efecto especial
    frmMapa.objPais(byPaisDesde).IniciarDestello
    frmMapa.objPais(byPaisDesde).ZOrder 0
    frmMapa.Refresh
    Pausa CintPausaMsMovimiento, False
    frmMapa.objPais(byPaisDesde).FinalizarDestello
    
    frmMapa.objPais(byPaisHasta).IniciarDestello
    frmMapa.objPais(byPaisHasta).ZOrder 0
    frmMapa.Refresh
    Pausa CintPausaMsMovimiento, False
    frmMapa.objPais(byPaisHasta).FinalizarDestello
    
    frmMapa.objPais(byPaisDesde).Restaurar
    frmMapa.objPais(byPaisHasta).Restaurar

    Exit Sub
ErrorHandle:
    ReportErr "EfectoMover", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub


'Rutinas de manejo de estado del cliente

Public Sub CargarRegistroEstado(estadoOrigen As enuEstadoCli, eventoOrigen As enuEventosCli, estadoDestino As enuEstadoCli)
    'Subrutina utilizada para facilitar la carga de la matriz de estados
    On Error GoTo ErrorHandle
    
    GMatrizEstados(estadoOrigen, eventoOrigen) = estadoDestino
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarRegistroEstado", "mdlCliente", Err.Description, Err.Number, Err.Source
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
    ReportErr "CargarMatrizEstados", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub DescripcionEstadosCliente()
    'Carga un vector con las descripciones de los estados del cliente
    On Error GoTo ErrorHandle
    
    'Regional - Carga de estados
    GvecEstadoCliente(enuEstadoCli.estDesconectado) = ObtenerTextoRecurso(CintPrincipalEstadoDesconectado)
    GvecEstadoCliente(enuEstadoCli.estConectado) = ObtenerTextoRecurso(CintPrincipalEstadoConectado)
    GvecEstadoCliente(enuEstadoCli.estValidado) = ObtenerTextoRecurso(CintPrincipalEstadoValidado)
    GvecEstadoCliente(enuEstadoCli.estEsperandoTurno) = ObtenerTextoRecurso(CintPrincipalEstadoEsperandoTurno)
    GvecEstadoCliente(enuEstadoCli.estAgregando) = ObtenerTextoRecurso(CintPrincipalEstadoAgregando)
    GvecEstadoCliente(enuEstadoCli.estAtacando) = ObtenerTextoRecurso(CintPrincipalEstadoAtacando)
    GvecEstadoCliente(enuEstadoCli.estMoviendo) = ObtenerTextoRecurso(CintPrincipalEstadoMoviendo)
    GvecEstadoCliente(enuEstadoCli.estTarjetaTomada) = ObtenerTextoRecurso(CintPrincipalEstadoTarTomada)
    GvecEstadoCliente(enuEstadoCli.estTarjetaCobrada) = ObtenerTextoRecurso(CintPrincipalEstadoTarCobrada)
    GvecEstadoCliente(enuEstadoCli.estTarjetaCobradaTomada) = ObtenerTextoRecurso(CintPrincipalEstadoTarTomadaCobrada)
    GvecEstadoCliente(enuEstadoCli.estPartidaPausada) = ObtenerTextoRecurso(CintPrincipalEstadoPausada)
    GvecEstadoCliente(enuEstadoCli.estPartidaFinalizada) = ObtenerTextoRecurso(CintPrincipalEstadoFinalizada)
    GvecEstadoCliente(enuEstadoCli.estInconsistente) = ObtenerTextoRecurso(CintPrincipalEstadoInconsistente)
    
    Exit Sub
ErrorHandle:
    ReportErr "DescripcionEstadosCliente", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub ActualizarEstadoCliente(NuevoEvento As enuEventosCli)
    'Cambia el estado del cliente de acuerdo a un evento pasado por parametro
    On Error GoTo ErrorHandle
    
    GEstadoCliente = GMatrizEstados(GEstadoCliente, NuevoEvento)
    mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panEstado).Text = GvecEstadoCliente(GEstadoCliente)
    
    ActualizarControles
    '###E Agregar a las desactivaciones de opciones de menu
'''    If GEstadoCliente > estEsperandoTurno And GEstadoCliente <> estInconsistente Then
'''        GblnMapaHabilitado = True
'''    Else
'''        GblnMapaHabilitado = False
'''    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "ActualizarEstadoCliente", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub



'Rutinas de manejo de los controles del cliente
Public Sub CargarRegistroControl(intEstado As enuEstadoCli, intControl As enuControles, blnValor As Boolean)
    'Subrutina utilizada para facilitar la carga de la matriz de controles
    On Error GoTo ErrorHandle
    
    GMatrizControles(intEstado, intControl) = blnValor
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarRegistroControl", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CargarMatrizControles()
    On Error GoTo ErrorHandle
    
    CargarRegistroControl estDesconectado, coConexion, True
    CargarRegistroControl estDesconectado, coResincronizar, False
    CargarRegistroControl estDesconectado, coOpciones, False
    CargarRegistroControl estDesconectado, coAgregar, False
    CargarRegistroControl estDesconectado, coAtacar, False
    CargarRegistroControl estDesconectado, coMover, False
    CargarRegistroControl estDesconectado, coTomarTarjeta, False
    CargarRegistroControl estDesconectado, coFinTurno, False
    CargarRegistroControl estDesconectado, coMapa, False
    CargarRegistroControl estDesconectado, coCobrarTarjeta, False
    CargarRegistroControl estDesconectado, coCanjearTarjeta, False
    CargarRegistroControl estDesconectado, coVerMision, False
    CargarRegistroControl estDesconectado, coVerTarjetas, False
    CargarRegistroControl estDesconectado, coGuardar, False
    CargarRegistroControl estDesconectado, coCambiarAdm, False
    CargarRegistroControl estDesconectado, coAsignarJV, False
    CargarRegistroControl estDesconectado, coBajarServidor, False
    CargarRegistroControl estDesconectado, coEnviarChat, False
    CargarRegistroControl estDesconectado, coPausar, False
    
    CargarRegistroControl estConectado, coConexion, True
    CargarRegistroControl estConectado, coResincronizar, False
    CargarRegistroControl estConectado, coOpciones, False
    CargarRegistroControl estConectado, coAgregar, False
    CargarRegistroControl estConectado, coAtacar, False
    CargarRegistroControl estConectado, coMover, False
    CargarRegistroControl estConectado, coTomarTarjeta, False
    CargarRegistroControl estConectado, coFinTurno, False
    CargarRegistroControl estConectado, coMapa, False
    CargarRegistroControl estConectado, coCobrarTarjeta, False
    CargarRegistroControl estConectado, coCanjearTarjeta, False
    CargarRegistroControl estConectado, coVerMision, False
    CargarRegistroControl estConectado, coVerTarjetas, False
    CargarRegistroControl estConectado, coGuardar, False
    CargarRegistroControl estConectado, coCambiarAdm, False
    CargarRegistroControl estConectado, coAsignarJV, False
    CargarRegistroControl estConectado, coBajarServidor, True
    CargarRegistroControl estConectado, coEnviarChat, False
    CargarRegistroControl estConectado, coPausar, False
    
    CargarRegistroControl estValidado, coConexion, True
    CargarRegistroControl estValidado, coResincronizar, False
    CargarRegistroControl estValidado, coOpciones, False
    CargarRegistroControl estValidado, coAgregar, False
    CargarRegistroControl estValidado, coAtacar, False
    CargarRegistroControl estValidado, coMover, False
    CargarRegistroControl estValidado, coTomarTarjeta, False
    CargarRegistroControl estValidado, coFinTurno, False
    CargarRegistroControl estValidado, coMapa, False
    CargarRegistroControl estValidado, coCobrarTarjeta, False
    CargarRegistroControl estValidado, coCanjearTarjeta, False
    CargarRegistroControl estValidado, coVerMision, False
    CargarRegistroControl estValidado, coVerTarjetas, False
    CargarRegistroControl estValidado, coGuardar, False
    CargarRegistroControl estValidado, coCambiarAdm, False
    CargarRegistroControl estValidado, coAsignarJV, True
    CargarRegistroControl estValidado, coBajarServidor, True
    CargarRegistroControl estValidado, coEnviarChat, True
    CargarRegistroControl estValidado, coPausar, False
    
    CargarRegistroControl estEsperandoTurno, coConexion, True
    CargarRegistroControl estEsperandoTurno, coResincronizar, True
    CargarRegistroControl estEsperandoTurno, coOpciones, True
    CargarRegistroControl estEsperandoTurno, coAgregar, False
    CargarRegistroControl estEsperandoTurno, coAtacar, False
    CargarRegistroControl estEsperandoTurno, coMover, False
    CargarRegistroControl estEsperandoTurno, coTomarTarjeta, False
    CargarRegistroControl estEsperandoTurno, coFinTurno, False
    CargarRegistroControl estEsperandoTurno, coMapa, False
    CargarRegistroControl estEsperandoTurno, coCobrarTarjeta, False
    CargarRegistroControl estEsperandoTurno, coCanjearTarjeta, False
    CargarRegistroControl estEsperandoTurno, coVerMision, True
    CargarRegistroControl estEsperandoTurno, coVerTarjetas, True
    CargarRegistroControl estEsperandoTurno, coGuardar, True
    CargarRegistroControl estEsperandoTurno, coCambiarAdm, True
    CargarRegistroControl estEsperandoTurno, coAsignarJV, True
    CargarRegistroControl estEsperandoTurno, coBajarServidor, True
    CargarRegistroControl estEsperandoTurno, coEnviarChat, True
    CargarRegistroControl estEsperandoTurno, coPausar, True
    
    CargarRegistroControl estAgregando, coConexion, True
    CargarRegistroControl estAgregando, coResincronizar, True
    CargarRegistroControl estAgregando, coOpciones, True
    CargarRegistroControl estAgregando, coAgregar, True
    CargarRegistroControl estAgregando, coAtacar, False
    CargarRegistroControl estAgregando, coMover, False
    CargarRegistroControl estAgregando, coTomarTarjeta, False
    CargarRegistroControl estAgregando, coFinTurno, True
    CargarRegistroControl estAgregando, coMapa, True
    CargarRegistroControl estAgregando, coCobrarTarjeta, False
    CargarRegistroControl estAgregando, coCanjearTarjeta, True
    CargarRegistroControl estAgregando, coVerMision, True
    CargarRegistroControl estAgregando, coVerTarjetas, True
    CargarRegistroControl estAgregando, coGuardar, True
    CargarRegistroControl estAgregando, coCambiarAdm, True
    CargarRegistroControl estAgregando, coAsignarJV, True
    CargarRegistroControl estAgregando, coBajarServidor, True
    CargarRegistroControl estAgregando, coEnviarChat, True
    CargarRegistroControl estAgregando, coPausar, True
    
    CargarRegistroControl estAtacando, coConexion, True
    CargarRegistroControl estAtacando, coResincronizar, True
    CargarRegistroControl estAtacando, coOpciones, True
    CargarRegistroControl estAtacando, coAgregar, False
    CargarRegistroControl estAtacando, coAtacar, True
    CargarRegistroControl estAtacando, coMover, True
    CargarRegistroControl estAtacando, coTomarTarjeta, True
    CargarRegistroControl estAtacando, coFinTurno, True
    CargarRegistroControl estAtacando, coMapa, True
    CargarRegistroControl estAtacando, coCobrarTarjeta, True
    CargarRegistroControl estAtacando, coCanjearTarjeta, False
    CargarRegistroControl estAtacando, coVerMision, True
    CargarRegistroControl estAtacando, coVerTarjetas, True
    CargarRegistroControl estAtacando, coGuardar, True
    CargarRegistroControl estAtacando, coCambiarAdm, True
    CargarRegistroControl estAtacando, coAsignarJV, True
    CargarRegistroControl estAtacando, coBajarServidor, True
    CargarRegistroControl estAtacando, coEnviarChat, True
    CargarRegistroControl estAtacando, coPausar, True
    
    CargarRegistroControl estMoviendo, coConexion, True
    CargarRegistroControl estMoviendo, coResincronizar, True
    CargarRegistroControl estMoviendo, coOpciones, True
    CargarRegistroControl estMoviendo, coAgregar, False
    CargarRegistroControl estMoviendo, coAtacar, False
    CargarRegistroControl estMoviendo, coMover, True
    CargarRegistroControl estMoviendo, coTomarTarjeta, True
    CargarRegistroControl estMoviendo, coFinTurno, True
    CargarRegistroControl estMoviendo, coMapa, True
    CargarRegistroControl estMoviendo, coCobrarTarjeta, True
    CargarRegistroControl estMoviendo, coCanjearTarjeta, False
    CargarRegistroControl estMoviendo, coVerMision, True
    CargarRegistroControl estMoviendo, coVerTarjetas, True
    CargarRegistroControl estMoviendo, coGuardar, True
    CargarRegistroControl estMoviendo, coCambiarAdm, True
    CargarRegistroControl estMoviendo, coAsignarJV, True
    CargarRegistroControl estMoviendo, coBajarServidor, True
    CargarRegistroControl estMoviendo, coEnviarChat, True
    CargarRegistroControl estMoviendo, coPausar, True
    
    CargarRegistroControl estTarjetaTomada, coConexion, True
    CargarRegistroControl estTarjetaTomada, coResincronizar, True
    CargarRegistroControl estTarjetaTomada, coOpciones, True
    CargarRegistroControl estTarjetaTomada, coAgregar, False
    CargarRegistroControl estTarjetaTomada, coAtacar, False
    CargarRegistroControl estTarjetaTomada, coMover, False
    CargarRegistroControl estTarjetaTomada, coTomarTarjeta, False
    CargarRegistroControl estTarjetaTomada, coFinTurno, True
    CargarRegistroControl estTarjetaTomada, coMapa, True
    CargarRegistroControl estTarjetaTomada, coCobrarTarjeta, True
    CargarRegistroControl estTarjetaTomada, coCanjearTarjeta, False
    CargarRegistroControl estTarjetaTomada, coVerMision, True
    CargarRegistroControl estTarjetaTomada, coVerTarjetas, True
    CargarRegistroControl estTarjetaTomada, coGuardar, True
    CargarRegistroControl estTarjetaTomada, coCambiarAdm, True
    CargarRegistroControl estTarjetaTomada, coAsignarJV, True
    CargarRegistroControl estTarjetaTomada, coBajarServidor, True
    CargarRegistroControl estTarjetaTomada, coEnviarChat, True
    CargarRegistroControl estTarjetaTomada, coPausar, True
    
    CargarRegistroControl estTarjetaCobrada, coConexion, True
    CargarRegistroControl estTarjetaCobrada, coResincronizar, True
    CargarRegistroControl estTarjetaCobrada, coOpciones, True
    CargarRegistroControl estTarjetaCobrada, coAgregar, False
    CargarRegistroControl estTarjetaCobrada, coAtacar, False
    CargarRegistroControl estTarjetaCobrada, coMover, False
    CargarRegistroControl estTarjetaCobrada, coTomarTarjeta, True
    CargarRegistroControl estTarjetaCobrada, coFinTurno, True
    CargarRegistroControl estTarjetaCobrada, coMapa, True
    CargarRegistroControl estTarjetaCobrada, coCobrarTarjeta, True
    CargarRegistroControl estTarjetaCobrada, coCanjearTarjeta, False
    CargarRegistroControl estTarjetaCobrada, coVerMision, True
    CargarRegistroControl estTarjetaCobrada, coVerTarjetas, True
    CargarRegistroControl estTarjetaCobrada, coGuardar, True
    CargarRegistroControl estTarjetaCobrada, coCambiarAdm, True
    CargarRegistroControl estTarjetaCobrada, coAsignarJV, True
    CargarRegistroControl estTarjetaCobrada, coBajarServidor, True
    CargarRegistroControl estTarjetaCobrada, coEnviarChat, True
    CargarRegistroControl estTarjetaCobrada, coPausar, True
    
    CargarRegistroControl estTarjetaCobradaTomada, coConexion, True
    CargarRegistroControl estTarjetaCobradaTomada, coResincronizar, True
    CargarRegistroControl estTarjetaCobradaTomada, coOpciones, True
    CargarRegistroControl estTarjetaCobradaTomada, coAgregar, False
    CargarRegistroControl estTarjetaCobradaTomada, coAtacar, False
    CargarRegistroControl estTarjetaCobradaTomada, coMover, False
    CargarRegistroControl estTarjetaCobradaTomada, coTomarTarjeta, False
    CargarRegistroControl estTarjetaCobradaTomada, coFinTurno, True
    CargarRegistroControl estTarjetaCobradaTomada, coMapa, True
    CargarRegistroControl estTarjetaCobradaTomada, coCobrarTarjeta, True
    CargarRegistroControl estTarjetaCobradaTomada, coCanjearTarjeta, False
    CargarRegistroControl estTarjetaCobradaTomada, coVerMision, True
    CargarRegistroControl estTarjetaCobradaTomada, coVerTarjetas, True
    CargarRegistroControl estTarjetaCobradaTomada, coGuardar, True
    CargarRegistroControl estTarjetaCobradaTomada, coCambiarAdm, True
    CargarRegistroControl estTarjetaCobradaTomada, coAsignarJV, True
    CargarRegistroControl estTarjetaCobradaTomada, coBajarServidor, True
    CargarRegistroControl estTarjetaCobradaTomada, coEnviarChat, True
    CargarRegistroControl estTarjetaCobradaTomada, coPausar, True
    
    CargarRegistroControl estPartidaPausada, coConexion, True
    CargarRegistroControl estPartidaPausada, coResincronizar, False
    CargarRegistroControl estPartidaPausada, coOpciones, False
    CargarRegistroControl estPartidaPausada, coAgregar, False
    CargarRegistroControl estPartidaPausada, coAtacar, False
    CargarRegistroControl estPartidaPausada, coMover, False
    CargarRegistroControl estPartidaPausada, coTomarTarjeta, False
    CargarRegistroControl estPartidaPausada, coFinTurno, False
    CargarRegistroControl estPartidaPausada, coMapa, False
    CargarRegistroControl estPartidaPausada, coCobrarTarjeta, False
    CargarRegistroControl estPartidaPausada, coCanjearTarjeta, False
    CargarRegistroControl estPartidaPausada, coVerMision, False
    CargarRegistroControl estPartidaPausada, coVerTarjetas, False
    CargarRegistroControl estPartidaPausada, coGuardar, False
    CargarRegistroControl estPartidaPausada, coCambiarAdm, False
    CargarRegistroControl estPartidaPausada, coAsignarJV, False
    CargarRegistroControl estPartidaPausada, coBajarServidor, True
    CargarRegistroControl estPartidaPausada, coEnviarChat, False
    CargarRegistroControl estPartidaPausada, coPausar, True
    
    CargarRegistroControl estPartidaFinalizada, coConexion, True
    CargarRegistroControl estPartidaFinalizada, coResincronizar, True
    CargarRegistroControl estPartidaFinalizada, coOpciones, True
    CargarRegistroControl estPartidaFinalizada, coAgregar, False
    CargarRegistroControl estPartidaFinalizada, coAtacar, False
    CargarRegistroControl estPartidaFinalizada, coMover, False
    CargarRegistroControl estPartidaFinalizada, coTomarTarjeta, False
    CargarRegistroControl estPartidaFinalizada, coFinTurno, False
    CargarRegistroControl estPartidaFinalizada, coMapa, False
    CargarRegistroControl estPartidaFinalizada, coCobrarTarjeta, False
    CargarRegistroControl estPartidaFinalizada, coCanjearTarjeta, False
    CargarRegistroControl estPartidaFinalizada, coVerMision, True
    CargarRegistroControl estPartidaFinalizada, coVerTarjetas, True
    CargarRegistroControl estPartidaFinalizada, coGuardar, True
    CargarRegistroControl estPartidaFinalizada, coCambiarAdm, True
    CargarRegistroControl estPartidaFinalizada, coAsignarJV, False
    CargarRegistroControl estPartidaFinalizada, coBajarServidor, True
    CargarRegistroControl estPartidaFinalizada, coEnviarChat, True
    CargarRegistroControl estPartidaFinalizada, coPausar, False
    
    CargarRegistroControl estInconsistente, coConexion, True
    CargarRegistroControl estInconsistente, coResincronizar, True
    CargarRegistroControl estInconsistente, coOpciones, False
    CargarRegistroControl estInconsistente, coAgregar, False
    CargarRegistroControl estInconsistente, coAtacar, False
    CargarRegistroControl estInconsistente, coMover, False
    CargarRegistroControl estInconsistente, coTomarTarjeta, False
    CargarRegistroControl estInconsistente, coFinTurno, False
    CargarRegistroControl estInconsistente, coMapa, False
    CargarRegistroControl estInconsistente, coCobrarTarjeta, False
    CargarRegistroControl estInconsistente, coCanjearTarjeta, False
    CargarRegistroControl estInconsistente, coVerMision, False
    CargarRegistroControl estInconsistente, coVerTarjetas, False
    CargarRegistroControl estInconsistente, coGuardar, False
    CargarRegistroControl estInconsistente, coCambiarAdm, False
    CargarRegistroControl estInconsistente, coAsignarJV, False
    CargarRegistroControl estInconsistente, coBajarServidor, True
    CargarRegistroControl estInconsistente, coEnviarChat, True
    CargarRegistroControl estInconsistente, coPausar, False
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarMatrizControles", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub ActualizarControles()
    'De acuerdo al estado habilita/deshabilita controles
    On Error GoTo ErrorHandle
    Dim intControl As enuControles
    Dim blnValor As Boolean
    
    For intControl = coConexion To enuControles.coPausar
        blnValor = GMatrizControles(GEstadoCliente, intControl)
        Select Case intControl
            Case enuControles.coConexion
                mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbConexion).Enabled = blnValor
'                mdifrmPrincipal.mnuPartidaConectar.Enabled = blnValor
'                mdifrmPrincipal.mnuPartidaDesconectar.Enabled = False
            Case enuControles.coResincronizar
                mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbResincronizacion).Enabled = blnValor
                mdifrmPrincipal.mnuPartidaResincronizar.Enabled = blnValor
            Case enuControles.coOpciones
                mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbOpciones).Enabled = blnValor
                mdifrmPrincipal.MnuOpciones.Enabled = blnValor
            Case enuControles.coAgregar
                mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbAgregar).Enabled = blnValor
                mdifrmPrincipal.mnuJuegoAgregar.Enabled = blnValor
                mdifrmPrincipal.mnuAgregar.Enabled = blnValor
                mdifrmPrincipal.mnuAgregar1.Enabled = blnValor
                mdifrmPrincipal.mnuAgregar10.Enabled = blnValor
                mdifrmPrincipal.mnuAgregar5.Enabled = blnValor
                mdifrmPrincipal.mnuAgregarTodas.Enabled = blnValor
            Case enuControles.coAtacar
                mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbAtacar).Enabled = blnValor
                mdifrmPrincipal.mnuJuegoAtacar.Enabled = blnValor
                mdifrmPrincipal.mnuAtacar.Enabled = blnValor
            Case enuControles.coMover
                mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbMover).Enabled = blnValor
                mdifrmPrincipal.mnuJuegoMover.Enabled = blnValor
                mdifrmPrincipal.mnuMover.Enabled = blnValor
            Case enuControles.coTomarTarjeta
                mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbTomarTarjeta).Enabled = blnValor
                mdifrmPrincipal.mnuJuegoTomar.Enabled = blnValor
            Case enuControles.coFinTurno
                mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbFinTurno).Enabled = blnValor
                mdifrmPrincipal.mnuJuegoFinalizar.Enabled = blnValor
            Case enuControles.coMapa
                'Variable global que indica si el mapa está o no 'habilitado'
                GblnMapaHabilitado = blnValor
            Case enuControles.coCobrarTarjeta
                frmTarjetas.cmdCobrar.Enabled = blnValor
            Case enuControles.coCanjearTarjeta
                frmTarjetas.cmdCanjear.Enabled = blnValor
            Case enuControles.coVerMision
                mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbMision).Enabled = blnValor
                mdifrmPrincipal.mnuVerMision.Enabled = blnValor
            Case enuControles.coVerTarjetas
                mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbVerTarjetas).Enabled = blnValor
                mdifrmPrincipal.mnuVerTarjetas.Enabled = blnValor
            Case enuControles.coGuardar
                mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbGuardar).Enabled = blnValor
                mdifrmPrincipal.mnuPartidaGuardar.Enabled = blnValor
            Case enuControles.coCambiarAdm
                mdifrmPrincipal.MnuCambiarAdministrador.Enabled = blnValor
            Case enuControles.coAsignarJV
                mdifrmPrincipal.mnuAsignarJV.Enabled = blnValor
            Case enuControles.coBajarServidor
                mdifrmPrincipal.MnuBajarServidor.Enabled = blnValor
            Case enuControles.coEnviarChat
                frmChat.cmdEnviar.Enabled = blnValor
            Case enuControles.coPausar
                mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbPausa).Enabled = blnValor
                mdifrmPrincipal.mnuPartidaPausar.Enabled = blnValor
        End Select
    Next intControl
    
    Exit Sub
ErrorHandle:
    ReportErr "ActualizarControles", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub


'Rutinas de comunicación

Public Sub EnviarMensaje(strMensaje As String)
    On Error GoTo ErrorHandle
    
    mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panIcoIntercambio).Picture = mdifrmPrincipal.imgLstFichas.ListImages(7).Picture
    mdifrmPrincipal.Winsock1.SendData strMensaje
    mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panIcoIntercambio).Picture = Nothing
    
    Exit Sub
ErrorHandle:
    ReportErr "EnviarMensaje", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub DistribuirMensaje(intTipoMensaje As enuTipoMensaje, vecParametros() As String, Optional intIndiceOrigenMensaje As Integer)
    On Error GoTo ErrorHandle
    
    Dim i As Integer

    'De acuerdo al tipo de mensaje
    Select Case intTipoMensaje
        Case msgConfirmarAdm
            '###E
            ActualizarEstadoCliente eveConfirmacionAdm
            '###Version
            '(solo para el caso que se cambie el ADM)
            If UBound(vecParametros) = -1 Then
                'Servidor 1.0.0
                cConfirmarAdm 2
            Else
                cConfirmarAdm CInt(vecParametros(0))
            End If
        Case msgJugadoresConectados
            '###E
            ActualizarEstadoCliente eveJugadoresConectados
            '###Version
            If UBound(vecParametros) = 12 Then
                'Servidor 1.0.0
                ReDim Preserve vecParametros(30)
                'Mueve el campo Tipo de Mensaje al Final
                vecParametros(UBound(vecParametros)) = vecParametros(12)
                'Pone en cero el tipo de Inteligencia del Jugador
                For i = 1 To UBound(GvecJugadores)
                    vecParametros(11 + i) = "0"
                Next i
            End If
            
            cRefrescarConexiones vecParametros
            'Habilita/Deshabilita la opcion Pausa, si son todos Robots o no.
            EvaluarHabilitacionPausa
        
        Case msgPais
            '###Version
            If UBound(vecParametros) = 3 Then
                'Servidor 1.0.0
                cActualizarPais CByte(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2)), CInt(vecParametros(3)), 0
            Else
                cActualizarPais CByte(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2)), CInt(vecParametros(3)), CInt(vecParametros(4))
            End If
            
        Case msgMisionAsignada
            '###Version
            If GlngVersionServidor < 10200 Then
                'Servidor 1.0.0
                cMostrarMision vecParametros(0), 0
            Else
                cMostrarMision vecParametros(0), CInt(vecParametros(1))
            End If
        Case msgAckAltaJugador
            cConfirmarAlta CInt(vecParametros(1)), CStr(vecParametros(2)), CInt(vecParametros(0))
        Case msgOrdenRonda
            cActualizarRonda vecParametros
'        Case msgComienzoTurno Es redundante con msgInicioTurno
'            cInicioTurno CInt(vecParametros(0)), CInt(vecParametros(1)), CBool(vecParametros(2))
        Case msgChatEntrante
            cMensajeChatEntrante CInt(vecParametros(0)), vecParametros(1)
        Case msgPartidasGuardadas
            cTipoPartidasGuardadas vecParametros
        Case msgAYA
            cEnviarIAA
        Case msgOpciones
            cRecibirOpciones vecParametros
        Case msgOpcionesDefault
            cRecibirOpcionesDefault vecParametros
        Case msgBajaAdm
            cConfirmarBajaAdm
        Case msgInicioTurno
            '###Version
            If UBound(vecParametros) = 1 Then
                'servidor 1.0.0
                cInicioTurno CInt(vecParametros(0)), CInt(vecParametros(1)), False
            Else
                cInicioTurno CInt(vecParametros(0)), CInt(vecParametros(1)), CBool(vecParametros(2))
            End If
        Case msgAckInicioPartida
            cConfirmarInicioPartida
        Case msgTropasDisponibles
            cTropasDisponibles vecParametros
        Case msgAckFinTurno
            cConfirmarFinTurno CBool(vecParametros(0))
        Case msgTipoRonda
            cActualizarTipoRonda CInt(vecParametros(0))
        Case msgAckAtaque
            cAckAtacar CInt(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2)), _
                       CInt(vecParametros(3)), CInt(vecParametros(4)), CInt(vecParametros(5)), _
                       CByte(vecParametros(6)), CInt(vecParametros(7)), CInt(vecParametros(8)), _
                       CByte(vecParametros(9)), CInt(vecParametros(10)), CInt(vecParametros(11))
        Case msgAckMovimiento
            '###Version
            If UBound(vecParametros) = 6 Then
                'Servidor 1.0.0
                cAckMover CByte(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2)), _
                          CByte(vecParametros(3)), CInt(vecParametros(4)), CInt(vecParametros(5)), _
                          CInt(vecParametros(6)), 0
            Else
                cAckMover CByte(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2)), _
                          CByte(vecParametros(3)), CInt(vecParametros(4)), CInt(vecParametros(5)), _
                          CInt(vecParametros(6)), CInt(vecParametros(7))
            End If
        Case msgTarjeta
            cMostrarTarjeta CByte(vecParametros(0)), CByte(vecParametros(1)), _
                            IIf(Trim(UCase(vecParametros(2))) = "S", True, False)
        Case msgTarjetasJugador
            cActualizarTarjetasJugador CInt(vecParametros(0)), CInt(vecParametros(1))
        Case msgAckCobroTarjeta
            cAckCobrarTarjeta CByte(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2))
        Case msgAckCanjeTarjeta
            cAckCanjearTarjeta CByte(vecParametros(0)), CByte(vecParametros(1)), CByte(vecParametros(2))
        Case msgMisionCumplida
            '###Version
            If UBound(vecParametros) = 1 Then
                'servidor 1.0.0
                cMisionCumplida CInt(vecParametros(0)), CStr(vecParametros(1)), 0
            Else
                cMisionCumplida CInt(vecParametros(0)), CStr(vecParametros(1)), CInt(vecParametros(2))
            End If
        Case msgEstadoTurnoCliente
            '###Version
            If UBound(vecParametros) = 0 Then
                'servidor 1.0.0
                cActualizarEstadoTurno CInt(vecParametros(0)), ""
            Else
                cActualizarEstadoTurno CInt(vecParametros(0)), CStr(vecParametros(1))
            End If
        Case msgAckGuardarPartida
            cAckGuardarPartida CInt(vecParametros(0))
        Case msgCanjesJugador
            cActualizarCanjesJugador CInt(vecParametros(0)), CInt(vecParametros(1))
        Case msgLog
            cRecibirLog vecParametros
        Case msgIpServidor
            cMostrarIpServidor CStr(vecParametros(0))
        Case msgVersionServidor
            cRecibirVersionServidor CInt(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2))
        Case msgLimitrofes
            cRecibirLimitrofes vecParametros
        Case msgAckPausarPartida
            cAckPausarPartida
        Case msgAckContinuarPartida
            cAckContinuarPartida
        Case enuTipoMensaje.msgError
            cMostrarError CStr(vecParametros(0)), CInt(vecParametros(1))
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "DistribuirMensaje - msg:" & intTipoMensaje, "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cActualizarPais(byPais As Byte, intColor As Integer, intCantidad As Integer, intOrigen As enuOrigenMsgPais, intTropasFijas As Integer)
    On Error GoTo ErrorHandle
    
    'Actualiza el mapa
    frmMapa.objPais(byPais).Color = intColor
    frmMapa.objPais(byPais).CantTropas = intCantidad
    frmMapa.objPais(byPais).TropasFijas = intTropasFijas
    
    'Actualiza el detalle del jugador seleccionado
    'frmJugadores.ActualizarDetalleJugadores intColor
    frmJugadores.ActualizarDetalleJugadorSeleccionado
    
    'Segun el origen del mensaje realiza un efecto determinado
    Select Case intOrigen
        Case enuOrigenMsgPais.orAgregado, enuOrigenMsgPais.orCobroTarjeta
            frmMapa.EfectoDestelloPais byPais
    End Select
        
    Exit Sub
ErrorHandle:
    ReportErr "cActualizarPais", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMostrarMision(strNuevaMision As String, intCodigoMision As Integer)
    On Error GoTo ErrorHandle
    
    '###Version
    If intCodigoMision = 0 Then
        GstrMision = strNuevaMision
    Else
        GstrMision = ObtenerTextoRecurso(intCodigoMision)
    End If
    
    If GstrMision = "" Then
        GstrMision = strNuevaMision
    End If
    frmMision.Actualizar
    frmMision.Visible = True
    mdifrmPrincipal.ActualizarMenu

    Exit Sub
ErrorHandle:
    ReportErr "cMostrarMision", "mdlCliente", Err.Description, Err.Number, Err.Source
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
    
    frmJugadores.ActualizarRonda

    Exit Sub
ErrorHandle:
    ReportErr "cActualizarRonda", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMensajeChatEntrante(intRemitente As Integer, strMensaje As String)

    On Error GoTo ErrorHandle

    'Posiciona el punto de inserción al final de la cadena
    frmChat.txtRecibido.SelStart = Len(frmChat.txtRecibido.Text)
    frmChat.txtRecibido.SelLength = 1
    
    If frmChat.txtRecibido.Text <> "" Then
        frmChat.txtRecibido.SelText = vbCrLf
    End If
    
    'Color del remitente
    frmChat.txtRecibido.SelColor = GvecColores(intRemitente)
    frmChat.txtRecibido.SelBold = True

    frmChat.txtRecibido.SelText = "" & GvecJugadores(intRemitente).strNombre & ": "
    
    frmChat.txtRecibido.SelColor = vbBlack
    frmChat.txtRecibido.SelBold = False
    
    'frmChat.txtRecibido.SelColor = RGB(0, 255, 0) 'vbBlack
    frmChat.txtRecibido.SelText = strMensaje
    
    'Posiciona el punto de inserción al final de la cadena
    frmChat.txtRecibido.SelStart = Len(frmChat.txtRecibido.Text)
    frmChat.txtRecibido.SelLength = 1
    
    Exit Sub
ErrorHandle:
    ReportErr "cMensajeChatEntrante", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMensajeChatSaliente(strMensaje As String)
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgChatSaliente, strMensaje)

    Exit Sub
ErrorHandle:
    ReportErr "cMensajeChatSaliente", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cRefrescarConexiones(pVecJugadores() As String)
    'Actualiza en el cliente las conexiones actuales
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim intCantJugadoresConectados As Integer
    'Dim intTipoJugadoresConectados As enuTipoJugadoresConectados
    
    'Toma del mensaje el Tipo que se encuentra en la última posición
    'del vector Jugadores
    GintTipoJugadoresConectados = pVecJugadores(UBound(pVecJugadores))
    
    'Muestra los jugadores conectados deshabilitando/habilitando
    'las opciones de color
    intCantJugadoresConectados = 0
    For i = 1 To UBound(GvecJugadores)
        'Actualiza el vector de jugadores
        GvecJugadores(i).strNombre = pVecJugadores(i - 1)
        GvecJugadores(i).intEstado = pVecJugadores(i + 5)
        GvecJugadores(i).byTipoJugador = pVecJugadores(i + 11)
        GvecJugadores(i).strDirIP = pVecJugadores(i + 17)
        GvecJugadores(i).strVersion = pVecJugadores(i + 23)
        
        'Cuenta la cant de jugadores conectados
        'para luego saber si habilitar o no el boton "Iniciar Parida"
        If GvecJugadores(i).strNombre <> "" Then
            intCantJugadoresConectados = intCantJugadoresConectados + 1
        End If
    Next i
    
    frmSeleccionColor.Actualizar GintTipoJugadoresConectados
    'El botón "Iniciar Partida" solo se habilita
    'si hay mas de un jugador conectado.
    If intCantJugadoresConectados > 1 Then
        frmSeleccionColor.cmdIniciarPartida.Enabled = True
    Else
        frmSeleccionColor.cmdIniciarPartida.Enabled = False
    End If
    frmJugadores.Actualizar
    
    Exit Sub
ErrorHandle:
    ReportErr "cRefrescarConexiones", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub
                                        
Public Sub cConfirmarAlta(intColorAsignado As Integer, strNombreAsignado As String, intCodAck As enuAckAltaJugador)
    On Error GoTo ErrorHandle
    Dim strMensaje As String
    Dim i As Integer
    
    Select Case intCodAck
        Case enuAckAltaJugador.ackOk
            '###E
            ActualizarEstadoCliente eveConfirmacionAlta
            
            GintMiColor = intColorAsignado
            GstrMiNombre = strNombreAsignado
            
            'Muestra el color en la barra de estado
            mdifrmPrincipal.StatusBar1.Panels(panIcoColor).Picture = mdifrmPrincipal.imgLstFichas.ListImages(intColorAsignado).Picture
            
            'Cambia el aspecto de la pantalla de selección si ya hay un color seleccionado
            frmSeleccionColor.CambiarEstadoPantalla 1
            
            'LIMPIEZA DE VALORES ANTERIORES
            'Limpia el formulario de Tarjetas
            For i = 1 To UBound(GvecTarjetas)
                GvecTarjetas(i).byPais = 0
                frmTarjetas.shpTarjetaSel(i - 1).Visible = False
            Next i
            frmTarjetas.Actualizar
            
            'Limpia el formulario de Jugadores
            For i = 1 To UBound(GvecJugadores)
                GvecJugadores(i).intCanje = 0
                GvecJugadores(i).intCantidadTarjetas = 0
                GvecJugadores(i).intTropasDisponibles = 0
                GvecJugadores(i).strNombre = ""
            Next i
            frmJugadores.Actualizar

        Case enuAckAltaJugador.ackColorUsado
            strMensaje = ObtenerTextoRecurso(CintGralMsgColorAsignado) 'El color seleccionado ya ha sido asignado a otro usuario.
            frmSeleccionColor.StatusBar1.SimpleText = ObtenerTextoRecurso(CintGralMsgColorAsignadoCaption) '"Seleccione un nuevo color."
        Case enuAckAltaJugador.ackNombreUsado
            strMensaje = ObtenerTextoRecurso(CintGralMsgNombreAsignado) '"El nombre seleccionado ya ha sido asignado a otro usuario."
            frmSeleccionColor.StatusBar1.SimpleText = ObtenerTextoRecurso(CintGralMsgNombreAsignadoCaption) '"Ingrese un nuevo nombre."
        Case enuAckAltaJugador.ackNombreYColorUsados
            strMensaje = ObtenerTextoRecurso(CintGralMsgNombreColorAsignados) '"El nombre y el color seleccionados ya han sido asignados a otro/s usuario/s"
            frmSeleccionColor.StatusBar1.SimpleText = ObtenerTextoRecurso(CintGralMsgNombreColorAsignadosCaption) '"Seleccione un nuevo color e ingrese un nuevo nombre."
        Case enuAckAltaJugador.ackColorInexistente
            strMensaje = ObtenerTextoRecurso(CintGralMsgColorInvalido) '"El color seleccionado no corresponde a ningún jugador de la partida."
            frmSeleccionColor.StatusBar1.SimpleText = ObtenerTextoRecurso(CintGralMsgColorInvalidoCaption) '"Seleccione un nuevo color."
        Case enuAckAltaJugador.ackColorConectado
            strMensaje = ObtenerTextoRecurso(CintGralMsgColorUtilizado) '"El color seleccionado ya está siendo utilizado por otro jugador."
            frmSeleccionColor.StatusBar1.SimpleText = ObtenerTextoRecurso(CintGralMsgColorUtilizadoCaption) '"Ingrese un nuevo color."
        Case enuAckAltaJugador.ackServidorPausado
            strMensaje = ObtenerTextoRecurso(CintGralMsgServidorPausado) '"No es posible conectarse, porque el Servidor se encuentra Pausado."
            'Solo en este error se cierra el form...
            GblnSeCierra = True 'Para que el form no pregunte nada al cerrarse
            Unload frmSeleccionColor
            GblnSeCierra = False
            MsgBox strMensaje, vbExclamation, ObtenerTextoRecurso(CintGralMsgServidorPausadoCaption) '"Servidor no disponible."
            cDesconectar
            Exit Sub
        Case Else
    End Select
    
    'Parte comun a todos los errores
    If intCodAck <> enuAckAltaJugador.ackOk Then
        frmSeleccionColor.cmdAceptar.Enabled = True
        frmSeleccionColor.optColor(intColorAsignado).Value = False
        MsgBox strMensaje, vbExclamation, ObtenerTextoRecurso(CintGralMsgColorNombreOtroCaption) 'Resultado de la Selección
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cConfirmarAlta", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cConectarAdm(strServer As String, intPuerto As Integer, blnActivarServidor As Boolean)
    On Error GoTo ErrorHandle
    Dim dblExito As Double
    'Levanta el servidor y se conecta automáticamente
    
    'Levanta el servidor, solo si es localhost
    '### no levanta el servidor si pone la direccion IP real de la máquina
    If blnActivarServidor Then
        dblExito = Shell(App.Path & "\..\Servidor\TEGNet_Server.exe " & CStr(intPuerto))
    End If
    
    '###CHEQUAR ERROR
    
    'Intenta conectarse al servidor
    cConectar strServer, intPuerto
    
    Exit Sub
ErrorHandle:
    ReportErr "cConectarAdm", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub
    
Public Sub cConectar(strServer As String, intPort As Integer)
    On Error GoTo ErrorHandle
    
    'Conecta al cliente con el servidor
    mdifrmPrincipal.Winsock1.RemoteHost = strServer
    mdifrmPrincipal.Winsock1.RemotePort = intPort
    
    'Almacena las variables globales
    GstrServidor = strServer
    GintPuerto = intPort
    
    ' Invoca el método Connect para iniciar
    ' una conexión.
    mdifrmPrincipal.Winsock1.Connect
    
    'Valores por defecto de las variables de conexion
    GlngVersionServidor = 0
    
    Exit Sub
ErrorHandle:
    Screen.MousePointer = vbDefault
    If frmComienzo.Visible And frmComienzo.fraEtapas(enuEtapas.etaUnirse).Visible Then
        frmComienzo.Label4.Caption = ObtenerTextoRecurso(CintComienzoErrorConectando)
    Else
        ReportErr "cConectar", "mdlCliente", Err.Description, Err.Number, Err.Source
    End If
        
    '###E
    If GEstadoCliente < estConectado Then
        'Cierra el puerto para poder volver a conectarse
        cDesconectar
    End If

End Sub

Public Sub cConfirmarAdm(intCantConexiones As Integer)
    On Error GoTo ErrorHandle
    
    'Hablita al administrador
    SetearAdm True
    
    'Si está activa la ventana de selección de colores,
    'se actualiza en modo administrador
    If frmSeleccionColor.Visible Then
        frmSeleccionColor.CambiarEstadoPantalla 1
    End If
    
    'ATENCION
    '--------
    'Este messagebox provoca error (no muestra las partidas guardadas)
    'pero solo cuando se ejecuta desde el fuente
    '"Usted ha sido designado Administrador de la Partida."
    MsgBox ObtenerTextoRecurso(CintGralMsgAdmDesignado), vbInformation, ObtenerTextoRecurso(CintGralMsgAdmDesignadoCaption)
    
    If GEstadoCliente >= estValidado Or intCantConexiones > 1 Then
    
    Else
        'Envía el mensaje de Tipo de Partida
        cTipoPartida
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cConfirmarAdm", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cTipoPartida()
    On Error GoTo ErrorHandle
    'Informa al servidor el tipo de partida
    'seleccionado por el cliente (Nueva o guardada)
    EnviarMensaje ArmarMensajeParam(msgTipoPartida, TipoPartida)
    
    Exit Sub
ErrorHandle:
    ReportErr "cTipoPartida", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cTipoPartidasGuardadas(vecPartidas() As String)
    On Error GoTo ErrorHandle
    'Verifica el último valor del vector para saber a
    'que proceso reenviar el mensaje
    Dim intTipo As enuTipoPartidaGuardada
    
    'Toma el último valor del vector
    intTipo = CInt(vecPartidas(UBound(vecPartidas)))
    
    'Elimina el último valor del vector
    ReDim Preserve vecPartidas(UBound(vecPartidas) - 1)
    
    'LLama al procedimiento que corresponda
    If intTipo = tpgIniciarPartida Then
        cPartidasGuardadas vecPartidas
    Else
        cMostrarPartidasGuardadas vecPartidas
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cTipoPartidasGuardadas", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cPartidasGuardadas(vecPartidas() As String)
    On Error GoTo ErrorHandle
    'Recibe del servidor las partidas guardadas y las
    'muestra en pantalla
    Dim i As Integer
    Dim existeUltima As Boolean
    Dim datFecha As Date
    Dim strNombre As String
    
    'Pasa a la etapa 3 del Wizard
    frmComienzo.MostrarEtapa etaGuardada
    
    'Limpia el List
    frmComienzo.lstPartidasGuardadas.Clear
    ReDim GvecNombresPartidas(UBound(vecPartidas))
    
    existeUltima = False
    For i = LBound(vecPartidas) To UBound(vecPartidas)
        'Si el elemento está vacío
        If Trim(vecPartidas(i)) <> "" Then
            strNombre = Mid$(vecPartidas(i), 11)
            datFecha = CDate(Left(vecPartidas(i), 10))
            If strNombre = strNombrePartidaActual Then
                existeUltima = True
            Else
                frmComienzo.lstPartidasGuardadas.AddItem CStr(datFecha) & " - " & strNombre
                'Guarda los nombres en un vector paralelo, para recuperarlo facilmente
                GvecNombresPartidas(frmComienzo.lstPartidasGuardadas.ListCount - 1) = strNombre
            End If
        End If
    Next i
    
    'Oculta la opcion de la ultima partida si la misma no existe
    If Not existeUltima Then
        frmComienzo.optUltima.Visible = False
        frmComienzo.optVieja.Value = True
        frmComienzo.optVieja.Visible = False
        frmComienzo.lblPartidasGuardadas.Visible = True
    Else
        frmComienzo.optUltima.Visible = True
        frmComienzo.optUltima.Value = True
        frmComienzo.optVieja.Visible = True
        frmComienzo.lblPartidasGuardadas.Visible = False
    End If
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrorHandle:
    ReportErr "cPartidasGuardadas", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cBajarServidor()
    On Error GoTo ErrorHandle
    
    'Baja el servidor
    EnviarMensaje ArmarMensajeParam(msgBajarServidor)
        
    Exit Sub
ErrorHandle:
    ReportErr "cBajarServidor", "mdlCliente", Err.Description, Err.Number, Err.Source

End Sub

Public Sub cAltaJugador(intColorSeleccionado As Integer, strNickNameSeleccionado As String)
    'Envia al servidor el color y NickName seleccionado por el jugador
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgAltaJugador, CStr(intColorSeleccionado), strNickNameSeleccionado)
        
    Exit Sub
ErrorHandle:
    ReportErr "cAltaJugador", "mdlCliente", Err.Description, Err.Number, Err.Source

End Sub

Public Sub cEnviarIAA()
    'Envia al servidor el mensaje que indica que el cliente está vivo
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgIAA)
        
    Exit Sub
ErrorHandle:
    ReportErr "cEnviarIAA", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cRecibirOpciones(vecOpcionesMsg() As String)
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim blnMostrarEnPantalla
    
    'Toma del último valor del vector el flag de mostrar en pantalla
    blnMostrarEnPantalla = CBool(vecOpcionesMsg(UBound(vecOpcionesMsg)))
    ReDim Preserve vecOpcionesMsg(UBound(vecOpcionesMsg) - 1)
    
    ReDim GvecOpciones(0 To (UBound(vecOpcionesMsg) - LBound(vecOpcionesMsg) + 1) / 2 - 1)
    
    'En base al vector recibido carga un vector global con los campos separados
    For i = LBound(vecOpcionesMsg) To UBound(vecOpcionesMsg) Step 2 'Toma de a pares
        GvecOpciones(Int(i / 2)).Id = vecOpcionesMsg(i)
        GvecOpciones(Int(i / 2)).Valor = vecOpcionesMsg(i + 1)
    Next i
    
    'Si soy el administrador (lo muestra solo la primera vez)
    'o bien
    'A los no administradores se les muestran las opciones cada vez que
    'el Adm las cambia (salvo la primera)
    If (GsoyAdministrador And blnMostrarEnPantalla) _
    Or (Not GsoyAdministrador And Not blnMostrarEnPantalla) _
    Then
        frmOpciones.Visible = True
    End If
    
    Exit Sub
ErrorHandle:
        ReportErr "sRecibirOpciones", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cRecibirOpcionesDefault(vecOpcionesMsg() As String)
    'Guarda las opciones por defecto en un vector global
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    ReDim GvecOpcionesDefault(0 To (UBound(vecOpcionesMsg) - LBound(vecOpcionesMsg) + 1) / 2 - 1)
    
    'En base al vector recibido carga un vector global con los campos separados
    For i = LBound(vecOpcionesMsg) To UBound(vecOpcionesMsg) Step 2 'Toma de a pares
        GvecOpcionesDefault(Int(i / 2)).Id = vecOpcionesMsg(i)
        GvecOpcionesDefault(Int(i / 2)).Valor = vecOpcionesMsg(i + 1)
    Next i
    
    Exit Sub
ErrorHandle:
        ReportErr "sRecibirOpcionesDefault", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cEnviarOpciones()
    'Toma de la pantalla de Opciones las opciones de la partida actual
    On Error GoTo ErrorHandle
    Dim vecOpcionesMsg() As String
    Dim i As Integer
        
    'Arma el vector de a pares para el mensaje
    ReDim vecOpcionesMsg(LBound(GvecOpciones) To UBound(GvecOpciones))
    For i = LBound(GvecOpciones) To UBound(GvecOpciones)
        vecOpcionesMsg(i) = GvecOpciones(i).Id & "|" & GvecOpciones(i).Valor
    Next i
    
    EnviarMensaje ArmarMensaje(msgOpciones, vecOpcionesMsg)
    
    Exit Sub
ErrorHandle:
    ReportErr "cEnviarOpciones", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cEnviarOpcionesDefault()
    'Toma de la pantalla de Opciones las opciones de la partida actual
    'y las guarda como default
    On Error GoTo ErrorHandle
    Dim vecOpcionesMsg() As String
    Dim i As Integer
        
    'Arma el vector de a pares para el mensaje
    ReDim vecOpcionesMsg(LBound(GvecOpcionesDefault) To UBound(GvecOpcionesDefault))
    For i = LBound(GvecOpcionesDefault) To UBound(GvecOpcionesDefault)
        vecOpcionesMsg(i) = GvecOpcionesDefault(i).Id & "|" & GvecOpcionesDefault(i).Valor
    Next i
    
    EnviarMensaje ArmarMensaje(msgOpcionesDefault, vecOpcionesMsg)
    
    Exit Sub
ErrorHandle:
    ReportErr "cEnviarOpcionesDefault", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub MostrarOpciones(vecOpcionesAux() As typOpcion)
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    For i = LBound(vecOpcionesAux) To UBound(vecOpcionesAux)
        Select Case vecOpcionesAux(i).Id
            Case enuOpciones.opTurnoDuracion
                frmOpciones.txtDuracion.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opTurnoTolerancia
                frmOpciones.txtTolerancia.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opRondaTropas1ra
                frmOpciones.txtPrimeraRonda.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opRondaTropas2da
                frmOpciones.txtSegundaRonda.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opRondaTipo
                If Trim(UCase(vecOpcionesAux(i).Valor)) = "R" Then
                    'Primero Rotativo
                    frmOpciones.optRotativa.Value = True
                Else
                    'Ronda fija
                    frmOpciones.optFija.Value = True
                End If
            Case enuOpciones.opBonusTarjetaPropia
                frmOpciones.txtTarjetaPropio.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opBonusAfrica
                frmOpciones.txtAfrica.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opBonusANorte
                frmOpciones.txtANorte.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opBonusASur
                frmOpciones.txtASur.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opBonusAsia
                frmOpciones.txtAsia.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opBonusEuropa
                frmOpciones.txtEuropa.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opBonusOceania
                frmOpciones.txtOceania.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opMisionDestruir
                If Trim(UCase(vecOpcionesAux(i).Valor)) = "S" Then
                    'Con Destruir
                    frmOpciones.chkDestruir.Value = 1
                Else
                    'Sin Destruir
                    frmOpciones.chkDestruir.Value = 0
                End If
            Case enuOpciones.opMisionTipo
                If Trim(UCase(vecOpcionesAux(i).Valor)) = "C" Then
                    'Conquistar el mundo
                    frmOpciones.optConquistarMundo.Value = True
                Else
                    'Por misiones
                    frmOpciones.optMisiones.Value = True
                End If
            Case enuOpciones.opMisionObjetivoComun
                frmOpciones.txtObjetivoComun.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opTropasInicial
                frmOpciones.txtTropasInicio.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opCanjeNro1
                frmOpciones.txtBonusCanje1.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opCanjeNro2
                frmOpciones.txtBonusCanje2.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opCanjeNro3
                frmOpciones.txtBonusCanje3.Text = vecOpcionesAux(i).Valor
            Case enuOpciones.opCanjeIncremento
                frmOpciones.txtBonusCanjeIncremento.Text = vecOpcionesAux(i).Valor
        End Select
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "MostrarOpciones", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CapturarOpciones(ByRef vecOpcionesAux() As typOpcion)
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    For i = LBound(vecOpcionesAux) To UBound(vecOpcionesAux)
        Select Case vecOpcionesAux(i).Id
            Case enuOpciones.opTurnoDuracion
                vecOpcionesAux(i).Valor = frmOpciones.txtDuracion.Text
            Case enuOpciones.opTurnoTolerancia
                vecOpcionesAux(i).Valor = frmOpciones.txtTolerancia.Text
            Case enuOpciones.opRondaTropas1ra
                vecOpcionesAux(i).Valor = frmOpciones.txtPrimeraRonda.Text
            Case enuOpciones.opRondaTropas2da
                vecOpcionesAux(i).Valor = frmOpciones.txtSegundaRonda.Text
            Case enuOpciones.opRondaTipo
                If frmOpciones.optRotativa.Value = True Then
                    'Primero Rotativo
                    vecOpcionesAux(i).Valor = "R"
                Else
                    'Ronda fija
                    vecOpcionesAux(i).Valor = "F"
                End If
            Case enuOpciones.opBonusTarjetaPropia
                vecOpcionesAux(i).Valor = frmOpciones.txtTarjetaPropio.Text
            Case enuOpciones.opBonusAfrica
                vecOpcionesAux(i).Valor = frmOpciones.txtAfrica.Text
            Case enuOpciones.opBonusANorte
                vecOpcionesAux(i).Valor = frmOpciones.txtANorte.Text
            Case enuOpciones.opBonusASur
                vecOpcionesAux(i).Valor = frmOpciones.txtASur.Text
            Case enuOpciones.opBonusAsia
                vecOpcionesAux(i).Valor = frmOpciones.txtAsia.Text
            Case enuOpciones.opBonusEuropa
                vecOpcionesAux(i).Valor = frmOpciones.txtEuropa.Text
            Case enuOpciones.opBonusOceania
                vecOpcionesAux(i).Valor = frmOpciones.txtOceania.Text
            Case enuOpciones.opMisionDestruir
                If frmOpciones.chkDestruir.Value = 1 Then
                    'Con Destruir
                    vecOpcionesAux(i).Valor = "S"
                Else
                    'Sin Destruir
                    vecOpcionesAux(i).Valor = "N"
                End If
            Case enuOpciones.opMisionTipo
                If frmOpciones.optConquistarMundo.Value = True Then
                    'Conquistar el mundo
                    vecOpcionesAux(i).Valor = "C"
                Else
                    'Por misiones
                    vecOpcionesAux(i).Valor = "M"
                End If
            Case enuOpciones.opMisionObjetivoComun
                vecOpcionesAux(i).Valor = frmOpciones.txtObjetivoComun.Text
            Case enuOpciones.opTropasInicial
                vecOpcionesAux(i).Valor = frmOpciones.txtTropasInicio.Text
            Case enuOpciones.opCanjeNro1
                vecOpcionesAux(i).Valor = frmOpciones.txtBonusCanje1.Text
            Case enuOpciones.opCanjeNro2
                vecOpcionesAux(i).Valor = frmOpciones.txtBonusCanje2.Text
            Case enuOpciones.opCanjeNro3
                vecOpcionesAux(i).Valor = frmOpciones.txtBonusCanje3.Text
            Case enuOpciones.opCanjeIncremento
                vecOpcionesAux(i).Valor = frmOpciones.txtBonusCanjeIncremento.Text
        End Select
    Next i
    
    
    Exit Sub
ErrorHandle:
        ReportErr "CapturarOpciones", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cPedirConexionesActuales()
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgPedirConexionesActuales)
    
    Exit Sub
ErrorHandle:
    ReportErr "cPedirConexionesActuales", "mdlInterfase", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cDesconectar()
    'Desconecta al cliente sin bajar el servidor
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    'Cambia el icono de la barra
    mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbConexion).Image = 12
    'Habilita y Deshabilita las opciones de menu
    mdifrmPrincipal.mnuPartidaDesconectar.Enabled = False
    mdifrmPrincipal.mnuPartidaConectar.Enabled = True

    
    mdifrmPrincipal.Winsock1.Close
    'mdifrmPrincipal.Toolbar1.Buttons(1).Value = tbrUnpressed
    mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panIcoColor).Picture = Nothing
    mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panTipoRonda).Text = ""
    
    'Desactiva el timer
    mdifrmPrincipal.Timer1.Interval = 0
    mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panTimer).Text = ""
    
    'mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panEstado).Text = ""
    GintMiColor = 0
    ActualizarEstadoCliente eveCierreConexion
    
    If GsoyAdministrador Then
        'Deshabilita al administrador
        SetearAdm False
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cDesconectar", "mdlInterfase", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMostrarError(strMensajeError As String, intCodigoError As Integer)
    'Muestra en pantalla el mensaje de error recibido del servidor
    On Error GoTo ErrorHandle
    Dim strMensaje As String
    
    strMensaje = ObtenerTextoRecurso(intCodigoError + enuIndiceArchivoRecurso.pmsErrores)
    'Si no encuentra el mensaje de error en el archivo de recursos, muestra la descripcion
    'que llega en el mensaje
    Mensaje IIf(strMensaje = "", strMensajeError, strMensaje), ObtenerTextoRecurso(CintGralMsgErrServidorCaption) '"Mensaje del Servidor"
    
    Exit Sub
ErrorHandle:
    ReportErr "cMostrarError", "mdlInterfase", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cCambiarAdm(intNuevoAdm As Integer)
    'Envía al servidor la petición de cambio de administrador
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgCambiarAdm, intNuevoAdm)
    
    Exit Sub
ErrorHandle:
    ReportErr "cCambiarAdm", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cConfirmarBajaAdm()
    'Informa al cliente que ya no es mas Administrador
    On Error GoTo ErrorHandle
    
    SetearAdm False
    '"Usted ha dejado de ser Administrador"
    MsgBox ObtenerTextoRecurso(CintGralMsgAdmNoDesignado), vbInformation, ObtenerTextoRecurso(CintGralMsgAdmNoDesignadoCaption)
    
    Exit Sub
ErrorHandle:
    ReportErr "cConfirmarBajaAdm", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cIniciarPartida()
    'Indica al servidor que inicie la partida
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgIniciarPartida)
    
    Exit Sub
ErrorHandle:
    ReportErr "cIniciarPartida", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cConfirmarInicioPartida()
    'Cambia el estado a jugando
    On Error GoTo ErrorHandle
    
    'Carga los paises limítrofes (solo si no se cargaron todavía)
    frmMapa.CargarLimitrofesArchivo
    
    ActualizarEstadoCliente eveInicioPartida
    
    'Habilita/Deshabilita la opcion Pausa, si son todos Robots o no.
    EvaluarHabilitacionPausa
    
    VisibilidadMapa True
    
    '###
    'Si el formulario esta cargado hay que descargarlo
    'frmSeleccionColor.Hide
    'If frmSeleccionColor.Visible = True Then
        Unload frmSeleccionColor
    'End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cConfirmarInicioPartida", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cInicioTurno(intColorTurno As Integer, intTimerTurno As Integer, blnResincronizacion As Boolean)
    On Error GoTo ErrorHandle
    
    Dim byPais As Byte
    
    GintColorActual = intColorTurno
    frmJugadores.ActualizarTurno GintColorActual
    
    'Muestra el tiempo disponible
    If intTimerTurno = CintValorTimerInfinito Then
        'Si el timer es infinito no activa el timer
        mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panTimer).Text = "-"
        mdifrmPrincipal.Timer1.Interval = 0
    Else
        mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panTimer).Text = intTimerTurno
        GintTimerTurno = intTimerTurno
        'Activa el timer
        mdifrmPrincipal.Timer1.Interval = 1000
        GsngInicioTimerTurno = Timer
    End If
    
    'Si el turno es mio
    If intColorTurno = GintMiColor Then
        'Solo cambia el estado y limpia los fijos si no se trata de una resincronizacion
        If Not blnResincronizacion Then
            'Limpia las cantidades de tropas fijas de cada pais
            For byPais = 1 To frmMapa.objPais.Count - 1
                frmMapa.objPais(byPais).TropasFijas = 0
            Next byPais
            
            'Cambia el estado
            If GintTipoRonda = trInicio Or GintTipoRonda = trRecuento Then
                ActualizarEstadoCliente eveInicioTurnoRecuento
            Else
                ActualizarEstadoCliente eveInicioTurnoAccion
            End If
        End If
        
        'Actualiza los iconos del mouse del mapa
        frmMapa.ActualizarIconosMouse
        
        'Resetea la cantidad de conquistas
        GintCantConquistas = 0
        
        'Resalta icono del systray
        intFlagSysTray = 10 'Titila 5 veces
        mdifrmPrincipal.tmrSysTray.Enabled = True
        
        Beep
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cInicioTurno", "mdlCliente", Err.Description, Err.Number, Err.Source
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
    
    If CInt(vecMensaje(0)) = GintMiColor And blnExisteNoLibre Then
        'Si al llegar el mensaje, el mismo se refiere al jugador activo
        'y hay tropas que corresponden solo a un continente determinado
        'muestra el formulario con el detalle de las tropas disponibles
        'forzando la selección del jugador
        frmJugadores.optDetalle(GvecJugadores(GintMiColor).intOrdenRonda - 1).Value = False
        frmJugadores.optDetalle(GvecJugadores(GintMiColor).intOrdenRonda - 1).Value = True
        MostrarFormulario frmTropasDisponibles
        mdifrmPrincipal.ActualizarMenu
    Else
        'Si al llegar el mensaje, el mismo se refiere a otro jugador,
        'actualiza el formulario con el detalle de las tropas disponibles
        
        'Busca la opción seleccionada en el formulario de jugadores
        For i = 0 To frmJugadores.optDetalle.Count - 1
            If frmJugadores.optDetalle(i).Value = True Then
                'Fuerza el evento click para que actualice los valores
                frmJugadores.optDetalle(i).Value = False
                frmJugadores.optDetalle(i).Value = True
            End If
        Next i
    End If
        
    
    Exit Sub
ErrorHandle:
    ReportErr "cTropasDisponibles", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMostrarTropasDisponibles(intColor As Integer)
    'Muestra el detalle de las tropas disponibles
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    For i = LBound(GvecJugadores(intColor).vecDetalleTropasDisponibles) To UBound(GvecJugadores(intColor).vecDetalleTropasDisponibles)
        frmTropasDisponibles.lblTDI(i).Caption = GvecJugadores(intColor).vecDetalleTropasDisponibles(i)
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "cMostrarTropasDisponibles", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cFinTurno()
    'Informa al servidor el fin voluntario del turno
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgFinTurno)
    
    Exit Sub
ErrorHandle:
    ReportErr "cFinTurno", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cConfirmarFinTurno(blnExpiroTimer As Boolean)
    On Error GoTo ErrorHandle
    
    'Desactiva el timer del cliente
    mdifrmPrincipal.Timer1.Interval = 0
    
    'Vuelve el systray a la normalidad
    mdifrmPrincipal.tmrSysTray.Enabled = False
    SysTrayChangeIcon mdifrmPrincipal.hwnd, mdifrmPrincipal.imgSysTrayOK
    
    'Cambia el estado a esperando turno
    ActualizarEstadoCliente eveAckFinTurno
    
    If blnExpiroTimer Then
        'Avisa en pantalla
        'El tiempo del turno ha expirado
        Mensaje ObtenerTextoRecurso(CintGralMsgTurnoExpirado), ObtenerTextoRecurso(CintGralMsgTurnoExpiradoCaption)
    End If
    
    'Oculta los paises seleccionados
    frmMapa.DeseleccionarDestino
    frmMapa.DeseleccionarOrigen
    'Oculta todos los paises del mapa
    frmMapa.LimpiarMapa
    
    Exit Sub
ErrorHandle:
    ReportErr "cConfirmarFinTurno", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cActualizarTipoRonda(intTipoRonda As enuTipoRonda)
    On Error GoTo ErrorHandle
    
    GintTipoRonda = intTipoRonda
    
    'Regional - Carga del tipo de ronda
    'Dado el tipo de ronda muestra en la statusbar la descripción del mismo
    Select Case intTipoRonda
        Case enuTipoRonda.trInicio
            mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panTipoRonda).Text = ObtenerTextoRecurso(CintPrincipalRondaInicio) '"Ronda de Inicio"
        Case enuTipoRonda.trAccion
            mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panTipoRonda).Text = ObtenerTextoRecurso(CintPrincipalRondaAccion) '"Ronda de Acción"
        Case enuTipoRonda.trRecuento
            mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panTipoRonda).Text = ObtenerTextoRecurso(CintPrincipalRondaRecuento) '"Ronda de Recuento"
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "cActualizarTipoRonda", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cAgregarTropas(intPais As Byte, intCantidad As Integer)
    'Agrega tropas a un pais determinado
    On Error GoTo ErrorHandle
    
    'Valida que exista un pais seleccionado
    If intPais <= 0 Then
        Exit Sub
    End If
    
    'Valida que el pais sea propio
    If frmMapa.objPais(intPais).Color = GintMiColor Then
        'Valida la cantidad de tropas disponibles (el destino lo valida el servidor)
        If GvecJugadores(GintMiColor).intTropasDisponibles >= intCantidad And intCantidad > 0 Then
            'Envia el mensaje
            EnviarMensaje ArmarMensajeParam(msgAgregarTropas, intPais, intCantidad)
        End If
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "cAgregarTropas", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cAtacar(byPaisDesde As Byte, byPaisHasta As Byte)
    On Error GoTo ErrorHandle
    
    'valida que esten seleccionados el origen y el destino
    If byPaisDesde <= 0 Or byPaisHasta <= 0 Then
        Exit Sub
    End If
    
    'Valida que el pais origen sea mio
    If frmMapa.objPais(byPaisDesde).Color <> GintMiColor Then
        Exit Sub
    End If
    
    'Valida que el pais destino sea enemigo
    If frmMapa.objPais(byPaisHasta).Color = GintMiColor Then
        Exit Sub
    End If
    
    'Valida que el pais origen tenga mas de una tropa
    If frmMapa.objPais(byPaisDesde).CantTropas < 2 Then
        Exit Sub
    End If
    
    EnviarMensaje ArmarMensajeParam(msgAtaque, byPaisDesde, byPaisHasta)
    
    Exit Sub
ErrorHandle:
    ReportErr "cAtacar", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub Focalizar(ByRef ctrControl As Control)
    'Función que permite pasar el foco a un control sin tirar error si hay
    'un formulario modal abierto
    On Error Resume Next
    ctrControl.SetFocus
End Sub

Public Sub cAckAtacar(intDadoDesde1 As Integer, intDadoDesde2 As Integer, intDadoDesde3 As Integer, _
                      intDadoHasta1 As Integer, intDadoHasta2 As Integer, intDadoHasta3 As Integer, _
                      byPaisDesde As Byte, intColorDesde As Integer, intCantDesde As Integer, _
                      byPaisHasta As Byte, intColorHasta As Integer, intCantHasta As Integer)

    On Error GoTo ErrorHandle
    Dim strValorIngresado As String
    Dim intTropasAMover As Integer
    
    'Muestra los dados
    MostrarFormulario frmDados
    mdifrmPrincipal.ActualizarMenu
    
    'Después de mostrar los dados le pasa el foco al chat.
    If frmChat.Visible Then
        Focalizar frmChat.txtEnviado
    End If

    MostrarAtaque intDadoDesde1, intDadoDesde2, intDadoDesde3, _
                  intDadoHasta1, intDadoHasta2, intDadoHasta3, _
                  byPaisDesde, byPaisHasta
    
    'Si hubo conquista
    If intColorDesde = GintMiColor And intColorHasta = GintMiColor Then
        
        'Incrementa la cantidad de conquistas
        GintCantConquistas = GintCantConquistas + 1
        '### Habilita la opcion de tomar tarjeta
        
        intTropasAMover = 0
        'Si la cantidad de tropas en el origen es 2 no pregunta cuantas
        'tropas pasar
        If intCantDesde > 1 Then
            Do
                'Pregunta la cantidad de tropas a pasar
                '"Ingrese la cantidad de tropas a mover al pais conquistado"
                'strValorIngresado = InputBox(ObtenerTextoRecurso(CintGralMsgTropasMover), CompilarMensaje(ObtenerTextoRecurso(CintGralMsgTropasMoverCaption), Array(frmMapa.objPais(byPaisHasta).Nombre)), 1)
                intTropasAMover = frmConquista.TropasAMover(frmMapa.objPais(byPaisDesde).CantTropas - 1, frmMapa.objPais(byPaisHasta).Nombre)
            Loop While (intTropasAMover = 0)
        End If
        
        If intTropasAMover > 1 Then
            cMover byPaisDesde, byPaisHasta, intTropasAMover - 1, tmConquista
        End If
        
    End If
    
    'Actualiza los dos paises
    cActualizarPais byPaisDesde, intColorDesde, intCantDesde, orAtaque, 0
    cActualizarPais byPaisHasta, intColorHasta, intCantHasta, orAtaque, 0

    Exit Sub
ErrorHandle:
    ReportErr "cAckAtacar", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMover(byPaisDesde As Byte, byPaisHasta As Byte, intCantidadFichas As Integer, intTipoMovimiento As enuTipoMovimiento)
    On Error GoTo ErrorHandle
    
    'valida que esten seleccionados el origen y el destino
    If byPaisDesde <= 0 Or byPaisHasta <= 0 Then
        Exit Sub
    End If
    
    'Valida que el pais origen sea mio
    If frmMapa.objPais(byPaisDesde).Color <> GintMiColor Then
        Exit Sub
    End If
    
    If intTipoMovimiento <> tmConquista Then
        'Valida que el pais destino sea mio
        If frmMapa.objPais(byPaisHasta).Color <> GintMiColor Then
            Exit Sub
        End If
    Else
        'Valida que el pais destino sea enemigo
        If frmMapa.objPais(byPaisHasta).Color = GintMiColor Then
            Exit Sub
        End If
    End If
    
    'Valida que el pais origen tenga la cantidad de fichas necesarias
    If frmMapa.objPais(byPaisDesde).CantTropas <= intCantidadFichas Then
        Exit Sub
    End If
    
    EnviarMensaje ArmarMensajeParam(msgMovimiento, byPaisDesde, byPaisHasta, intCantidadFichas, intTipoMovimiento)
    
    Exit Sub
ErrorHandle:
    ReportErr "cMover", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cAckMover(byPaisDesde As Byte, intColorDesde As Integer, intCantDesde As Integer, _
                     byPaisHasta As Byte, intColorHasta As Integer, intCantHasta As Integer, _
                     intTipoMovimiento As enuTipoMovimiento, intCantidadTropas As Integer)
    On Error GoTo ErrorHandle
    
    'Cambia el estado del jugador que hizo el movimiento
    'Solo si el tipo de movimiento fue Movimiento
    If intTipoMovimiento = tmMovimiento Then
        ActualizarEstadoCliente eveAckMoverTropa
    
        'Efecto especial
        EfectoMover byPaisDesde, byPaisHasta
    
        'Actualiza los dos paises
        cActualizarPais byPaisDesde, intColorDesde, intCantDesde, orMovimiento, frmMapa.objPais(byPaisDesde).TropasFijas
        cActualizarPais byPaisHasta, intColorHasta, intCantHasta, orMovimiento, frmMapa.objPais(byPaisHasta).TropasFijas + intCantidadTropas
    Else
        'Actualiza los dos paises
        cActualizarPais byPaisDesde, intColorDesde, intCantDesde, orMovimiento, frmMapa.objPais(byPaisDesde).TropasFijas
        cActualizarPais byPaisHasta, intColorHasta, intCantHasta, orMovimiento, frmMapa.objPais(byPaisHasta).TropasFijas
    End If
    
    
    Exit Sub
ErrorHandle:
    ReportErr "cAckMover", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cTomarTarjeta()
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim blnPuedeTomarTarjeta As Boolean
    
    'Valida que al menos haya conquistado un pais
'''    If GintCantConquistas < 1 Then
'''        Exit Sub
'''    End If
    
    'Verifica que ya no tenga 5 tarjetas
    blnPuedeTomarTarjeta = False
    For i = 1 To UBound(GvecTarjetas)
        If GvecTarjetas(i).byPais = 0 Then
            blnPuedeTomarTarjeta = True
        End If
    Next i
    
    If Not blnPuedeTomarTarjeta Then
        Exit Sub
    End If
    
    EnviarMensaje ArmarMensajeParam(msgPedidoTarjeta)
    
    Exit Sub
ErrorHandle:
    ReportErr "cTomarTarjeta", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMostrarTarjeta(byPais As Byte, byFigura As Byte, blnCobrada As Boolean)
    'Muestra en pantalla el formulario de tarjetas con la nueva tarjeta
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
    
    MostrarFormulario frmTarjetas
    frmTarjetas.Actualizar
    mdifrmPrincipal.ActualizarMenu
    
    Exit Sub
ErrorHandle:
    ReportErr "cMostrarTarjeta", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cActualizarTarjetasJugador(intColor As Integer, intCantTarjetas As Integer)
    'Actualiza la cantidad de tarjetas del jugador especificado
    On Error GoTo ErrorHandle
    
    GvecJugadores(intColor).intCantidadTarjetas = intCantTarjetas
    
    'Por si está seleccionado
    frmJugadores.ActualizarDetalleJugadorSeleccionado
    
    Exit Sub
ErrorHandle:
    ReportErr "cActualizarTarjetasJugadores", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cActualizarCanjesJugador(intColor As Integer, intCantCanjes As Integer)
    'Actualiza la cantidad de canjes del jugador especificado
    On Error GoTo ErrorHandle
    
    GvecJugadores(intColor).intCanje = intCantCanjes
    
    'Por si está seleccionado
    frmJugadores.ActualizarDetalleJugadorSeleccionado
    
    Exit Sub
ErrorHandle:
    ReportErr "cActualizarCanjesJugador", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cCobrarTarjeta()
    On Error GoTo ErrorHandle
    Dim intTarjetaSeleccionada As Integer
    Dim i As Integer
    
    'Busca la tarjeta seleccionada
    For i = 0 To frmTarjetas.shpTarjetaSel.Count - 1
        If frmTarjetas.shpTarjetaSel(i).Visible Then
            intTarjetaSeleccionada = i + 1
            Exit For
        End If
    Next i
    
    'Valida que exist alguna tarjeta seleccionada
    If intTarjetaSeleccionada <= 0 Then
        Exit Sub
    End If
    
    'Valida que la tarjeta seleccionada corresponda a un pais mio
    If frmMapa.objPais(GvecTarjetas(intTarjetaSeleccionada).byPais).Color <> GintMiColor Then
        Exit Sub
    End If
    
    'Valida que la tarjeta seleccionada no haya sido cobrada con anterioridad
    If GvecTarjetas(intTarjetaSeleccionada).blCobrada Then
        Exit Sub
    End If
    
    'Envia el mensaje
    EnviarMensaje ArmarMensajeParam(msgCobroTarjeta, GvecTarjetas(intTarjetaSeleccionada).byPais)
    
    'Deselecciona la tarjeta en cuestion
    frmTarjetas.shpTarjetaSel(intTarjetaSeleccionada - 1).Visible = False
    
    Exit Sub
ErrorHandle:
    ReportErr "cCobrarTarjeta", "mdlCliente", Err.Description, Err.Number, Err.Source
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
                frmTarjetas.Actualizar
                Exit For
            End If
        Next i
    End If
    
    'Actualiza el mapa
    cActualizarPais byPais, intColor, intCantidad, orCobroTarjeta, (intCantidad - frmMapa.objPais(byPais).CantTropas) + frmMapa.objPais(byPais).TropasFijas
    
    'Deselecciona las tarjetas seleccionadas
    frmTarjetas.DeseleccionarTarjetas
    
    Exit Sub
ErrorHandle:
    ReportErr "cCobrarTarjeta", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cCanjearTarjeta()
    On Error GoTo ErrorHandle
    Dim vecTarjetaSeleccionada(0 To 2) As Integer
    Dim i As Integer
    Dim intCantTarjetasSeleccionadas
    Dim blnCanjeValido As Boolean
    
    intCantTarjetasSeleccionadas = 0
    blnCanjeValido = False
    
    'Busca las tarjetas seleccionadas
    For i = 0 To frmTarjetas.shpTarjetaSel.Count - 1
        If frmTarjetas.shpTarjetaSel(i).Visible = True Then
            intCantTarjetasSeleccionadas = intCantTarjetasSeleccionadas + 1
            vecTarjetaSeleccionada(intCantTarjetasSeleccionadas - 1) = i + 1
        End If
    Next i
    
    If intCantTarjetasSeleccionadas = 3 Then
        'Busca si hay algun comodin
        If GvecTarjetas(vecTarjetaSeleccionada(0)).byFigura = figComodin _
        Or GvecTarjetas(vecTarjetaSeleccionada(1)).byFigura = figComodin _
        Or GvecTarjetas(vecTarjetaSeleccionada(2)).byFigura = figComodin Then
            'Si hay algun comodin el canje es válido
            blnCanjeValido = True
        Else
            'Si no hay comodin se fija si son todas iguales o todas distintas
            If GvecTarjetas(vecTarjetaSeleccionada(0)).byFigura = GvecTarjetas(vecTarjetaSeleccionada(1)).byFigura _
            And GvecTarjetas(vecTarjetaSeleccionada(1)).byFigura = GvecTarjetas(vecTarjetaSeleccionada(2)).byFigura Then
                'Son las tres iguales
                blnCanjeValido = True
            ElseIf GvecTarjetas(vecTarjetaSeleccionada(0)).byFigura <> GvecTarjetas(vecTarjetaSeleccionada(1)).byFigura _
               And GvecTarjetas(vecTarjetaSeleccionada(1)).byFigura <> GvecTarjetas(vecTarjetaSeleccionada(2)).byFigura _
               And GvecTarjetas(vecTarjetaSeleccionada(0)).byFigura <> GvecTarjetas(vecTarjetaSeleccionada(2)).byFigura Then
                'Son las tres distintas
                blnCanjeValido = True
            End If
        End If
    End If
    
    'Si está todo bien envia el mensaje
    If blnCanjeValido Then
        EnviarMensaje ArmarMensajeParam(msgCanjeTarjeta, GvecTarjetas(vecTarjetaSeleccionada(0)).byPais, _
                                                         GvecTarjetas(vecTarjetaSeleccionada(1)).byPais, _
                                                         GvecTarjetas(vecTarjetaSeleccionada(2)).byPais)
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cCanjearTarjeta", "mdlCliente", Err.Description, Err.Number, Err.Source
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
    frmTarjetas.Actualizar
    
    'Deselecciona las tarjetas seleccionadas
    frmTarjetas.DeseleccionarTarjetas
    
    Exit Sub
ErrorHandle:
    ReportErr "cAckCanjearTarjeta", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMisionCumplida(intColorGanador As Integer, strMisionGanador As String, intMisionId As Integer)
    On Error GoTo ErrorHandle
    Dim strMensaje As String
    
    '###E
    'Actualiza el estado del cliente
    ActualizarEstadoCliente eveMisionCumplida
    
    'Desactiva el timer
    mdifrmPrincipal.Timer1.Interval = 0
    
    'Pepito ha logrado cumplir su misión:
    strMensaje = CompilarMensaje(ObtenerTextoRecurso(CintGralMsgMisionCumplida), Array(GvecJugadores(intColorGanador).strNombre))
    frmMisionCumplida.lblGanador = strMensaje
    
    strMensaje = ObtenerTextoRecurso(intMisionId)
    If strMensaje = "" Then
        strMensaje = strMisionGanador
    End If
    frmMisionCumplida.lblMision = strMensaje
    
    MostrarFormulario frmMisionCumplida, vbModal
    
    Exit Sub
ErrorHandle:
    ReportErr "cMisionCumplida", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cActualizarEstadoTurno(intEstadoTurno As enuEstadoCli, strEstadoClientePausado As String)
    'Actualiza el estado del cliente (forzado por la resincronizacion)
    On Error GoTo ErrorHandle
    
    GEstadoCliente = intEstadoTurno
    
    If strEstadoClientePausado = CStr(estPartidaPausada) Then
        'Si la Partida está pausada, deja al Cliente Pausado
        mdifrmPrincipal.mnuPartidaPausar_Click
    Else
        'Actualiza el Cliente, segun su Estado.
        mdifrmPrincipal.StatusBar1.Panels(enuPaneles.panEstado).Text = GvecEstadoCliente(GEstadoCliente)
    
        '### Matriz de controles (habilitar/deshabilitar opciones)
        ActualizarControles
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cActualizarEstadoTurno", "mdlCliente", Err.Description, Err.Number, Err.Source
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
    ReportErr "cResincronizar", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cReconectarJugador(intColor As Integer)
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgReconexion, intColor)
    
    Exit Sub
ErrorHandle:
    ReportErr "cReconectarJugador", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMostrarPartidasGuardadas(vecPartidas() As String)
    On Error GoTo ErrorHandle
    'Muestra el formulario de Guardar Partida
    Dim i As Integer
    Dim datFecha As Date
    Dim strNombre As String
    
    'Limpia el List
    frmGuardarPartida.lstPartidasGuardadas.Clear
    ReDim GvecNombresPartidas(UBound(vecPartidas))
    
    For i = LBound(vecPartidas) To UBound(vecPartidas)
        'Si el elemento está vacío
        If Trim(vecPartidas(i)) <> "" Then
            strNombre = Mid$(vecPartidas(i), 11)
            datFecha = CDate(Left(vecPartidas(i), 10))
            If strNombre <> strNombrePartidaActual Then
                frmGuardarPartida.lstPartidasGuardadas.AddItem CStr(datFecha) & " - " & strNombre
                'Guarda los nombres en un vector paralelo, para recuperarlo facilmente
                GvecNombresPartidas(frmGuardarPartida.lstPartidasGuardadas.ListCount - 1) = strNombre
            End If
        End If
    Next i
    
    If Not frmGuardarPartida.Visible Then
        MostrarFormulario frmGuardarPartida, vbModal
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cMostrarPartidasGuardadas", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cPedirPartidasGuardadas()
    On Error GoTo ErrorHandle
    'Le pide al servidor los nombres de las partidas guardadas,
    'que serán mostrados en el formulario de Guardar Partida
    
    EnviarMensaje ArmarMensajeParam(msgPedidoPartidasGuardadas)
    
    Exit Sub
ErrorHandle:
    ReportErr "cPedirPartidasGuardadas", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cEliminarPartida(strNombrePartida As String)
    On Error GoTo ErrorHandle
    'Envia al servidor el nombre de la partida que se desea eliminar
    
    EnviarMensaje ArmarMensajeParam(msgEliminarPartida, strNombrePartida)
    
    Exit Sub
ErrorHandle:
    ReportErr "cEliminarPartida", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cGuardarPartida(strNombrePartida As String)
    On Error GoTo ErrorHandle
    'Envia al servidor el nombre de la partida que se desea guardar
    
    EnviarMensaje ArmarMensajeParam(msgGuardarPartida, strNombrePartida)
    
    Exit Sub
ErrorHandle:
    ReportErr "cGuardarPartida", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cAckGuardarPartida(intAck As enuAckGuardarPartida)
    On Error GoTo ErrorHandle
    
    If intAck = enuAckGuardarPartida.ackOk Then
        Unload frmGuardarPartida
    Else
        Select Case intAck
            Case enuAckGuardarPartida.ackNoAdministrador
                '"Imposible guardar partida. Usted no es Administrador."
                MsgBox ObtenerTextoRecurso(CintGralMsgErrGuardarPartidaNoAdm), vbExclamation, ObtenerTextoRecurso(CintGralMsgErrGuardarPartidaNoAdmCaption)
            Case enuAckGuardarPartida.ackErrorDesconocido
                '"Imposible guardar partida. Error desconocido."
                MsgBox ObtenerTextoRecurso(CintGralMsgErrGuardarPartidaDesconocido), vbExclamation, ObtenerTextoRecurso(CintGralMsgErrGuardarPartidaDesconocidoCaption)
        End Select
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cAckGuardarPartida", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub ReportErr(ByVal strFuncion As String, ByVal strModulo As String, ByVal strDesc As String, _
                    ByVal intErr As Long, ByVal strSource As String, _
                    Optional styIcono As VbMsgBoxStyle = vbCritical)
    'Reportes de Errores
    On Error GoTo ErrorHandle
    Dim strMsg As String
    If intErr <> 0 Then
        strMsg = ObtenerTextoRecurso(CintGralMsgErrRutina) & strFuncion & Chr(10) & _
                 ObtenerTextoRecurso(CintGralMsgErrModulo) & strModulo & vbCrLf & _
                 ObtenerTextoRecurso(CintGralMsgErrDescripcion) & strDesc & Space(10) & vbCrLf & _
                 ObtenerTextoRecurso(CintGralMsgErrOrigen) & strSource
        MsgBox strMsg, styIcono, CompilarMensaje(ObtenerTextoRecurso(CintGralMsgErrNumero), Array(intErr))
    End If
    
    Err.Clear

    Exit Sub

ErrorHandle:
    MsgBox ObtenerTextoRecurso(CintGralMsgErrRutinaError) & Chr(10) & ObtenerTextoRecurso(CintGralMsgErrDescripcion) & Err.Description, vbCritical, CompilarMensaje(ObtenerTextoRecurso(CintGralMsgErrNumero), Array(Err.Number))
    Close

End Sub

Public Sub MostrarFormulario(frmform As Form, Optional intModal As FormShowConstants = vbModeless)
    'Intenta hacer un Show del formulario especificado
    'La idea es evitar que se muestre un formulario no modal cuando
    'hay un modal abierto
    On Error Resume Next
    
    If intModal = vbModal Then
        frmform.Show intModal
    Else
        'Si no es modal, no hace el show para evitar que el
        'formulario principal tome el foco.
        frmform.Visible = True
        frmform.ZOrder 0
    End If

End Sub

Public Sub cRecibirLog(vecLog() As String)
    'Recibe del servidor un mensaje de log compuesto por un código de máscara y
    'sus parámetros.
    'En base a esto arma el mensaje final tomando los valores del archivo de recursos
    On Error GoTo ErrorHandle
    Dim intCodMascara As Integer
    Dim strMascara As String
    Dim strMensajeLog As String
    Dim vecParametros() As Variant
    Dim intParametro As Integer
    
    'Obtiene la máscara
    intCodMascara = vecLog(LBound(vecLog))
    strMascara = ObtenerTextoRecurso(intCodMascara)
    
    'Obtiene los parámetros
    If UBound(vecLog) > 0 Then
        ReDim vecParametros(UBound(vecLog) - 1)
        For intParametro = 1 To UBound(vecLog)
            If Left(vecLog(intParametro), 1) = CstrTipoParametroRecurso Then
                vecParametros(intParametro - 1) = ObtenerTextoRecurso(Mid(vecLog(intParametro), 2))
            Else
                vecParametros(intParametro - 1) = Mid(vecLog(intParametro), 2)
            End If
        Next
    End If
    
    'Resuelve el mensaje
    strMensajeLog = CompilarMensaje(strMascara, vecParametros)
    
    MostrarLog strMensajeLog
    
    Exit Sub
ErrorHandle:
    ReportErr "cRecibirLog", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub MostrarLog(strMensaje As String)
    On Error GoTo ErrorHandle
    
    frmLog.lstLog.AddItem Format(Time, "hh:mm:ss") & " - " & strMensaje
    'Se posiciona en el último elemento
    frmLog.lstLog.ListIndex = frmLog.lstLog.ListCount - 1
    frmLog.lstLog.ListIndex = -1
        
    'Agrega el mensaje de log al chat si la opción está habilitada
    If frmChat.chkLog.Value = 1 Then
        'Posiciona el punto de inserción al final de la cadena
        frmChat.txtRecibido.SelStart = Len(frmChat.txtRecibido.Text)
        frmChat.txtRecibido.SelLength = 1
        
        If frmChat.txtRecibido.Text <> "" Then
            frmChat.txtRecibido.SelText = vbCrLf
        End If
        
        'Color del remitente
        frmChat.txtRecibido.SelColor = RGB(120, 120, 120)
        frmChat.txtRecibido.SelBold = True
    
        frmChat.txtRecibido.SelText = strMensaje
        
        'Posiciona el punto de inserción al final de la cadena
        frmChat.txtRecibido.SelStart = Len(frmChat.txtRecibido.Text)
        frmChat.txtRecibido.SelLength = 1
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "MostrarLog", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMostrarIpServidor(strIpServidor As String)
    On Error GoTo ErrorHandle
    
    GstrIpServidor = strIpServidor
    
    frmSeleccionColor.StatusBar1.SimpleText = frmSeleccionColor.StatusBar1.SimpleText & " " & strIpServidor & "]"
        
    Exit Sub
ErrorHandle:
    ReportErr "MostrarLog", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub GrabarSeteo(strNombre As String, strValor As String)
    'Graba en la registry el valor de la clave pasada por parametro
    On Error Resume Next
    SaveSetting "TEGNet", "Personalizacion", strNombre, strValor
End Sub

Public Sub GrabarPersonalizacion()
    'Graba en la registry los valores de la personalización
    On Error GoTo ErrorHandle
    
    GrabarSeteo "mdifrmPrincipalWindowState", mdifrmPrincipal.WindowState
    If mdifrmPrincipal.WindowState = vbNormal Then
        GrabarSeteo "mdifrmPrincipalLeft", mdifrmPrincipal.Left
        GrabarSeteo "mdifrmPrincipalWidth", mdifrmPrincipal.Width
        GrabarSeteo "mdifrmPrincipalHeight", mdifrmPrincipal.Height
        GrabarSeteo "mdifrmPrincipalTop", mdifrmPrincipal.Top
    End If
        
'    GrabarSeteo "frmChatWindowState", frmChat.WindowState
'    If frmChat.WindowState = vbNormal Then
'        GrabarSeteo "frmChatTop", frmChat.Top
'        GrabarSeteo "frmChatLeft", frmChat.Left
'        GrabarSeteo "frmChatWidth", frmChat.Width
'        GrabarSeteo "frmChatHeight", frmChat.Height
'    End If
'    GrabarSeteo "frmChatChkLog", frmChat.chkLog.Value
'    GrabarSeteo "frmChatVisible", CInt(frmChat.Visible)
    GrabarSeteo "frmChatTop", frmChat.Top
    GrabarSeteo "frmChatLeft", frmChat.Left
    GrabarSeteo "frmChatWidth", frmChat.Width
    GrabarSeteo "frmChatHeight", frmChat.Height
    GrabarSeteo "frmChatChkLog", frmChat.chkLog.Value
    GrabarSeteo "frmChatVisible", CInt(frmChat.Visible)
    
    GrabarSeteo "frmMapaTop", frmMapa.Top
    GrabarSeteo "frmMapaLeft", frmMapa.Left
    GrabarSeteo "frmMapaVisible", CInt(frmMapa.Visible)
    
    GrabarSeteo "frmJugadoresTop", frmJugadores.Top
    GrabarSeteo "frmJugadoresLeft", frmJugadores.Left
    GrabarSeteo "frmJugadoresVisible", CInt(frmJugadores.Visible)
    
    GrabarSeteo "frmSeleccionTop", frmSeleccion.Top
    GrabarSeteo "frmSeleccionLeft", frmSeleccion.Left
    GrabarSeteo "frmSeleccionWidth", frmSeleccion.Width
    GrabarSeteo "frmSeleccionHeight", frmSeleccion.Height
    GrabarSeteo "frmSeleccionVisible", CInt(frmSeleccion.Visible)
    
    GrabarSeteo "frmDadosTop", frmDados.Top
    GrabarSeteo "frmDadosLeft", frmDados.Left
    GrabarSeteo "frmDadosVisible", CInt(frmDados.Visible)
    
    GrabarSeteo "frmMisionTop", frmMision.Top
    GrabarSeteo "frmMisionLeft", frmMision.Left
    GrabarSeteo "frmMisionVisible", CInt(frmMision.Visible)
    
    GrabarSeteo "frmTarjetasTop", frmTarjetas.Top
    GrabarSeteo "frmTarjetasLeft", frmTarjetas.Left
    GrabarSeteo "frmTarjetasVisible", CInt(frmTarjetas.Visible)
    
    GrabarSeteo "frmTropasDisponiblesTop", frmTropasDisponibles.Top
    GrabarSeteo "frmTropasDisponiblesLeft", frmTropasDisponibles.Left
    GrabarSeteo "frmTropasDisponiblesVisible", CInt(frmTropasDisponibles.Visible)
    
    GrabarSeteo "frmPropiedadesTop", frmPropiedades.Top
    GrabarSeteo "frmPropiedadesLeft", frmPropiedades.Left
    GrabarSeteo "frmPropiedadesVisible", CInt(frmPropiedades.Visible)
    
    GrabarSeteo "frmMisionesTop", frmMisiones.Top
    GrabarSeteo "frmMisionesLeft", frmMisiones.Left
    GrabarSeteo "frmMisionesVisible", CInt(frmMisiones.Visible)
    
    GrabarSeteo "frmLogTop", frmLog.Top
    GrabarSeteo "frmLogLeft", frmLog.Left
    GrabarSeteo "frmLogWidth", frmLog.Width
    GrabarSeteo "frmLogHeight", frmLog.Height
    GrabarSeteo "frmLogVisible", CInt(frmLog.Visible)
    
    GrabarSeteo "Version", App.Major * 10000 + App.Minor * 100 + App.Revision
    
    Exit Sub
ErrorHandle:
    ReportErr "GrabarPersonalizacion", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Function CargarSeteo(strNombre As String, strValorDefecto As String) As String
    'Obtiene de la registry el valor de la clave pasada por parametro
    On Error Resume Next
    CargarSeteo = GetSetting("TEGNet", "Personalizacion", strNombre, strValorDefecto)
End Function

Public Sub CargarPersonalizacion()
    'Carga de la registry los valores de la personalización
    On Error GoTo ErrorHandle
    Dim blnOrganizarVentanas As Boolean
    
    'Si la version que figura en la Registry es distinta (o no existe),
    'al final ejecuta la funcion Organizar Ventanas.
    If CargarSeteo("Version", 0) = (App.Major * 10000 + App.Minor * 100 + App.Revision) Then
        blnOrganizarVentanas = False
    Else
        blnOrganizarVentanas = True
    End If
    
    With mdifrmPrincipal
        .WindowState = CargarSeteo("mdifrmPrincipalWindowState", vbMaximized)
        If .WindowState = vbNormal Then
            .Width = CargarSeteo("mdifrmPrincipalWidth", .Width)
            .Height = CargarSeteo("mdifrmPrincipalHeight", .Height)
            .Left = CargarSeteo("mdifrmPrincipalLeft", .Left)
            .Top = CargarSeteo("mdifrmPrincipalTop", .Top)
        End If
    End With
    
    'Cargar este form, antes que el Seleccion
    frmJugadores.Top = CargarSeteo("frmJugadoresTop", 0)
    frmJugadores.Left = CargarSeteo("frmJugadoresLeft", frmMapa.Width + 2)
    frmJugadores.Visible = CargarSeteo("frmJugadoresVisible", "-1")
    
    'Cargar este form, antes que el Chat y el Log
    frmSeleccion.Left = CargarSeteo("frmSeleccionLeft", frmMapa.Width)
    frmSeleccion.Top = CargarSeteo("frmSeleccionTop", frmJugadores.Height)
    frmSeleccion.Width = CargarSeteo("frmSeleccionWidth", frmJugadores.Width)
    frmSeleccion.Height = CargarSeteo("frmSeleccionHeight", 930)
    frmSeleccion.Visible = CargarSeteo("frmSeleccionVisible", "-1")
    
    With frmChat
        .WindowState = CargarSeteo("frmChatWindowState", vbNormal)
        If .WindowState = vbNormal Then
            .Left = CargarSeteo("frmChatLeft", 0)
            .Top = CargarSeteo("frmChatTop", frmMapa.Height)
            .Width = CargarSeteo("frmChatWidth", frmMapa.Width)
            .Height = CargarSeteo("frmChatHeight", frmSeleccion.Height)
        End If
        .chkLog = CargarSeteo("frmChatChkLog", 0)
        .Visible = CargarSeteo("frmChatVisible", "-1")
    End With
    
    frmMapa.Top = CargarSeteo("frmMapaTop", 0)
    frmMapa.Left = CargarSeteo("frmMapaLeft", 0)
    frmMapa.Visible = CargarSeteo("frmMapaVisible", "-1")
    
    frmDados.Left = CargarSeteo("frmDadosLeft", 45)
    'frmDados.Top = CargarSeteo("frmDadosTop", 3720)
    frmDados.Top = CargarSeteo("frmDadosTop", frmMapa.Height - frmDados.Height - 45)
    frmDados.Visible = CargarSeteo("frmDadosVisible", "0")
    
    frmMision.Top = CargarSeteo("frmMisionTop", frmMision.Top)
    frmMision.Left = CargarSeteo("frmMisionLeft", frmMision.Left)
    frmMision.Visible = CargarSeteo("frmMisionVisible", "0")
    
    frmTarjetas.Top = CargarSeteo("frmTarjetasTop", frmTarjetas.Top)
    frmTarjetas.Left = CargarSeteo("frmTarjetasLeft", frmTarjetas.Left)
    frmTarjetas.Visible = CargarSeteo("frmTarjetasVisible", "0")

    frmTropasDisponibles.Left = CargarSeteo("frmTropasDisponiblesLeft", 45)
    'frmTropasDisponibles.Top = CargarSeteo("frmTropasDisponiblesTop", 3255)
    frmTropasDisponibles.Top = CargarSeteo("frmTropasDisponiblesTop", frmMapa.Height - frmTropasDisponibles.Height - 45)
    frmTropasDisponibles.Visible = CargarSeteo("frmTropasDisponiblesVisible", "0")

    frmPropiedades.Top = CargarSeteo("frmPropiedadesTop", frmPropiedades.Top)
    frmPropiedades.Left = CargarSeteo("frmPropiedadesLeft", frmPropiedades.Left)
    frmPropiedades.Visible = CargarSeteo("frmPropiedadesVisible", "0")
    
    frmMisiones.Top = CargarSeteo("frmMisionesTop", frmMisiones.Top)
    frmMisiones.Left = CargarSeteo("frmMisionesLeft", frmMisiones.Left)
    frmMisiones.Visible = CBool(CargarSeteo("frmMisionesVisible", "0"))
    
    frmLog.WindowState = CargarSeteo("frmLogWindowState", vbNormal)
    If frmLog.WindowState = vbNormal Then
        frmLog.Width = CargarSeteo("frmLogWidth", frmMapa.Width)
        frmLog.Height = CargarSeteo("frmLogHeight", frmSeleccion.Height)
        frmLog.Left = CargarSeteo("frmLogLeft", 0)
        frmLog.Top = CargarSeteo("frmLogTop", frmMapa.Height)
    End If
    frmLog.Visible = CargarSeteo("frmLogVisible", "0")
    
    'Por ultimo...
    If blnOrganizarVentanas = True Then
        'Organizar Ventanas
        mdifrmPrincipal.mnuOrganizar_Click
    End If

    Exit Sub
ErrorHandle:
    ReportErr "CargarPersonalizacion", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cRecibirVersionServidor(intMajor As Integer, intMinor As Integer, intRevision As Integer)
    'Valida que la versión del servidor sea compatible con la versión del cliente
    On Error GoTo ErrorHandle
    
    GlngVersionServidor = intMajor * 10000 + intMinor * 100 + intRevision
    
    cEnviarVersionCliente
    
    Exit Sub
ErrorHandle:
    ReportErr "cRecibirVersionServidor", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cEnviarVersionCliente()
    'Informa al servidor de la version y tipo del cliente
    On Error GoTo ErrorHandle
    
    'Tipo de Cliente:
    EnviarMensaje ArmarMensajeParam(msgVersionCliente, enuInteligenciaJugador.hrHumano, App.Major, App.Minor, App.Revision)
    
    Exit Sub
ErrorHandle:
    ReportErr "cEnviarVersionCliente", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cMoverTodas(byPaisDesde As Byte, byPaisHasta As Byte)
    'Mueve todas las tropas posibles de un país a otro
    On Error GoTo ErrorHandle
    
    Dim intTropasMover As Integer
    
    'Calcula la cantidad de tropas que puede mover
    intTropasMover = frmMapa.objPais(byPaisDesde).CantTropas - frmMapa.objPais(byPaisDesde).TropasFijas - 1
    
    If intTropasMover > 0 Then
        cMover byPaisDesde, byPaisHasta, intTropasMover, tmMovimiento
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cEnviarVersionCliente", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Private Sub cRecibirLimitrofes(vecLimitrofes() As String)
    On Error GoTo ErrorHandle
    
    frmMapa.CargarLimitrofesServidor vecLimitrofes
    
    Exit Sub
ErrorHandle:
    ReportErr "cRecibirLimitrofes", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub EvaluarHabilitacionPausa()
    'Si el jugador es ADM, estoy Jugando y hay solo Robots, habilita la opcion de Pausa,
    'en caso contrario la deshabilita.
    On Error GoTo ErrorHandle
    
    Dim blnMostrar As Boolean
    Dim i As Integer
    
    blnMostrar = False
    If GsoyAdministrador And GEstadoCliente >= estEsperandoTurno And GEstadoCliente <> estInconsistente Then
    
        'Se fija si solo hay Robots
        blnMostrar = True
        For i = LBound(GvecJugadores) To UBound(GvecJugadores)
            If i <> GintMiColor And GvecJugadores(i).strNombre <> "" And GvecJugadores(i).byTipoJugador <> enuInteligenciaJugador.hrRobot Then
                blnMostrar = False
            End If
        Next i
        
    End If
    
    'Habilita/Deshabilita la Opcion
    mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbPausa).Visible = blnMostrar
    mdifrmPrincipal.mnuPartidaPausar.Visible = blnMostrar
    
    Exit Sub
ErrorHandle:
    ReportErr "EvaluarHabilitacionPausa", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cPausarPartida()
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgPausarPartida)
    
    Exit Sub
ErrorHandle:
    ReportErr "cPausarPartida", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cAckPausarPartida()
    On Error GoTo ErrorHandle
    
    'Guarda el estado Anterior (al cual luego deberá volver)
    CargarRegistroEstado estPartidaPausada, evePartidaContinuada, GEstadoCliente
    '###E
    ActualizarEstadoCliente evePartidaPausada
    
    'Actualiza el Caption de la Opcion y el Boton
    mdifrmPrincipal.mnuPartidaPausar.Caption = ObtenerTextoRecurso(CintPrincipalTipContinuar)
    mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbPausa).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipContinuar)
    mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbPausa).Image = 15
    
    'Detiene el Timer
    mdifrmPrincipal.Timer1.Interval = 0
    GsngInicioPausaTimerTurno = Timer
    
    'Actualiza los iconos del mouse del mapa
    '(no muestra ningun puntero "prohibido")
    frmMapa.ActualizarIconosMouse
    
    'Oculta los paises seleccionados
    frmMapa.DeseleccionarDestino
    frmMapa.DeseleccionarOrigen
    'Oculta todos los paises del mapa
    frmMapa.LimpiarMapa
    
    Exit Sub
ErrorHandle:
    ReportErr "cAckPausarPartida", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cContinuarPartida()
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgContinuarPartida)
    
    Exit Sub
ErrorHandle:
    ReportErr "cContinuarPartida", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

Public Sub cAckContinuarPartida()
    On Error GoTo ErrorHandle
    
    '###E
    ActualizarEstadoCliente evePartidaContinuada
    
    'Actualiza el Caption de la Opcion y el Boton
    mdifrmPrincipal.mnuPartidaPausar.Caption = ObtenerTextoRecurso(CintPrincipalTipPausar)
    mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbPausa).ToolTipText = ObtenerTextoRecurso(CintPrincipalTipPausar)
    mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbPausa).Image = 14

    'Reactiva el timer
    GsngInicioTimerTurno = GsngInicioTimerTurno + (Timer - GsngInicioPausaTimerTurno)
    mdifrmPrincipal.Timer1.Interval = 1000
    
    'Actualiza los iconos del mouse del mapa
    '(segun el estado muestra el puntero "prohibido" donde corresponda)
    frmMapa.ActualizarIconosMouse
    
    Exit Sub
ErrorHandle:
    ReportErr "cAckPausarPartida", "mdlCliente", Err.Description, Err.Number, Err.Source
End Sub

