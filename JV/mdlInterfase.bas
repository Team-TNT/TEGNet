Attribute VB_Name = "mdlInterfase"
Option Explicit

Public Const chrSEPARADOR = "|"
Public Const chrCOMIENZO = "#"
Public Const chrPREFIJOADM = "@"
Public Const chrSUFIJOADM = ""
Public Const chrMASCARAMENSAJE = "%"
Public Const strLOCALHOST = "localhost"
Public Const strNombrePartidaActual = "@ACTUAL@"
Public Const intCANTMAXJUGADORES = 6
Public Const CintValorTimerInfinito = 9999

Public Const CstrTipoParametroRecurso = "R"
Public Const CstrTipoParametroResuelto = "X"

Public Enum enuTipoMensaje
    'Conexion
    msgConfirmarAdm = 101
    msgTipoPartida = 102
    msgPartidasGuardadas = 103
    msgNombrePartida = 104
    msgJugadoresConectados = 105
    msgAltaJugador = 106
    msgAckAltaJugador = 107
    msgPedirConexionesActuales = 108
    msgIpServidor = 109
    msgVersionServidor = 110
    msgVersionCliente = 111
    'Desconexion
    msgBajarServidor = 201
    msgAYA = 202
    msgIAA = 203
    'Mantenimiento
    msgPais = 301
    msgMisionAsignada = 302
    msgOpciones = 303
    msgOpcionesDefault = 304
    msgCambiarAdm = 305
    msgBajaAdm = 306
    msgEstadoTurnoCliente = 307
    msgResincronizacion = 308
    msgReconexion = 309
    msgPedidoPartidasGuardadas = 310
    msgEliminarPartida = 311
    msgGuardarPartida = 312
    msgAckGuardarPartida = 313
    msgLog = 314
    msgLimitrofes = 315
    msgPaisContinente = 316
    msgPausarPartida = 317
    msgAckPausarPartida = 318
    msgContinuarPartida = 319
    msgAckContinuarPartida = 320
    msgPartidaPausada = 321
    msgPartidaContinuada = 322
    msgError = 399
    'Juego
    msgIniciarPartida = 401
    msgOrdenRonda = 402
    'msgComienzoTurno = 403     Es redundante con msgInicioTurno
    msgAckInicioPartida = 404
    msgInicioTurno = 405
    msgTropasDisponibles = 406
    msgFinTurno = 407
    msgAckFinTurno = 408
    msgTipoRonda = 409
    msgAgregarTropas = 410
    msgAtaque = 411
    msgAckAtaque = 412
    msgMovimiento = 413
    msgAckMovimiento = 414
    msgPedidoTarjeta = 415
    msgTarjetasJugador = 416
    msgTarjeta = 417
    msgCobroTarjeta = 418
    msgAckCobroTarjeta = 419
    msgCanjeTarjeta = 420
    msgAckCanjeTarjeta = 421
    msgMisionCumplida = 422
    msgAckAgregarTropas = 423 'Unicamente para el Jugador Virtual
    msgCanjesJugador = 424
    'Chat
    msgChatEntrante = 501
    msgChatSaliente = 502
End Enum

Public Enum enuTipoPartida
    tpNueva = 1
    tpGuardada = 2
End Enum

'Enumeración con las opciones
Public Enum enuOpciones
    opTurnoDuracion = 1
    opTurnoTolerancia = 2
    opRondaTropas1ra = 10
    opRondaTropas2da = 11
    opRondaTipo = 12
    opBonusTarjetaPropia = 20
    opBonusAfrica = 21
    opBonusANorte = 22
    opBonusASur = 23
    opBonusAsia = 24
    opBonusEuropa = 25
    opBonusOceania = 26
    opMisionDestruir = 30
    opMisionTipo = 31 'por misiones o a conquistar el mundo
    opMisionObjetivoComun = 32
    opTropasInicial = 40
    opCanjeNro1 = 50
    opCanjeNro2 = 51
    opCanjeNro3 = 52
    opCanjeIncremento = 53
End Enum

'------------------Enumeraciones para ACK-------------------------
Public Enum enuAckAltaJugador
    ackOk = 0
    ackColorUsado = -1
    ackNombreUsado = -2
    ackNombreYColorUsados = -3
    'Para reconexion
    ackColorInexistente = -4
    ackColorConectado = -5
    ackServidorPausado = -6
End Enum
'-----------------------------------------------------------------

'Enumeración de tipos de ronda
Public Enum enuTipoRonda
    trInicio
    trAccion
    trRecuento
End Enum

'Enumeracion de Origenes para el mensaje pais (para efecto visual)
Public Enum enuOrigenMsgPais
    orRepartoInicial
    orAgregado
    orAtaque
    orMovimiento
    orCobroTarjeta
End Enum

'Enumeración con los tipos de movimiento
Public Enum enuTipoMovimiento
    tmConquista
    tmMovimiento
End Enum

'Enumeración con las figuras de las tarjetas
Public Enum enuFigurasTarjetas
    figGlobo = 1
    figCanon = 2
    figBarco = 3
    figComodin = 4
End Enum

'Estados del cliente
Public Enum enuEstadoCli
    estDesconectado
    estConectado
    estValidado
    estEsperandoTurno
    estAgregando
    estAtacando
    estMoviendo
    estTarjetaCobrada
    estTarjetaTomada
    estTarjetaCobradaTomada
    estPartidaPausada
    estPartidaFinalizada
    estInconsistente
End Enum

'Flag Tipo del mensaje Jugadores conectados
Public Enum enuTipoJugadoresConectados
    tjcIngreso = 1
    tjcReconexion = 2
End Enum

'Estado de la conexión de los clientes
Public Enum enuEstadoConexion
    conNoJuega
    conConectado
    conDesconectado
End Enum

'Tipo de mensaje de partida guardada
Public Enum enuTipoPartidaGuardada
    tpgIniciarPartida
    tpgGuardarPartida
End Enum

'ACK de Guardar Partida
Public Enum enuAckGuardarPartida
    ackOk = 0
    ackNoAdministrador = -1
    ackErrorDesconocido = -99
End Enum

Public Enum enuErrores
    errNoAccionNoAdm = 1
    errNoAltaPartidaIniciada
    errNoFinTurnoNoTurno
    errNoAgregarNoTurno
    errNoAgregarRondaAccion
    errNoAgregarNoPais
    errNoAgregarNoTropas
    errNoAgregarNoCantidad
    errNoAtaqueNoTurno
    errNoAtaqueNoRondaAccion
    errNoAtaqueNoOrigen
    errNoAtaqueNoDestino
    errNoAtaqueNoTropas
    errNoAtaqueNoLimitrofes
    errNoMovimientoNoTurno
    errNoMovimientoNoRondaAccion
    errNoMovimientoNoOrigen
    errNoMovimientoNoDestino
    errNoMovimientoNoTropas
    errNoMovimientoNoLimitrofes
    errNoTomarNoTurno
    errNoTomarNoRondaAccion
    errNoTomarNoConquista1
    errNoTomarNoConquista2
    errNoTomarNoMasTarjetas
    errNoCobrarNoTurno
    errNoCobrarNoRondaAccion
    errNoCobrarNoPais
    errNoCobrarYaCobrada
    errNoAccionNoPartida
    errNoAccionNoEstado
    errNoCanjeNoTurno
    errNoCanjeNoTarjetas
    errNoCanjeNoFiguras
    errNoPausaHayHumanos
    errNoContinuarNoPausa
End Enum

'Enumeración Si/No
Public Enum enuSiNo
    snSi = 1
    snNo = 2
End Enum

'Enumeración Humano/Robot
Public Enum enuInteligenciaJugador
    hrHumano = 1
    hrRobot = 2
End Enum

'Enumeración que contiene el inicio de cada tipo de parámetro.
Public Enum enuIndiceArchivoRecurso
    pmsColores = 0
    pmsContinentes = 10
    pmsPaises = 20
    pmsMisiones = 100
    pmsFiguras = 120
    pmsErrores = 130
    pmsOpciones = 200
    pmsSiNo = 260
    pmsInteligenciaJugador = 270
End Enum

'Tipo de Partida (nueva o guardada) elegida por el cliente
Public TipoPartida As enuTipoPartida

Public Function ArmarMensajeParam(TipoMensaje As enuTipoMensaje, ParamArray vecValores()) As String
    'Idem ArmarMensaje con argumentos variables
    '(separados por coma, sin utilizar vectores)
    
    Dim vectorParametros() As String
    Dim i As Integer
    'Si el ParamArray viene vacío
    If UBound(vecValores) > -1 Then
        ReDim vectorParametros(UBound(vecValores))
    Else
        'Creo un vector con un solo elemento (como vector vacío)
        ReDim vectorParametros(0)
    End If
    
    For i = 0 To UBound(vecValores)
        vectorParametros(i) = vecValores(i)
    Next
    
    ArmarMensajeParam = ArmarMensaje(TipoMensaje, vectorParametros)
End Function

Public Function ArmarMensaje(TipoMensaje As enuTipoMensaje, vecValores() As String) As String
    Dim strMensaje As String
    Dim i As Integer
    
    strMensaje = chrCOMIENZO & CStr(TipoMensaje) & chrSEPARADOR
    For i = 0 To UBound(vecValores)
        strMensaje = strMensaje & CStr(vecValores(i)) & chrSEPARADOR
    Next
    ArmarMensaje = strMensaje
End Function

Public Sub InterpretarMensaje(strMensaje As String, Optional intIndiceOrigenMensaje As Integer)
    Dim chrActual As String
    Dim intTipoMensaje As enuTipoMensaje
    Dim strParametro As String
    Dim vecParametros() As String
    Dim i As Integer
    
    i = 1
    strParametro = ""

    '###Debug de los Mensajes
    '(esto es algo exclusivo del servidor)
'    If MODO_DEBUG = "1" Then
'        Open App.Path & "\Mensajes.txt" For Append As #1
'        Print #1, Tiempo & " - IN De: "; Mid$(GvecColores(IndiceAColor(intIndiceOrigenMensaje)) & Space(9), 1, 9) & " - #" & strMensaje
'        Close #1
'    End If

    'Obtiene el tipo de mensaje
    chrActual = Mid$(strMensaje, i, 1)
    While i < Len(strMensaje) And chrActual <> chrSEPARADOR
        strParametro = strParametro & chrActual
        i = i + 1
        chrActual = Mid$(strMensaje, i, 1)
    Wend
    
    intTipoMensaje = CInt(strParametro)
    
    'Salteo el primer separador
    i = i + 1
    
    'Obtiene el resto de los parametros
    vecParametros = Split(Mid$(strMensaje, i, Len(strMensaje) - i), chrSEPARADOR)

    DistribuirMensaje intTipoMensaje, vecParametros, intIndiceOrigenMensaje
    
End Sub

Public Sub SepararMensajes(strMensajes As String, Optional intIndiceOrigenMensaje As Integer)
    'La cadena puede recibir mas de un mensaje,
    'la rutina se encargará de separar estos mensajes
    Dim intPosicion As Integer
    intPosicion = InStrRev(strMensajes, chrCOMIENZO)
    If intPosicion <> 1 Then
        SepararMensajes Mid$(strMensajes, 1, intPosicion - 1), intIndiceOrigenMensaje
    End If
    InterpretarMensaje Mid$(strMensajes, intPosicion + 1, Len(strMensajes) - intPosicion + 1), intIndiceOrigenMensaje
End Sub

Public Function Aleatorio(intDesde As Integer, intHasta As Integer) As Integer
    On Error GoTo ErrorHandle
    
    Aleatorio = Int(Rnd * (intHasta - intDesde + 1)) + intDesde
    
    Exit Function
ErrorHandle:
    ReportErr "Aleatorio", "mdlInterfase", Err.Description, Err.Number, Err.Source
End Function

Public Function Dado() As Integer
    'Devuelve un valor de 1 a 6
    On Error GoTo ErrorHandle
    
    Dado = Aleatorio(1, 6)
        
    Exit Function
ErrorHandle:
    ReportErr "Dado", "mdlServidor", Err.Description, Err.Number, Err.Source
End Function

Public Sub Pausa(intCantMs As Integer, Optional blnDoEvents As Boolean = True)
    On Error GoTo ErrorHandle
    Dim sngInicio As Single
    
    Sleep CLng(intCantMs)
    
    Exit Sub
ErrorHandle:
    ReportErr "Pausa", "mdlInterfase", Err.Description, Err.Number, Err.Source
End Sub

Public Sub IniciarPausa(intCantMs As Integer, ByRef sngPausa As Single)
    On Error GoTo ErrorHandle
    
    sngPausa = Timer   ' Establece la hora de inicio.
    
    Exit Sub
ErrorHandle:
    ReportErr "IniciarPausa", "mdlInterfase", Err.Description, Err.Number, Err.Source
End Sub

Public Sub FinalizarPausa(intCantMs As Integer, sngPausa As Single, Optional blnDoEvents As Boolean = False)
    On Error GoTo ErrorHandle
    Dim sngTimer As Single
    
    sngTimer = Timer
    If sngPausa < sngTimer Then
        If intCantMs - (sngTimer - sngPausa) > 0 Then
            Sleep CLng(intCantMs - (sngTimer - sngPausa))
        End If
    Else
        'Pasó la medianoche
        Sleep CLng(intCantMs)
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "FinalizarPausa", "mdlInterfase", Err.Description, Err.Number, Err.Source
End Sub

Public Function ValidaTexto(strTexto As String, Optional intLongitud As Integer = 0) As Boolean
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim blnAux As Boolean
    
    blnAux = True
    
    For i = 1 To Len(strTexto)
        If InStr("'" & chrSEPARADOR & chrCOMIENZO & chrPREFIJOADM & chrSUFIJOADM, Mid$(strTexto, i, 1)) <> 0 Then
            'Si encontro uno
            blnAux = False
        End If
    Next
    
    If intLongitud > 0 Then
        If intLongitud < Len(strTexto) Then
            blnAux = False
        End If
    End If
    
    ValidaTexto = blnAux
    
    Exit Function
ErrorHandle:
    ReportErr "ValidaTexto", "mdlInterfase", Err.Description, Err.Number, Err.Source
    ValidaTexto = False
End Function

Public Function ValidaEntero(strNumero As String, Optional intMinimo As Integer = 0, Optional intMaximo As Integer = 0) As Boolean
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim blnAux As Boolean
    
    blnAux = True
    
    For i = 1 To Len(strNumero)
        If InStr("0123456789" & Chr(vbKeyBack), Mid$(strNumero, i, 1)) = 0 Then
            'Si no es un numero
            blnAux = False
        End If
    Next
    
    If intMinimo > 0 Then
        If CDbl(intMinimo) > CDbl(strNumero) Then
            blnAux = False
        End If
    End If
    
    If intMaximo > 0 Then
        If CDbl(intMaximo) < CDbl(strNumero) Then
            blnAux = False
        End If
    End If
    
    ValidaEntero = blnAux
    
    Exit Function
ErrorHandle:
    ReportErr "ValidaEntero", "mdlInterfase", Err.Description, Err.Number, Err.Source
    ValidaEntero = False
End Function

Public Function CompilarMensaje(strMascara As String, Optional vecValores As Variant) As String
   ' On Error GoTo ErrorHandle
    'Es similar al printf de C.
    'Dada una mascara arma un mensaje insertàndole los valores
    Dim strMensaje As String
    Dim intPos As Integer
    Dim intNumero As Integer
    
    strMensaje = strMascara
    intPos = InStr(strMascara, chrMASCARAMENSAJE)
    While intPos <> 0
        intNumero = CInt(Mid$(strMensaje, intPos + 1, 2))
        strMensaje = Left$(strMensaje, intPos - 1) & vecValores(intNumero - 1) & Mid$(strMensaje, intPos + 3)
        intPos = InStr(intPos + 1, strMensaje, chrMASCARAMENSAJE)
    Wend
    
    CompilarMensaje = strMensaje
    
    Exit Function
ErrorHandle:
    ReportErr "CompilarMensaje", "mdlInterfase", Err.Description, Err.Number, Err.Source
End Function
