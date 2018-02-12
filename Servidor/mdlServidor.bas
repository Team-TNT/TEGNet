Attribute VB_Name = "mdlServidor"
Option Explicit

Public MODO_DEBUG As String

Public Const CintCantMaxConexiones = 7 '6 jugadores + 1 en reconexion
Public Const CstrErrorNoAdm = "Usted no es el Administrador y no tiene permisos para realizar la acción especificada."
Public Const CintCantMaxTarjetas = 5

'Tipo que contiene el indice y el estado de una conexion (se utiliza para el AYA)
Public Type typRtaAYA
    Indice As Integer
    Estado As enuRtaAYA
End Type

'Enumeración que contiene los distintos estados del servidor
Public Enum enuEstadoSrv
    estEsperandoAdm = 0
    estConfigurandoServidor = 1
    estEsperandoJugadores = 2
    estEjecutandoPartida = 3
    estPartidaDetenida = 4
End Enum

'Enumeración que contiene los estados del jugador activo
Public Enum enuEstadoActivo
    estAgregando
    estAtacando
    estMoviendo
    estTarjetaTomada
    estTarjetaCobrada
    estTarjetaTomadaCobrada
End Enum

'Enumeración que contiene los eventos para el cambio de estado
Public Enum enuEventosServidor
    eveAgregarTropas
    eveAtaque
    eveMovimiento
    evePedidoTarjeta
    eveCobroTarjeta
    eveCanjeTarjeta
'    eveFinTurnoInicial
    eveFinTurnoRecuento
    eveFinTurnoAccion
End Enum

'Enumeración de los posibles estados de respuesta a un AYA
Public Enum enuRtaAYA
    rtaEsperandoRta = 0
    rtaEstaVivo = 1
    rtaEstaMuerto = -1
End Enum

'Enumeración que contiene los códigos de las máscaras
Public Enum enuMascara
    mscConexionFisica = 501
    mscValidacion
    mscDesconexion
    mscInicioPartida
    mscInicioTurno
    mscMisionCumplida
    mscAgregadoTropas
    mscAtaque
    mscResultadoAtaque
    mscConquista
    mscMovimientoConquista1
    mscMovimientoConquistaN
    mscMovimientoTropas
    mscTarjetaTomada
    mscTarjetaCobrada
    mscCanje
    mscJugadorEliminado
    mscJugadorEliminadoTarjetasAsesino
    mscJugadorEliminadoTarjetasMazo
    mscOpcionModificada
    mscAgregado1Tropa
    mscMovimientoConquistaN1
    mscMovimiento1Tropa
    mscReconexion
End Enum
'--------------------------------------------------------------

'### Borrar
Public vecEstadosServidor(10) As String
Public vecEstadosActivo(10) As String

'Estados del servidor
Dim GMatrizEstados(0 To 5, 0 To 8) As enuEstadoActivo

Public GintCantConexiones As Integer
'Numero representativo de la partida actual (Par_Id)
Public GintPartidaActiva As Integer

'Estado del servidor
Public GEstadoServidor As enuEstadoSrv
Public GEstadoActivo As enuEstadoActivo

'Contiene el estado de la respuesta a un AYA
Public GintRtaAYA As enuRtaAYA
Public GvecRtaAYA() As typRtaAYA

'Contiene el indice del pedido de AYA actual (para quien se pidio el AYA)
' Public intIndiceAYA As Integer
'Contiene la cantidad de milisegundos a esperar por un ACK de un AYA
Public GintEsperaAckAYA As Long

'Vector que contiene la relacion entre el socket y el color
'contiene el color que le corresponde a cada socket
Public GvecColoresSock() As Integer

'Vector que contiene la relacion entre el Color y el Nombre
'contiene el nombre que le corresponde a cada color
Public GvecNombreJugadorColor(6) As String

'Indice del Cliente que envió el mensaje que se está procesando
Public GintOrigenMensaje As Integer

'Contiene el indice del socket que representa al administrador
Public GintIndiceAdm As Integer

'Contiene el valor actual del timer del turno
Public GintValorTimerActual As Integer
'Contiene el valor total del timer
Public GintValorTimerTotal As Integer

'Vector con los nombres de los colores
Public GvecColores(6) As String

'Vector con los nombres de los paises
Public GvecPaises() As String

'Vector con los nombres de los continentes
Public GvecContinentes(6) As String

'Valor inicial del timer del turno
Public GsngInicioTimerTurno As Single

'Flag que indica si se debe preguntar antes de hacer el unload
Public blnPreguntar As Boolean


Public Sub DistribuirMensaje(TipoMensaje As enuTipoMensaje, vecParametros() As String, intIndiceOrigenMensaje As Integer)
'Public Sub DistribuirMensaje(TipoMensaje As Integer, vecParametros() As String)
    On Error GoTo ErrorHandle
    
    'Si la Partida esta Pausada y hay Adm
    ', no acepta msgs que no sean del Adm
    '(excepo los mensajes de IAA de los jugadores conectados)
    If (TipoMensaje <> msgIAA) And (GEstadoServidor = estPartidaDetenida) And (GintOrigenMensaje <> GintIndiceAdm) And (GintIndiceAdm <> -1) Then
        If TipoMensaje = msgReconexion Then
            'Envia mensaje de error,
            'avisando que el Servidor esta Pausado.
            EnviarMensaje ArmarMensajeParam(msgAckAltaJugador, enuAckAltaJugador.ackServidorPausado, 0, ""), intIndiceOrigenMensaje
            'Libera el socket ocupado
            'frmServer.wskServer(intIndiceOrigenMensaje).Close
            Exit Sub
        Else
            'Ignora el mensaje
            Exit Sub
        End If
    End If
    
    'De acuerdo al tipo de mensaje
    Select Case TipoMensaje
        Case msgTipoPartida
            sTipoPartida CInt(vecParametros(0))
        Case msgBajarServidor
            sBajarServidor
        Case msgNombrePartida
            sPartidaGuardada CStr(vecParametros(0))
        Case msgAltaJugador
            sAltaJugador CInt(vecParametros(0)), CStr(vecParametros(1))
        Case msgIAA
            sRecibirIAA GintOrigenMensaje
        Case msgPedirConexionesActuales
            sConexionesActuales
        Case msgOpciones
            sRecibirOpciones vecParametros
        Case msgOpcionesDefault
            sRecibirOpcionesDefault vecParametros
        Case msgCambiarAdm
            sCambiarAdm CInt(vecParametros(0)), True
        Case msgFinTurno
            sFinTurno False
        Case msgChatSaliente
            sEnviarMensajeChat CStr(vecParametros(0))
        Case msgIniciarPartida
            sIniciarPartida
        Case msgAgregarTropas
            sAgregarTropas CInt(vecParametros(0)), CInt(vecParametros(1))
        Case msgAtaque
            sAtacar CInt(vecParametros(0)), CInt(vecParametros(1))
        Case msgMovimiento
            sMover CInt(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2)), CInt(vecParametros(3))
        Case msgCobroTarjeta
            sCobrarTarjeta CInt(vecParametros(0))
        Case msgPedidoTarjeta
            sTomarTarjeta
        Case msgCanjeTarjeta
            sCanjearTarjeta CInt(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2))
        Case msgReconexion
            sReconectarJugador CInt(vecParametros(0))
        Case msgResincronizacion
            sResincronizar GintOrigenMensaje
        Case msgPedidoPartidasGuardadas
            sEnviarPartidasGuardadas tpgGuardarPartida
        Case msgEliminarPartida
            sEliminarPartida vecParametros(0)
        Case msgVersionCliente
            sRecibirVersionCliente intIndiceOrigenMensaje, CByte(vecParametros(0)), CInt(vecParametros(1)), CInt(vecParametros(2)), CInt(vecParametros(3))
        Case msgGuardarPartida
            sGuardarPartida vecParametros(0)
        Case msgPausarPartida
            sPausarPartida
        Case enuTipoMensaje.msgContinuarPartida
            sContinuarPartida

    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "DistribuirMensaje", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub EnviarMensaje(strMensaje As String, Optional intDestinatario As Integer)
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    Dim intDestino As Integer
    
    If intDestinatario = 0 Then
        'Broadcast
        For i = 1 To frmServer.wskServer.Count - 1
            intDestino = i
            If frmServer.wskServer(i).State <> 0 Then 'Si no está cerrado
                frmServer.wskServer(i).SendData strMensaje
            End If
        Next
    Else
        'Unicast
        intDestino = intDestinatario
        
        If frmServer.wskServer(intDestinatario).State <> 0 Then 'Si no está cerrado
            frmServer.wskServer(intDestinatario).SendData strMensaje
        End If
    End If
    
    '###Borrar Mensaje
    If MODO_DEBUG = "1" Then
        Open App.Path & "\Mensajes.txt" For Append As #1
        Print #1, Tiempo & " - OUT A: "; Mid$(IIf(intDestinatario = 0, "Broadcast", GvecColores(IndiceAColor(intDestinatario))) & Space(9), 1, 9) & " - " & strMensaje
        Close #1
    End If
    
    
    Exit Sub
ErrorHandle:
    Select Case Err.Number
    Case 40006: 'Protocolo o estado de la conexión inválido para la operación
        'Este error se produce cuando el socket del cliente se encuentra cerrado
        'pero el servidor no lo detectó
        'En este caso no envía error y cierra el socket si está abierto
        If frmServer.wskServer(intDestino).State <> 0 Then
            frmServer.wskServer(intDestino).Close
        End If
        Resume Next
    Case Else
        ReportErr "EnviarMensaje", "mdlServidor", Err.Description, Err.Number, Err.Source
    End Select
End Sub

Public Sub sConfirmarAdm()
    'Confirma al jugador que fue designado Administrador
    On Error GoTo ErrorHandle
    
    '###E
    If GEstadoServidor < estEjecutandoPartida Then
        CambiarEstadoServidor estConfigurandoServidor
    End If
    
    EnviarMensaje ArmarMensajeParam(msgConfirmarAdm, GintCantConexiones), GintIndiceAdm

    Exit Sub
ErrorHandle:
    ReportErr "sConfirmarAdm", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sConfirmarConexion()
    'Valida que no se supere la cantidad máxima de jugadores (6)
    'y envía al cliente los colores disponibles y las conexiones
    'anteriores.
    '###
    'Buscar en en la BD las conexiones actuales
    
    'Enviar al cliente los colores disponibles
    
End Sub

Public Sub sTipoPartida(intTipoPartida As enuTipoPartida)
'Public Sub sTipoPartida(intTipoPartida As Integer)
    'Recibe del cliente el tipo de partida
    '(Nueva o Guardada)
    'Si es nueva la crea y si es Guardada
    'envia el mensaje con las partidas existentes
    On Error GoTo ErrorHandle
    
    'Valida que la acción sea solicitada por el administrador
    If GintOrigenMensaje <> GintIndiceAdm Then
        sEnviarError CstrErrorNoAdm, errNoAccionNoAdm, GintOrigenMensaje
        Exit Sub
    End If
    
    If intTipoPartida = tpNueva Then
        'Crea una nueva partida en la BD
        sPartidaNueva
    ElseIf intTipoPartida = tpGuardada Then
        'Envía partidas existentes
        sEnviarPartidasGuardadas tpgIniciarPartida
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "sTipoPartida", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sEnviarPartidasGuardadas(intTipo As enuTipoPartidaGuardada)
    'Envia a los clientes la lista de jugadores conectados hasta el momento
    On Error GoTo ErrorHandle
    
    Dim rsPartidas As Recordset
    Dim vecPartidas() As String
    
    Set rsPartidas = EjecutarConsulta("SELECT FORMAT(par_Fecha_actu, 'YYYY-MM-DD') & " & _
                                      "par_nombre AS Partida FROM PARTIDAS " & _
                                      "WHERE Par_Id>0 ORDER BY par_nombre")
    RecordsetAVector rsPartidas, 0, vecPartidas
    
    'Inserta en el vector el tipo de partida guardada
    ReDim Preserve vecPartidas(UBound(vecPartidas) + 1)
    vecPartidas(UBound(vecPartidas)) = CStr(intTipo)
    
    EnviarMensaje ArmarMensaje(msgPartidasGuardadas, vecPartidas), GintIndiceAdm
    
    rsPartidas.Close
    Set rsPartidas = Nothing
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarPartidasGuardadas", "mdlServidor", Err.Description, Err.Number, Err.Source
    Set rsPartidas = Nothing

End Sub
Public Sub sBajarServidor()
    'Finaliza la ejecución del servidor
    On Error GoTo ErrorHandle
    
    'Valida que la acción sea solicitada por el administrador
    If GintOrigenMensaje <> GintIndiceAdm Then
        sEnviarError CstrErrorNoAdm, errNoAccionNoAdm, GintOrigenMensaje
        Exit Sub
    End If
    
    blnPreguntar = False
    Unload frmServer

    Exit Sub
ErrorHandle:
    ReportErr "sBajarServidor", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sPartidaNueva()
    'Crea en la Base de datos una nueva partida (@ACTUAL@)
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim varMaxPartida As Variant
    
    'Borra la última partida (@ACTUAL@)
    strSQL = "DELETE FROM Partidas WHERE Par_Nombre = '" & strNombrePartidaActual & "'"
    EjecutarComando strSQL
    
    'Obtiene el número de la última partida
    strSQL = "SELECT MAX(Par_Id) FROM Partidas"
    varMaxPartida = EjecutarConsultaValor(strSQL)
    If IsNull(varMaxPartida) Then
        varMaxPartida = 0
    End If
    
    GintPartidaActiva = CInt(varMaxPartida) + 1
    
    'Crea una nueva partida en la BD
    strSQL = "INSERT INTO Partidas(Par_Id, Par_Nombre, Par_Fecha_Creacion, Par_Fecha_Actu)" & _
             "VALUES (" & CStr(GintPartidaActiva) & ", '" & strNombrePartidaActual & "', Date() , Date())"
    EjecutarComando strSQL
    
    'Copia las opciones por default para la nueva partida
    strSQL = "INSERT INTO Opciones(Opc_Id, Opc_Valor, Par_Id) " & _
             "SELECT Opc_Id, Opc_Valor, '" & GintPartidaActiva & "'" & _
             "FROM Opciones WHERE Par_Id = 0"
    
    EjecutarComando strSQL
    
    'Envia al administrador las opciones por defecto
    sEnviarOpcionesDefault
    'Envia al administrador las opciones de la partida actual
    sEnviarOpciones True
    
    Exit Sub
ErrorHandle:
    ReportErr "sPartidaNueva", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sPartidaGuardada(strNombre As String)
    'Guarda el Id de la partida en una variable global (GintPartidaActiva)
    'Cambia el estado del servidor a Jugando
    'Cambia el estado del jugador activo
    On Error GoTo ErrorHandle
    Dim strSQL As String
    
    'Valida que la acción sea solicitada por el administrador
    If GintOrigenMensaje <> GintIndiceAdm Then
        sEnviarError CstrErrorNoAdm, errNoAccionNoAdm, GintOrigenMensaje
        Exit Sub
    End If
    
    'Copia la partida seleccionada y la pone como la 'actual'
    'Solo en el caso que no sea la actual
    If Trim(strNombre) <> strNombrePartidaActual Then
        CopiarPartida strNombre, strNombrePartidaActual
    End If
    
    strSQL = "SELECT Par_Id FROM Partidas WHERE Par_Nombre = '" & strNombrePartidaActual & "'"
    GintPartidaActiva = CInt(EjecutarConsultaValor(strSQL))
    
    '###E
    'Cambia el estado del servidor
    CambiarEstadoServidor estEjecutandoPartida
    
    'Cambia el estado del jugador activo
    strSQL = "SELECT Par_Activo_Estado " & _
             "FROM Partidas " & _
             "WHERE Par_Id = " & GintPartidaActiva
    GEstadoActivo = CInt(EjecutarConsultaValor(strSQL))
    '### Borrar
    frmServer.lblEstadoJugadorActivo.Caption = vecEstadosActivo(GEstadoActivo)
    
    'Envia al administrador las opciones por defecto
    sEnviarOpcionesDefault
    'Envia al administrador las opciones de la partida actual
    sEnviarOpciones True
    
    Exit Sub
ErrorHandle:
    ReportErr "sPartidaGuardada", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sConexionesActuales(Optional intDestino As Integer)
    'Envia a los clientes los nombres y colores de los jugadores conectados
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    Dim strSQL As String
    Dim rsJugadores As Recordset
    Dim vecJugadoresConectados(0 To 30) As String
    Dim intEstado As enuEstadoConexion
    Dim strTag As String
    
    strSQL = "SELECT Col_Id, Jug_Nombre FROM Jugadores WHERE Par_Id=" & GintPartidaActiva
    Set rsJugadores = EjecutarConsulta(strSQL)
    
    While Not rsJugadores.EOF
        'Si hay un Adm...
        If GintIndiceAdm >= 0 Then
            'Si el jugador es el Administrador le agrega '@' al nombre
            If rsJugadores!Col_Id = IndiceAColor(GintIndiceAdm) Then
                vecJugadoresConectados(rsJugadores!Col_Id - 1) = chrPREFIJOADM & rsJugadores!Jug_Nombre & chrSUFIJOADM
            Else
                vecJugadoresConectados(rsJugadores!Col_Id - 1) = rsJugadores!Jug_Nombre
            End If
        Else
            vecJugadoresConectados(rsJugadores!Col_Id - 1) = rsJugadores!Jug_Nombre
        End If
        
        'Guarda la relacion Nombre-Color
        '(vector usado para los msgs del LOG)
        GvecNombreJugadorColor(rsJugadores!Col_Id) = vecJugadoresConectados(rsJugadores!Col_Id - 1)
        
        rsJugadores.MoveNext
    Wend
    
    'Agrega el estado de la conexion para cada jugador (conectado/desconectado)
    For i = 1 To 6
        If ColorAIndice(i) = 0 Then
            If vecJugadoresConectados(i - 1) = "" Then
                'Si el jugador nunca se conecto
                 intEstado = conNoJuega
            Else
                'Se esta cargando una partida guardada
                intEstado = conDesconectado
            End If
        Else
            'Si el jugador está o estaba conectado
            If frmServer.wskServer(ColorAIndice(i)).State = 0 Then
                'El jugador esta desconectado
                intEstado = conDesconectado
            Else
                'El jugador esta conectado
                intEstado = conConectado
            End If
        End If
        vecJugadoresConectados(5 + i) = CStr(intEstado)
    Next i
    
    'Agrega el Tipo de Inteligencia, la IP y la Version de cada Jugador
    For i = 1 To 6
        'Si no existe el Jugador manda los datos en blanco
        If vecJugadoresConectados(i - 1) = "" Then
            vecJugadoresConectados(11 + i) = "0"
            vecJugadoresConectados(17 + i) = ""
            vecJugadoresConectados(23 + i) = ""
        Else
            vecJugadoresConectados(17 + i) = frmServer.wskServer(ColorAIndice(i)).LocalIP
            strTag = frmServer.wskServer(ColorAIndice(i)).Tag
            If strTag = "" Then
                '# Version 1.0.0
                vecJugadoresConectados(11 + i) = "0"
                vecJugadoresConectados(23 + i) = "1.0.0"
            Else
                vecJugadoresConectados(11 + i) = Mid$(strTag, 1, 1)
                vecJugadoresConectados(23 + i) = Val(Mid$(strTag, 2, Len(strTag) - 5)) & "." & Val(Mid$(strTag, Len(strTag) - 3, 2)) & "." & Val(Mid$(strTag, Len(strTag) - 1))
            End If
        End If
    Next i
    
    'En la última posicion,
    'agrega el flag Tipo al vector de Jugadores,
    'segun el estado del servidor
    Select Case GEstadoServidor
        Case enuEstadoSrv.estEjecutandoPartida, enuEstadoSrv.estPartidaDetenida
            'Reconexion
            vecJugadoresConectados(UBound(vecJugadoresConectados)) = enuTipoJugadoresConectados.tjcReconexion
        Case Else
            vecJugadoresConectados(UBound(vecJugadoresConectados)) = enuTipoJugadoresConectados.tjcIngreso
    End Select
    
    EnviarMensaje ArmarMensaje(msgJugadoresConectados, vecJugadoresConectados), intDestino
    
    rsJugadores.Close
    Set rsJugadores = Nothing
    
    Exit Sub
ErrorHandle:
    ReportErr "sConexionesActuales", "mdlServidor", Err.Description, Err.Number, Err.Source
    Set rsJugadores = Nothing
End Sub

Public Sub sAltaJugador(intNuevoColor As Integer, strNuevoNombre As String)
    'Da de alta un nuevo jugador si no existe alguno conectado
    'con el mismo color o el mismo nombre
    On Error GoTo ErrorHandle
    
    Dim strSQL As String
    Dim colorOK As Boolean
    Dim nombreOK As Boolean
    Dim CodErr As enuAckAltaJugador
    
    'Valida que el jugador no se pueda dar de alta en una partida que
    ' se está ejecutando (solo puede reconectarse)
    If GEstadoServidor >= estEjecutandoPartida Then
        sEnviarError "En el servidor se está ejecutando una partida y usted no puede darse de alta con un nuevo jugador. Intente reconectarse.", errNoAltaPartidaIniciada, GintOrigenMensaje
        frmServer.wskServer(GintOrigenMensaje).Close
    End If
    
    strSQL = "SELECT COUNT(*) FROM Jugadores " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             " AND Jug_Nombre LIKE '" & Trim(strNuevoNombre) & "'"
    nombreOK = IIf(CInt(EjecutarConsultaValor(strSQL)) > 0, False, True)
    
    strSQL = "SELECT COUNT(*) FROM Jugadores " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             " AND col_Id = " & intNuevoColor & ""
    colorOK = IIf(CInt(EjecutarConsultaValor(strSQL)) > 0, False, True)
    
    If colorOK And nombreOK Then
        '###V Virtual
        'Lo da de alta en la base de datos
        strSQL = "INSERT INTO Jugadores (Par_Id, Jug_Nombre, Col_Id, Jug_Tipo, Jug_Nro_Canje) " & _
                 "VALUES(" & GintPartidaActiva & ", '" & Trim(strNuevoNombre) & "'," & intNuevoColor & ", '" & _
                 IIf(GintOrigenMensaje = GintIndiceAdm, "A", "N") & "', 0)"
        
        EjecutarComando strSQL
        
        'Guarda la relación Socket-Color
        GvecColoresSock(GintOrigenMensaje) = intNuevoColor
        
        'Prepara la BD para el nuevo jugador, luego serán todos UPDATES
        '-Tropas Disponibles x Continente
        strSQL = "INSERT INTO Tropas_Disponibles (Col_Id, Par_Id, Con_Id, TDI_Cantidad) " & _
                 "SELECT Jug.Col_Id, Jug.Par_Id, Con.Con_Id, 0 " & _
                 "FROM Jugadores Jug, Continentes Con " & _
                 "WHERE Jug.Par_Id= " & GintPartidaActiva & _
                 "  AND Jug.Col_Id = " & intNuevoColor
        EjecutarComando strSQL
        
        strSQL = "INSERT INTO Tropas_Disponibles (Col_Id, Par_Id, TDI_Cantidad) " & _
                 "SELECT Jug.Col_Id, Jug.Par_Id, 0 " & _
                 "FROM Jugadores Jug " & _
                 "WHERE Jug.Par_Id=" & GintPartidaActiva & _
                 "  AND Jug.Col_Id = " & intNuevoColor
        EjecutarComando strSQL
        
        '###Log
        GuardarLog strNuevoNombre & " ha sido validado con el color " & GvecColores(intNuevoColor)
        sEnviarLog mscValidacion, CstrTipoParametroResuelto & strNuevoNombre, CstrTipoParametroRecurso & CStr(intNuevoColor + enuIndiceArchivoRecurso.pmsColores)
        
        'Confirma Color a Cliente
        sConfirmarAlta intNuevoColor, strNuevoNombre, 0
    
    Else
        'Envia mensaje de error
        If Not colorOK And nombreOK Then
            CodErr = ackColorUsado
        ElseIf colorOK And Not nombreOK Then
            CodErr = AckNombreUsado
        Else
            CodErr = ackNombreYColorUsados
        End If
        sConfirmarAlta intNuevoColor, strNuevoNombre, CodErr
    End If

    'Envía mensaje de usuarios conectados
    sConexionesActuales
    
    'Si el jugador es el Administrador...
    If GintOrigenMensaje = GintIndiceAdm Then
        EnviarMensaje ArmarMensajeParam(msgIpServidor, frmServer.wskServer(0).LocalIP), GintIndiceAdm
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "sAltaJugador", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sReconectarJugador(intColor As Integer)
    'Verifica que el color sea un jugador de la partida
    'Verifica que ese color no esté conectado (AYA)
    'Actualiza el vector con los colores y los indices
    'Resincroniza el jugador
    On Error GoTo ErrorHandle
    Dim intOrigenMensaje As Integer
    Dim strSQL As String
    Dim colorOK As Boolean
    Dim CodErr As enuAckAltaJugador
    
    'Guarda el indice del jugador que intenta reconectarse
    intOrigenMensaje = GintOrigenMensaje
    
    'Verifica que el color sea un jugador de la partida
    strSQL = "SELECT COUNT(*) FROM Jugadores " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             " AND col_Id = " & intColor & ""
    colorOK = IIf(CInt(EjecutarConsultaValor(strSQL)) = 0, False, True)
    
    If Not colorOK Then
        'Envia mensaje de error,
        'porque el color seleccionado no pertenece a ningun jugador
        EnviarMensaje ArmarMensajeParam(msgAckAltaJugador, enuAckAltaJugador.ackColorInexistente, intColor, ""), intOrigenMensaje
        'Desconecta al cliente que intentó reconectarse
'        frmServer.wskServer(intOrigenMensaje).Close
        Exit Sub
    End If
    
    'Verifica que ese color no esté conectado (AYA)
    'Verifica si está cerrado el socket del color seleccionado
    If ColorAIndice(intColor) <> 0 Then
        If frmServer.wskServer(ColorAIndice(intColor)).State <> 0 Then
            'Socket no cerrado
            'Envia AYA
            If estaVivo(ColorAIndice(intColor)) Then
                'Envia mensaje de error,
                'porque el color seleccionado está conectado
                EnviarMensaje ArmarMensajeParam(msgAckAltaJugador, enuAckAltaJugador.ackColorConectado, intColor, ""), intOrigenMensaje
                'Desconecta al cliente que intentó reconectarse
    '            frmServer.wskServer(intOrigenMensaje).Close
                Exit Sub
            Else
                'Si no esta vivo
                'Cierra el socket
                frmServer.wskServer(ColorAIndice(intColor)).Close
            End If
        Else
            'Si no esta vivo
            'Cierra el socket
            frmServer.wskServer(ColorAIndice(intColor)).Close
        End If
    End If
    
    'Guarda la relación Socket-Color
    GvecColoresSock(ColorAIndice(intColor)) = 0
    GvecColoresSock(intOrigenMensaje) = intColor
    
    'Confirma Color a Cliente
    sConfirmarAlta intColor, "", 0

    'Envía mensaje de usuarios conectados
    sConexionesActuales

    '###Log
    GuardarLog GvecNombreJugadorColor(intColor) & " se ha reconectado."
    sEnviarLog mscReconexion, CstrTipoParametroResuelto & GvecNombreJugadorColor(intColor)
    
    'Antes de confirmar el inicio de partida, envía los países limitrofes y los paises por continente
    sEnviarLimitrofes intOrigenMensaje
    sEnviarPaisContinente intOrigenMensaje
    
    'Inicio partida
    sConfirmarInicioPartida intOrigenMensaje
    
    'Resincroniza al jugador que se reconecta
    sResincronizar intOrigenMensaje
        
    Exit Sub
ErrorHandle:
    ReportErr "sReconectarJugador", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sConfirmarAlta(intColorAsignado As Integer, strNombreAsignado As String, intCodAck As enuAckAltaJugador)
    On Error GoTo ErrorHandle
    
    '###E
    If GintIndiceAdm = GintOrigenMensaje And GEstadoServidor < estEjecutandoPartida Then
        CambiarEstadoServidor estEsperandoJugadores
    End If
    
    EnviarMensaje ArmarMensajeParam(msgAckAltaJugador, intCodAck, intColorAsignado, strNombreAsignado), GintOrigenMensaje
    
    Exit Sub
ErrorHandle:
    ReportErr "sConfirmarAlta", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sEnviarMensajeChat(strMensaje As String)
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgChatEntrante, GvecColoresSock(GintOrigenMensaje), strMensaje)
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarMensajeChat", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sEnviarAYA(intCliente As Integer)
    On Error GoTo ErrorHandle
    
    'Envia el AYA al cliente
    EnviarMensaje ArmarMensajeParam(msgAYA), intCliente
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarAYA", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sRecibirIAA(intCliente As Integer)
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    Dim intEsperado As Integer
    Dim respondieronTodos As Boolean
    
    intEsperado = -1
    respondieronTodos = True
    
    For i = LBound(GvecRtaAYA) To UBound(GvecRtaAYA)
        If GvecRtaAYA(i).Indice = intCliente Then
            'Guarda la posición del cliente que respondio
            intEsperado = i
        End If
    Next i
    
    'Si el cliente a la espera de AYA es el que respondió
    If intEsperado >= 0 Then
        GvecRtaAYA(intEsperado).Estado = rtaEstaVivo
    End If
    
    'Si ya respondieron todos, desactiva el timer
    For i = LBound(GvecRtaAYA) To UBound(GvecRtaAYA)
        If GvecRtaAYA(i).Estado = rtaEsperandoRta Then
            respondieronTodos = False
        End If
    Next i
    
    If respondieronTodos Then
        frmServer.tmrAYA.Interval = 0
        GintRtaAYA = rtaEstaVivo
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "sRecibirIAA", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Function estaVivo(intCliente As Integer) As Boolean
    On Error GoTo ErrorHandle
    
    'Envia el AYA
    sEnviarAYA intCliente
    
    ReDim GvecRtaAYA(0)
    
    'Activa el timer y se queda a la espera de respuesta
    GintRtaAYA = rtaEsperandoRta
    GvecRtaAYA(0).Indice = intCliente
    GvecRtaAYA(0).Estado = rtaEsperandoRta
    frmServer.tmrAYA.Interval = GintEsperaAckAYA
    
    While GintRtaAYA = rtaEsperandoRta
        DoEvents
    Wend
    
    If GintRtaAYA = rtaEstaVivo Then
        estaVivo = True
    Else
        estaVivo = False
    End If
    
    Exit Function
ErrorHandle:
    ReportErr "estaVivo", "mdlServidor", Err.Description, Err.Number, Err.Source
End Function

Public Sub sEnviarOpciones(Optional blnMostrarEnPantalla As Boolean = False, Optional intIndiceDestino As Integer = 0)
    'Toma de la base de datos las opciones de la partida actual
    On Error GoTo ErrorHandle
    
    Dim strSQL As String
    Dim rsOpciones As Recordset
    Dim vecOpciones() As String
    
    strSQL = "SELECT Opc_Id & '" & chrSEPARADOR & "' & Opc_Valor AS Opcion " & _
             "FROM Opciones WHERE Par_Id =" & GintPartidaActiva & ""
             
    Set rsOpciones = EjecutarConsulta(strSQL)
    
    RecordsetAVector rsOpciones, 0, vecOpciones
    
    'Agrega el flag de Mostrar en Pantalla
    ReDim Preserve vecOpciones(UBound(vecOpciones) + 1)
    vecOpciones(UBound(vecOpciones)) = CInt(blnMostrarEnPantalla)
    
    'Broadcast (Todos tienen que saber cuales son las opciones)
    EnviarMensaje ArmarMensaje(msgOpciones, vecOpciones), intIndiceDestino
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarOpciones", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sEnviarOpcionesDefault()
    'Toma de la base de datos las opciones por default
    On Error GoTo ErrorHandle
    
    Dim strSQL As String
    Dim rsOpciones As Recordset
    Dim vecOpciones() As String
    
    strSQL = "SELECT Opc_Id & '" & chrSEPARADOR & "' & Opc_Valor AS Opcion " & _
             "FROM Opciones WHERE Par_Id = 0"
             
    Set rsOpciones = EjecutarConsulta(strSQL)
    
    RecordsetAVector rsOpciones, 0, vecOpciones
    
    EnviarMensaje ArmarMensaje(msgOpcionesDefault, vecOpciones), GintIndiceAdm
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarOpcionesDefault", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sRecibirOpciones(vecOpcionesMsg() As String)
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    Dim strSQL As String
    Dim rsOpcionesActuales As Recordset
    Dim intSiNoVieja As Integer
    Dim intSiNoNueva As Integer
    
    'Valida que la acción sea solicitada por el administrador
    If GintOrigenMensaje <> GintIndiceAdm Then
        sEnviarError CstrErrorNoAdm, errNoAccionNoAdm, GintOrigenMensaje
        Exit Sub
    End If
    
    'Toma de la BD las opciones actuales
    strSQL = "SELECT Opc_Id, Opc_Valor FROM Opciones " & _
             "WHERE Par_Id = " & GintPartidaActiva
    Set rsOpcionesActuales = EjecutarConsulta(strSQL)
    
    'Compara las opciones recibidas con las actuales para informar
    'cuales se modificaron
    While Not rsOpcionesActuales.EOF
        i = LBound(vecOpcionesMsg)
        While rsOpcionesActuales!Opc_Id <> vecOpcionesMsg(i) And i < UBound(vecOpcionesMsg) - 1
            i = i + 2
        Wend
        
        If rsOpcionesActuales!Opc_Id = vecOpcionesMsg(i) Then
            If vecOpcionesMsg(i + 1) <> rsOpcionesActuales!Opc_Valor Then
                'Si se modificó la opción...
                If IsNumeric(rsOpcionesActuales!Opc_Valor) Then
                    '###Log
                    GuardarLog "El Administrador de la Partida ha modificado la opción número " & rsOpcionesActuales!Opc_Id & " de " & rsOpcionesActuales!Opc_Valor & " a " & vecOpcionesMsg(i + 1) & "."
                    sEnviarLog mscOpcionModificada, CstrTipoParametroRecurso & CStr(rsOpcionesActuales!Opc_Id + enuIndiceArchivoRecurso.pmsOpciones), CstrTipoParametroResuelto & CStr(rsOpcionesActuales!Opc_Valor), CstrTipoParametroResuelto & CStr(vecOpcionesMsg(i + 1))
                Else
                    Select Case rsOpcionesActuales!Opc_Valor
                        Case "S", "F", "M"
                            intSiNoVieja = 1
                            intSiNoNueva = 2
                        Case "N", "R", "C"
                            intSiNoVieja = 2
                            intSiNoNueva = 1
                    End Select
                
                    '###Log
                    GuardarLog "El Administrador de la Partida ha modificado la opción número " & rsOpcionesActuales!Opc_Id & " de " & rsOpcionesActuales!Opc_Valor & " a " & vecOpcionesMsg(i + 1) & "."
                    sEnviarLog mscOpcionModificada, CstrTipoParametroRecurso & CStr(rsOpcionesActuales!Opc_Id + enuIndiceArchivoRecurso.pmsOpciones), CstrTipoParametroRecurso & CStr(intSiNoVieja + enuIndiceArchivoRecurso.pmsSiNo), CstrTipoParametroRecurso & CStr(intSiNoNueva + enuIndiceArchivoRecurso.pmsSiNo)
                End If
                    
            End If
        End If
        rsOpcionesActuales.MoveNext
    Wend

    rsOpcionesActuales.Close
    Set rsOpcionesActuales = Nothing
    
    'En base al vector recibido actualiza las opciones en la BD
    For i = LBound(vecOpcionesMsg) To UBound(vecOpcionesMsg) Step 2 'Toma de a pares
        strSQL = "UPDATE Opciones SET " & _
                 "Opc_Valor = '" & vecOpcionesMsg(i + 1) & "' " & _
                 "WHERE Par_Id = " & GintPartidaActiva & " " & _
                 "  AND Opc_Id = " & vecOpcionesMsg(i) & " "
        
        EjecutarComando strSQL
    Next i
    
    'Envia las nuevas opciones (broadcast)
    sEnviarOpciones
    
    Exit Sub
ErrorHandle:
    ReportErr "sRecibirOpciones", "mdlServidor", Err.Description, Err.Number, Err.Source
    Set rsOpcionesActuales = Nothing
End Sub

Public Sub sRecibirOpcionesDefault(vecOpcionesMsg() As String)
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    Dim strSQL As String
    
    'Valida que la acción sea solicitada por el administrador
    If GintOrigenMensaje <> GintIndiceAdm Then
        sEnviarError CstrErrorNoAdm, errNoAccionNoAdm, GintOrigenMensaje
        Exit Sub
    End If
    
    'En base al vector recibido actualiza las opciones en la BD
    For i = LBound(vecOpcionesMsg) To UBound(vecOpcionesMsg) Step 2 'Toma de a pares
        strSQL = "UPDATE Opciones SET " & _
                 "Opc_Valor = '" & vecOpcionesMsg(i + 1) & "' " & _
                 "WHERE Par_Id = 0 " & _
                 "  AND Opc_Id = " & vecOpcionesMsg(i) & " "
        
        EjecutarComando strSQL
    Next i

    Exit Sub
ErrorHandle:
    ReportErr "sRecibirOpcionesDefault", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CambiarEstadoServidor(NuevoEstado As enuEstadoSrv)
    'Cambia el estado del cliente de acuerdo al nuevo estado pasado por parametro
    On Error GoTo ErrorHandle
    
    GEstadoServidor = NuevoEstado
    
    frmServer.lblEstado = vecEstadosServidor(GEstadoServidor)
    
    Exit Sub
ErrorHandle:
        ReportErr "CambiarEstadoServidor", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sBajaJugador(intIndiceBaja As Integer)
    'Desconecta lógicamente al jugador involucrado
    On Error GoTo ErrorHandle
    
    Dim strSQL As String
    
    If GEstadoServidor < estEjecutandoPartida Then
        'Elimina al jugador de la BD
        strSQL = "DELETE FROM Jugadores " & _
                 "WHERE Col_Id = " & GvecColoresSock(intIndiceBaja) & _
                 "  AND Par_Id = " & GintPartidaActiva
        
        EjecutarComando strSQL
    End If
    
    'Actualiza la relación Color-Socket
    'GvecColoresSock(intIndiceBaja) = 0
    
    'Solo si el que se desconectó era un jugador...
    If GvecNombreJugadorColor(IndiceAColor(intIndiceBaja)) <> "" Then
        'Refresca a los clientes las conexiones actuales
        sConexionesActuales
    
        '###Log
        GuardarLog GvecNombreJugadorColor(IndiceAColor(intIndiceBaja)) & " se ha desconectado."
        sEnviarLog mscDesconexion, CstrTipoParametroResuelto & GvecNombreJugadorColor(IndiceAColor(intIndiceBaja))
    End If
    
    Exit Sub
ErrorHandle:
        ReportErr "sBajaJugador", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sEnviarError(strDescripcion As String, intCodigoError As enuErrores, Optional intClienteDestino As Integer = 0)
    'Envía un mensaje rápido a un cliente
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgError, strDescripcion, intCodigoError), intClienteDestino
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarError", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sCambiarAdm(intColorNuevoAdm As Integer, blnNoBajaAdm As Boolean)
    'Cambia al jugador administrador de la partida
    On Error GoTo ErrorHandle
    
    'Valida que la acción sea solicitada por el administrador
    If GintOrigenMensaje <> GintIndiceAdm Then
        sEnviarError CstrErrorNoAdm, errNoAccionNoAdm, GintOrigenMensaje
        Exit Sub
    End If
    
    If blnNoBajaAdm Then
        sConfirmarBajaAdm
    End If
    
    GintIndiceAdm = ColorAIndice(intColorNuevoAdm)
    sConfirmarAdm
    
    'Refresca los nombres de los clientes (@ al administrador)
    sConexionesActuales
    
    'Pasa las opciones por defecto de la partida al nuevo administrador
    sEnviarOpcionesDefault
    
    Exit Sub
ErrorHandle:
    ReportErr "sCambiarAdm", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sConfirmarBajaAdm()
    'Informa al administrador que ha dejado de serlo
    On Error GoTo ErrorHandle

    'Valida que la acción sea solicitada por el administrador
    If GintOrigenMensaje <> GintIndiceAdm Then
        sEnviarError CstrErrorNoAdm, errNoAccionNoAdm, GintOrigenMensaje
        Exit Sub
    End If
    
    EnviarMensaje ArmarMensajeParam(msgBajaAdm), GintOrigenMensaje
    
    Exit Sub
ErrorHandle:
    ReportErr "sConfirmarBajaAdm", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Function IndiceAColor(intIndice As Integer) As Integer
    'Dado el índice del socket devuelve el color del jugador correspondiente
    On Error GoTo ErrorHandle
    
    IndiceAColor = GvecColoresSock(intIndice)
    
    Exit Function
ErrorHandle:
    ReportErr "IndiceAColor", "mdlServidor", Err.Description, Err.Number, Err.Source
End Function

Public Function ColorAIndice(intColor As Integer) As Integer
    'Dado el color del jugador devuelve el índice de conexión del socket
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    
    For i = 1 To UBound(GvecColoresSock)
        If intColor = GvecColoresSock(i) Then
            ColorAIndice = i
            Exit For
        End If
    Next i
    
    Exit Function
ErrorHandle:
    ReportErr "ColorAIndice", "mdlServidor", Err.Description, Err.Number, Err.Source
End Function

Public Function ColorANombre(intColor As Integer) As String
    'Dado el Color de un jugador, devuelve su Nombre
    On Error GoTo ErrorHandle
    
    ColorANombre = "pepe"
    
    Exit Function
ErrorHandle:
    ReportErr "ColorANombre", "mdlServidor", Err.Description, Err.Number, Err.Source
End Function

Public Sub sIniciarPartida()
    'Inicia la partida
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim intColor As Integer
    '### Sincronizar Jugador??
    
    'Cambia el estado del servidor
    '###E
    CambiarEstadoServidor estEjecutandoPartida
    
    'Antes de confirmar el inicio de partida, envía los países limitrofes y los paises por continente
    sEnviarLimitrofes
    sEnviarPaisContinente
    
    sConfirmarInicioPartida
    
    '###Log
    GuardarLog "Ha comenzado la partida."
    sEnviarLog mscInicioPartida
    
    sRepartirPaises
    sActualizarMapa
    
    sEstablecerRonda
    sActualizarRonda
    
    'Necesita estar despues de establecer ronda
    sRepartirMisiones
    sEnviarMisiones
    
    'Informa a los clientes el tipo de ronda
    sInformarTipoRonda

    sTropasDisponibles
    sInformarTropasDisponiblesTodos
    
    'Envia las opciones (broadcast) a todos los jugadores
    sEnviarOpciones
    
    sInformarInicioTurno True
    
    Exit Sub
ErrorHandle:
    ReportErr "sIniciarPartida", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sRepartirPaises()
    'Reparte en forma aleatoria los paises entre los clientes y
    'los actualiza en la BD
    On Error GoTo ErrorHandle
    Dim intTropasInicial As Integer
    Dim intCantPaises As Integer
    Dim intCantJugadores As Integer
    Dim intCantMaxPaises As Integer
    Dim strSQL As String
    Dim vecCantPaisesAsignadosPorColor(intCANTMAXJUGADORES) As Integer
    Dim vecColoresActivos() As String
    Dim vecPaises() As Integer
    Dim intColor As Integer
    Dim intPais As Integer
    Dim intPaisSorteado As Integer
    Dim intIndiceSorteado As Integer
    Dim intCantJugadoresConMaxPaises As Integer
    Dim intMaxJugadoresConMaxPaises As Integer
    Dim intColorSorteado As Integer
    Dim i As Integer
    
    'Obtiene la cantidad de tropas por pais en el reparto inicial
    intTropasInicial = CInt(ValorOpcion(opTropasInicial))
    
    'Obtiene la cantidad de jugadores
    strSQL = "SELECT COUNT(*) FROM Jugadores WHERE Par_Id = " & GintPartidaActiva
    intCantJugadores = EjecutarConsultaValor(strSQL)
    ReDim vecColoresActivos(intCantJugadores - 1)
    
    'Obtiene la cantidad de paises
    strSQL = "SELECT COUNT(*) FROM Paises"
    intCantPaises = EjecutarConsultaValor(strSQL) - 1
    
    'Obtiene los colores activos
    strSQL = "SELECT Col_Id FROM Jugadores WHERE Par_Id = " & GintPartidaActiva
    RecordsetAVector EjecutarConsulta(strSQL), 0, vecColoresActivos
    
    ' Inicializa el generador de números aleatorios.
    Randomize
    
    'Calcula el maximo de paises por color
    intMaxJugadoresConMaxPaises = intCantPaises Mod intCantJugadores
    If intMaxJugadoresConMaxPaises = 0 Then
        intCantMaxPaises = intCantPaises / intCantJugadores
        'Todos pueden llegar al máximo
        intMaxJugadoresConMaxPaises = intCantJugadores
    Else
        intCantMaxPaises = Int(intCantPaises / intCantJugadores) + 1
    End If
    
    'Limpia las cantidades de cada color
    For intColor = LBound(vecCantPaisesAsignadosPorColor) To UBound(vecCantPaisesAsignadosPorColor)
        vecCantPaisesAsignadosPorColor(intColor) = 0
    Next intColor
    
    intCantJugadoresConMaxPaises = 0
    
    'Inicializa el vector de paises
    ReDim vecPaises(intCantPaises)
    For intPais = 1 To intCantPaises
        vecPaises(intPais) = intPais
    Next intPais
    
    For intPais = 1 To intCantPaises
    
        'Toma al azar un Pas_Id del vector
        intIndiceSorteado = Aleatorio(1, intCantPaises - intPais + 1)
        intPaisSorteado = vecPaises(intIndiceSorteado)
    
        intColorSorteado = vecColoresActivos(Aleatorio(LBound(vecColoresActivos), intCantJugadores - 1))
        
        If (vecCantPaisesAsignadosPorColor(intColorSorteado) = intCantMaxPaises - 1 And intCantJugadoresConMaxPaises < intMaxJugadoresConMaxPaises) Or _
           (vecCantPaisesAsignadosPorColor(intColorSorteado) < intCantMaxPaises - 1) Then
            
            vecCantPaisesAsignadosPorColor(intColorSorteado) = vecCantPaisesAsignadosPorColor(intColorSorteado) + 1
            
            'Asigna el pais al jugador sorteado en la BD
            strSQL = "INSERT INTO Tropas (Tro_Cantidad, Par_Id, Pas_Id, Col_Id, Tro_Fijos) " & _
                     " VALUES(" & intTropasInicial & ", " & GintPartidaActiva & ", " & intPaisSorteado & ", " & intColorSorteado & ", 0)"

            EjecutarComando strSQL
            
            If vecCantPaisesAsignadosPorColor(intColorSorteado) = intCantMaxPaises Then
                intCantJugadoresConMaxPaises = intCantJugadoresConMaxPaises + 1
            End If
            
            'Borra el pais sorteado del vector
            For i = intIndiceSorteado To intCantPaises - 1
                vecPaises(i) = vecPaises(i + 1)
            Next i

        Else
            'Vuelve a sortear el pais
            intPais = intPais - 1
        End If
        
    Next intPais
    
    Exit Sub
ErrorHandle:
    ReportErr "sRepartirPaises", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sRepartirMisiones()
    'Reparte en forma aleatoria las misiones entre los clientes y
    'los actualiza en la BD
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim strMisionDestruir As String
    Dim intCantMisiones As Integer
    Dim vecMisiones() As String
    Dim vecMisionesAsignadas() As Integer
    Dim vecColoresActivos() As String
    Dim intMisionSorteada As Integer
    Dim intIndiceMisionSorteada As Integer
    Dim varColorVictima As Variant
    Dim intColorVictima As Integer
    Dim blnExisteVictima As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim blnMisionYaSorteada As Boolean
    Dim blnConquistarMundo As Boolean
    
    'Obtiene los colores activos
    strSQL = "SELECT Col_Id FROM Jugadores WHERE Par_Id = " & GintPartidaActiva
    RecordsetAVector EjecutarConsulta(strSQL), 0, vecColoresActivos
    
    ReDim vecMisionesAsignadas(UBound(vecColoresActivos))
    
    'Inicializa vector misiones asignadas
    For i = LBound(vecMisionesAsignadas) To UBound(vecMisionesAsignadas)
        vecMisionesAsignadas(i) = -1
    Next i
    
    'Toma de la BD el tipo de mision (por objetivo o conquista mundial)
    blnConquistarMundo = IIf(Trim(UCase(ValorOpcion(opMisionTipo))) = "M", False, True)
    
    'Si se juega a conquistar el mundo, a todos le asigna la misión 0
    If blnConquistarMundo Then
        For i = LBound(vecColoresActivos) To UBound(vecColoresActivos)
            'Guarda en la BD la misión asignada
            strSQL = "UPDATE Jugadores SET Mis_Id = 0 " & _
                     "WHERE Col_Id = " & vecColoresActivos(i) & _
                     "  AND Par_Id = " & GintPartidaActiva
            EjecutarComando strSQL
        Next i
        Exit Sub
    End If
    
    'Si no se juega a conquistar el mundo se reparten las misiones...
    
    ' Inicializa el generador de números aleatorios.
    Randomize
    
    'Opcion que indica si el juego incluye Misiones de Destruir
    strMisionDestruir = ValorOpcion(opMisionDestruir)
    
    If strMisionDestruir = "S" Then
        'Incluye misiones destruir
        strSQL = "SELECT Mis_Id FROM Misiones WHERE Mis_Id > 0"
    Else
        'No incluye misiones destruir
        strSQL = "SELECT DISTINCT M.Mis_Id FROM Misiones M, Misiones_Objetivos O " & _
                 "WHERE M.Mis_Id = O.Mis_Id " & _
                 "  AND O.Col_Id IS NULL" & _
                 "  AND M.Mis_Id > 0"
    End If
    RecordsetAVector EjecutarConsulta(strSQL), 0, vecMisiones
    
    For i = LBound(vecColoresActivos) To UBound(vecColoresActivos)
        intIndiceMisionSorteada = Aleatorio(LBound(vecMisiones), UBound(vecMisiones))
        intMisionSorteada = vecMisiones(intIndiceMisionSorteada)
        blnMisionYaSorteada = False
        'Busca si la mision ya fue sorteada
        For j = LBound(vecMisionesAsignadas) To UBound(vecMisionesAsignadas)
            If vecMisionesAsignadas(j) = intMisionSorteada Then
                blnMisionYaSorteada = True
            End If
        Next j
        If blnMisionYaSorteada Then
            'Volver a sortear
            i = i - 1
        Else
            'Asigno la misión
            
            'Chequea que si la misión es de destruir la víctima exista
            strSQL = "SELECT Col_Id FROM Misiones_Objetivos " & _
                     "WHERE Mis_Id = " & intMisionSorteada
            varColorVictima = EjecutarConsultaValor(strSQL)
            
            If Not IsNull(varColorVictima) Then
                intColorVictima = CInt(varColorVictima)
                blnExisteVictima = False
                For j = LBound(vecColoresActivos) To UBound(vecColoresActivos)
                    If vecColoresActivos(j) = intColorVictima And j <> i Then
                        'Si existe la victima y no soy yo
                        blnExisteVictima = True
                    End If
                Next j
                
                If Not blnExisteVictima Then
                    'Si la victima no existe, toma la mision cuya víctima es el próximo
                    'jugador de la ronda (el de la derecha)
                    strSQL = "SELECT Mob.Mis_Id " & _
                             "FROM Misiones_Objetivos Mob, Jugadores Jug " & _
                             "WHERE Mob.Col_Id = Jug.Jug_Prox_Ronda " & _
                             "  AND Jug.Par_Id = " & GintPartidaActiva & _
                             "  AND Jug.Col_Id = " & vecColoresActivos(i)
                    intMisionSorteada = CInt(EjecutarConsultaValor(strSQL))
                End If
            End If
            
            vecMisionesAsignadas(i) = intMisionSorteada
            'Guarda en la BD la misión asignada
            strSQL = "UPDATE Jugadores SET Mis_Id = " & intMisionSorteada & " " & _
                     "WHERE Col_Id = " & vecColoresActivos(i) & _
                     "  AND Par_Id = " & GintPartidaActiva
            EjecutarComando strSQL
            
        End If
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "sRepartirMisiones", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sEstablecerRonda()
    On Error GoTo ErrorHandle
    Dim vecColoresActivos() As String
    Dim vecPosicionesRonda() As Integer
    Dim strSQL As String
    Dim i As Integer
    Dim j As Integer
    Dim intJugadorSorteado As Integer
    Dim blnJugadorYaSorteado As Boolean
    
    ' Inicializa el generador de números aleatorios.
    Randomize
    
    'Obtiene los colores activos
    strSQL = "SELECT Col_Id FROM Jugadores WHERE Par_Id = " & GintPartidaActiva
    RecordsetAVector EjecutarConsulta(strSQL), 0, vecColoresActivos
    
    ReDim vecPosicionesRonda(LBound(vecColoresActivos) To UBound(vecColoresActivos))
    
    'Inicializa el vector posicion
    For i = LBound(vecPosicionesRonda) To UBound(vecPosicionesRonda)
        vecPosicionesRonda(i) = -1
    Next i
    
    'Sortea los jugadores (indice del vector jugadores) para cada posicion en la ronda
    'Lo hace asi para que quede ordenado por posicion
    For i = LBound(vecPosicionesRonda) To UBound(vecPosicionesRonda)
        intJugadorSorteado = Aleatorio(LBound(vecColoresActivos), UBound(vecColoresActivos)) 'Sortea el jugador
        blnJugadorYaSorteado = False
        'Busca en el vector de ronda si el jugador ya fue sorteado
        For j = LBound(vecPosicionesRonda) To UBound(vecPosicionesRonda)
            If vecPosicionesRonda(j) = intJugadorSorteado Then
                blnJugadorYaSorteado = True
            End If
        Next j
        If blnJugadorYaSorteado Then
            'Volver a sortear
            i = i - 1
        Else
            'Asigno la posición en la ronda
            vecPosicionesRonda(i) = intJugadorSorteado
        End If
    Next i
    
    'Actualiza la BD
    For i = LBound(vecPosicionesRonda) To UBound(vecPosicionesRonda)
        If i < UBound(vecPosicionesRonda) Then
            strSQL = "UPDATE Jugadores SET Jug_Prox_Ronda = " & vecColoresActivos(vecPosicionesRonda(i + 1)) & " " & _
                     "WHERE Col_Id = " & vecColoresActivos(vecPosicionesRonda(i)) & " " & _
                     "  AND Par_Id = " & GintPartidaActiva
        Else
            'Si es el último de la ronda, el siguiente jugador es el primero del vector
            strSQL = "UPDATE Jugadores SET Jug_Prox_Ronda = " & vecColoresActivos(vecPosicionesRonda(LBound(vecPosicionesRonda))) & " " & _
                     "WHERE Col_Id = " & vecColoresActivos(vecPosicionesRonda(i)) & " " & _
                     "  AND Par_Id = " & GintPartidaActiva
        End If
        EjecutarComando strSQL
    Next i
    
    'Actualiza en la BD al primero de la ronda
    strSQL = "UPDATE Partidas SET Par_Ronda_Primero = " & vecColoresActivos(vecPosicionesRonda(LBound(vecPosicionesRonda))) & " " & _
             "WHERE Par_Id = " & GintPartidaActiva
    EjecutarComando strSQL
    
    'Pone al primero de la ronda como jugador activo, en estado Agregando/atacando
    strSQL = "UPDATE Partidas SET Par_Activo_Ronda = " & vecColoresActivos(vecPosicionesRonda(LBound(vecPosicionesRonda))) & ", " & _
             "Par_Activo_Estado = " & enuEstadoActivo.estAgregando & ", " & _
             "Par_Activo_Conquistas = 0, " & _
             "Par_Ronda_Nro = 1, " & _
             "Par_Ronda_Tipo = " & enuTipoRonda.trInicio & " " & _
             "WHERE Par_Id = " & CStr(GintPartidaActiva)
    EjecutarComando strSQL
    
    'Inicializa el estado del jugador activo
    GEstadoActivo = estAgregando
    '### Borrar
    frmServer.lblEstadoJugadorActivo.Caption = vecEstadosActivo(GEstadoActivo)

    
    Exit Sub
ErrorHandle:
    ReportErr "sEstablecerRonda", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sActualizarMapa(Optional intDestino As Integer)
    'Envía a el/los cliente/s todo el mapa pais por pais
    On Error GoTo ErrorHandle
    
    Dim strSQL As String
    Dim rsPaises As Recordset
    
    strSQL = "SELECT Col_Id, Pas_Id, Tro_Cantidad, Tro_Fijos FROM Tropas WHERE Par_Id = " & GintPartidaActiva & ""
    Set rsPaises = EjecutarConsulta(strSQL)
    
    While Not rsPaises.EOF
        EnviarMensaje ArmarMensajeParam(msgPais, rsPaises!Pas_Id, rsPaises!Col_Id, rsPaises!Tro_Cantidad, CStr(enuOrigenMsgPais.orRepartoInicial), CStr(rsPaises!Tro_Fijos)), intDestino
        rsPaises.MoveNext
    Wend
    
    rsPaises.Close
    Set rsPaises = Nothing
    
    Exit Sub
ErrorHandle:
    ReportErr "sActualizarMapa", "mdlServidor", Err.Description, Err.Number, Err.Source
    Set rsPaises = Nothing
End Sub

Public Sub sEnviarMision(intColor As Integer)
    'Comunica la misiones asignada al jugador pasado por parametro
    On Error GoTo ErrorHandle
    Dim rsMisiones As Recordset
    Dim rsObjetivos As Recordset
    Dim strSQL As String
    Dim vecMensaje() As String
    
    strSQL = "SELECT J.Col_Id, M.Mis_Id, M.Mis_Desc FROM Jugadores J, Misiones M " & _
             "WHERE J.Mis_Id = M.Mis_Id " & _
             "  AND J.Par_Id = " & GintPartidaActiva & _
             "  AND J.Col_Id = " & intColor
    Set rsMisiones = EjecutarConsulta(strSQL)
    
    'Primera posicion: Descripción fija de la misión (en español) (por compatibilidad)
    'Segunda posicion: Codigo de la descripción en el archivo de recursos
    'Resto: Objetivos para el jugador virtual
    While Not rsMisiones.EOF
        ReDim vecMensaje(0 To 1)
        vecMensaje(0) = Trim(rsMisiones!Mis_Desc)
        vecMensaje(1) = CStr(enuIndiceArchivoRecurso.pmsMisiones + rsMisiones!Mis_Id + 1)
        'Busca los objetivos asociados a esa mision
        strSQL = "SELECT Con_Id, Col_Id, Mob_Cant_Paises, Mob_Limitrofes FROM Misiones_Objetivos " & _
                 "WHERE Mis_Id = " & rsMisiones!Mis_Id
        Set rsObjetivos = EjecutarConsulta(strSQL)
        While Not rsObjetivos.EOF
            'Arma el vector con el mensaje a enviar
            ReDim Preserve vecMensaje(UBound(vecMensaje) + 4)
            'Continente
            If IsNull(rsObjetivos.Fields(0).Value) Then
                vecMensaje(UBound(vecMensaje) - 3) = 0
            Else
                vecMensaje(UBound(vecMensaje) - 3) = rsObjetivos.Fields(0).Value
            End If
            
            'Color
            If IsNull(rsObjetivos.Fields(1).Value) Then
                vecMensaje(UBound(vecMensaje) - 2) = 0
            Else
                vecMensaje(UBound(vecMensaje) - 2) = rsObjetivos.Fields(1).Value
            End If
            
            'Cantidad de paises
            If IsNull(rsObjetivos.Fields(2).Value) Then
                vecMensaje(UBound(vecMensaje) - 1) = 0
            Else
                vecMensaje(UBound(vecMensaje) - 1) = rsObjetivos.Fields(2).Value
            End If
            
            'Limitrofes
            If IsNull(rsObjetivos.Fields(3).Value) Then
                vecMensaje(UBound(vecMensaje) - 0) = "N"
            Else
                vecMensaje(UBound(vecMensaje) - 0) = rsObjetivos.Fields(3).Value
            End If
            
            rsObjetivos.MoveNext
            
        Wend
        
        EnviarMensaje ArmarMensaje(msgMisionAsignada, vecMensaje), ColorAIndice(rsMisiones!Col_Id)
        rsMisiones.MoveNext
    Wend
    
    rsMisiones.Close
    Set rsMisiones = Nothing
    
    rsObjetivos.Close
    Set rsObjetivos = Nothing
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarMision", "mdlServidor", Err.Description, Err.Number, Err.Source
    Set rsMisiones = Nothing
    Set rsObjetivos = Nothing
End Sub

Public Sub sEnviarMisiones()
    'Comunica las misiones repartidas a cada jugador
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    For i = 1 To UBound(GvecColoresSock)
        If GvecColoresSock(i) > 0 Then
            sEnviarMision IndiceAColor(i)
        End If
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarMision", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sActualizarRonda(Optional intDestino As Integer)
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim intPrimeroRonda As Integer
    Dim rsParesRonda As Recordset
    Dim intColorBuscado As Integer
    Dim vecRondaOrdenada() As String
    Dim i As Integer
    
    'Busca en la BD el primero de la ronda
    strSQL = "SELECT Par_Ronda_Primero FROM Partidas WHERE Par_Id = " & GintPartidaActiva
    intPrimeroRonda = EjecutarConsultaValor(strSQL)
    
    'Busca en la BD la secuencia de la ronda
    strSQL = "SELECT Col_Id, Jug_Prox_Ronda FROM Jugadores WHERE Par_Id = " & GintPartidaActiva
    Set rsParesRonda = EjecutarConsulta(strSQL)
    
    If Not rsParesRonda.EOF Then
        rsParesRonda.MoveLast
        ReDim vecRondaOrdenada(0 To rsParesRonda.RecordCount - 1)
        rsParesRonda.MoveFirst
    End If
    
    'Recorre el recordset buscando el próximo color para cada jugador
    intColorBuscado = intPrimeroRonda
    For i = 0 To UBound(vecRondaOrdenada)
        vecRondaOrdenada(i) = CStr(intColorBuscado)
        rsParesRonda.MoveFirst
        While Not rsParesRonda.EOF
            If rsParesRonda!Col_Id = intColorBuscado Then
                intColorBuscado = rsParesRonda!Jug_Prox_Ronda
                rsParesRonda.MoveLast
            End If
            rsParesRonda.MoveNext
        Wend
    Next i
    
    'broadcast
    EnviarMensaje ArmarMensaje(msgOrdenRonda, vecRondaOrdenada), intDestino
    
    rsParesRonda.Close
    Set rsParesRonda = Nothing
    
    Exit Sub
ErrorHandle:
    ReportErr "sActualizarRonda", "mdlServidor", Err.Description, Err.Number, Err.Source
    Set rsParesRonda = Nothing
End Sub

Public Sub sConfirmarInicioPartida(Optional intDestino As Integer)
    'Confirma a los clientes el inicio de la partida (cambia su estado a jugando)
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgAckInicioPartida), intDestino
    
    Exit Sub
ErrorHandle:
    ReportErr "sConfirmarInicioPartida", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sInformarInicioTurno(blnNoResincronizacion As Boolean, Optional intDestino As Integer)
    'Informa a los jugadores el turno actual (broadcast)
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim intColorActivoRonda As Integer
    Dim intTimerTurno As Integer
    Dim intTolerancia As Integer
    
    'Obtiene de la BD el jugador que tiene el turno
    strSQL = "SELECT Par_Activo_Ronda FROM Partidas WHERE Par_Id = " & CStr(GintPartidaActiva)
    intColorActivoRonda = EjecutarConsultaValor(strSQL)
    
    'Obtiene de la BD la tolerancia del turno
    intTolerancia = CInt(ValorOpcion(opTurnoTolerancia))
    
    If blnNoResincronizacion Then
        'Obtiene de la BD el Timer del turno
        intTimerTurno = CInt(ValorOpcion(opTurnoDuracion)) * 60
        
        '###Log
        GuardarLog "Ha comenzado el turno de " & GvecNombreJugadorColor(intColorActivoRonda)
        sEnviarLog mscInicioTurno, CstrTipoParametroResuelto & GvecNombreJugadorColor(intColorActivoRonda)
    Else
        'Toma el valor actual del timer (resincronizacion)
        If GintValorTimerActual = CintValorTimerInfinito Then
            'Si el timer es infinito no toma la tolerancia
            intTimerTurno = CintValorTimerInfinito
        Else
            'Sino resta la tolerancia
            intTimerTurno = GintValorTimerActual - intTolerancia
        End If
    End If
    
    EnviarMensaje ArmarMensajeParam(msgInicioTurno, intColorActivoRonda, intTimerTurno, CInt(Not blnNoResincronizacion)), intDestino
    
    'Activa el timer del servidor
    If blnNoResincronizacion Then
        ActivarTimer intTimerTurno + intTolerancia
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "sInformarInicioTurno", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sTropasDisponibles()
    'Calcula las tropas disponibles para cada jugador
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim intTipoRonda As enuTipoRonda
    Dim intCantTropas As Integer
    Dim intNroRonda As Integer
    Dim vecColoresActivos() As String
    Dim i As Integer
    Dim j As Integer
    Dim intCantPaises As Integer
    Dim vecContinentes() As String
    Dim rsContinentes As Recordset
    Dim opcContinente As enuOpciones
    Dim intCantTropasContinente As Integer
    
    'Toma de la BD el tipo de la ronda
    strSQL = "SELECT Par_Ronda_Tipo FROM Partidas WHERE Par_Id = " & GintPartidaActiva
    intTipoRonda = EjecutarConsultaValor(strSQL)
    
    'Toma de la BD el número de la ronda
    strSQL = "SELECT Par_Ronda_Nro FROM Partidas WHERE Par_Id = " & GintPartidaActiva
    intNroRonda = EjecutarConsultaValor(strSQL)
    
    Select Case intTipoRonda
        Case enuTipoRonda.trInicio:
            'Primera y segunda ronda de la partida
            
            'Toma de la BD la cantidad de tropas a poner
            intCantTropas = CInt(ValorOpcion(IIf(intNroRonda = 1, enuOpciones.opRondaTropas1ra, enuOpciones.opRondaTropas2da)))
            
            'Guarda en la BD las tropas disponibles para todos los jugadores
            strSQL = "UPDATE Tropas_Disponibles SET Tdi_Cantidad = " & intCantTropas & _
                     " WHERE Par_Id = " & GintPartidaActiva & _
                     "  AND Con_Id IS NULL"
            EjecutarComando strSQL
        
        Case enuTipoRonda.trAccion
            'Pone en 0 las tropas disponibles (pierden las no usadas)
            strSQL = "UPDATE Tropas_Disponibles " & _
                     "SET Tdi_Cantidad = 0 " & _
                     "WHERE Par_Id = " & GintPartidaActiva
            EjecutarComando strSQL
        
        Case enuTipoRonda.trRecuento:
            
            'Obtiene los colores activos
            strSQL = "SELECT Col_Id FROM Jugadores WHERE Par_Id = " & GintPartidaActiva
            RecordsetAVector EjecutarConsulta(strSQL), 0, vecColoresActivos
            
            'Por cada jugador...
            For i = LBound(vecColoresActivos) To UBound(vecColoresActivos)
                'Cuenta la cantidad de paises
                strSQL = "SELECT COUNT(*) FROM Tropas " & _
                         "WHERE Col_Id = " & vecColoresActivos(i) & _
                         "  AND Par_Id = " & GintPartidaActiva
                intCantPaises = EjecutarConsultaValor(strSQL)
                
                intCantTropas = Int(intCantPaises / 2)
                
                'La cantidad de tropas disponibles no puede ser menor a tres
                If intCantTropas < 3 Then
                    intCantTropas = 3
                End If
                
                'Guarda en la BD las tropas disponibles libres
                strSQL = "UPDATE Tropas_Disponibles SET Tdi_Cantidad = " & intCantTropas & _
                         " WHERE Par_Id = " & GintPartidaActiva & _
                         "  AND Col_Id = " & vecColoresActivos(i) & _
                         "  AND Con_Id IS NULL"
                EjecutarComando strSQL
                
                'Identifica los continentes conquistados
                strSQL = "SELECT V1.Con_Id " & _
                         "FROM Vw_Paises_x_Continente_x_Jugador V1, Vw_Paises_x_Continente V2 " & _
                         "WHERE V1.Con_Id = V2.Con_Id " & _
                         "  AND V1.Paises = V2.Paises " & _
                         "  AND V1.Par_Id = " & GintPartidaActiva & _
                         "  AND V1.Col_id = " & vecColoresActivos(i)
                Set rsContinentes = EjecutarConsulta(strSQL)
                
                If Not rsContinentes.EOF Then
                    
                    RecordsetAVector rsContinentes, 0, vecContinentes
                    
                    'Por cada continente conquistado...
                    For j = LBound(vecContinentes) To UBound(vecContinentes)
                        Select Case vecContinentes(j)
                            Case 1
                                opcContinente = opBonusAfrica
                            Case 2
                                opcContinente = opBonusANorte
                            Case 3
                                opcContinente = opBonusASur
                            Case 4
                                opcContinente = opBonusAsia
                            Case 5
                                opcContinente = opBonusEuropa
                            Case 6
                                opcContinente = opBonusOceania
                        End Select
                        
                        'Busca el bonus del continente
                        intCantTropasContinente = CInt(ValorOpcion(opcContinente))
                        
                        'Las tropas del bonus solo se pueden poner en el continente
                        strSQL = "UPDATE Tropas_Disponibles SET Tdi_Cantidad = " & intCantTropasContinente & _
                                 " WHERE Par_Id = " & GintPartidaActiva & _
                                 "  AND Con_Id = " & vecContinentes(j) & _
                                 "  AND Col_Id = " & vecColoresActivos(i)
                        EjecutarComando strSQL
                        
                    Next j
            
                End If
                
                rsContinentes.Close
                Set rsContinentes = Nothing
            
            Next i
            
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "sTropasDisponibles", "mdlServidor", Err.Description, Err.Number, Err.Source
    Set rsContinentes = Nothing
End Sub

Public Sub sInformarTropasDisponibles(intColor As Integer, Optional intDestino As Integer)
    'broadcast
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim rsTDI As Recordset
    Dim i As Integer
    Dim vecMensaje(0 To 7) As String 'Jugador + Libres + 6 continentes
    
    vecMensaje(0) = CStr(intColor)
    
    'Toma de la BD las tropas disponibles
    strSQL = "SELECT Con_Id, Tdi_Cantidad FROM Tropas_Disponibles " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Col_Id = " & intColor & _
             " ORDER BY Con_Id"
    Set rsTDI = EjecutarConsulta(strSQL)
    
    i = 1
    'Guarda las cantidades de tropas disponibles por continente
    'para enviar en el mensaje
    While Not rsTDI.EOF
        vecMensaje(i) = rsTDI!Tdi_Cantidad
        i = i + 1
        rsTDI.MoveNext
    Wend
    
    'broadcast
    EnviarMensaje ArmarMensaje(msgTropasDisponibles, vecMensaje), intDestino
    
    rsTDI.Close
    Set rsTDI = Nothing
    
    Exit Sub
ErrorHandle:
    ReportErr "sInformarTropasDisponibles", "mdlServidor", Err.Description, Err.Number, Err.Source
    Set rsTDI = Nothing
End Sub

Public Sub sInformarTropasDisponiblesTodos(Optional intDestino As Integer)
    'LLama a sInformarTropasDisponibles para todos los jugadores conectados
    On Error GoTo ErrorHandle
    Dim vecColoresActivos() As String
    Dim i As Integer
    Dim strSQL As String
    
    'Obtiene los colores activos
    strSQL = "SELECT Col_Id FROM Jugadores WHERE Par_Id = " & GintPartidaActiva
    RecordsetAVector EjecutarConsulta(strSQL), 0, vecColoresActivos
    
    For i = LBound(vecColoresActivos) To UBound(vecColoresActivos)
        sInformarTropasDisponibles CInt(vecColoresActivos(i)), intDestino
    Next i

    Exit Sub
ErrorHandle:
    ReportErr "sInformarTropasDisponiblesTodos", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sFinTurno(blnExpiroTimer As Boolean, Optional blnTurnoSalteado As Boolean = False)
    On Error GoTo ErrorHandle
    Dim intProximoJugador As Integer
    Dim intPrimerJugador As Integer
    Dim intJugadorActivo As Integer
    Dim intTipoRonda As enuTipoRonda
    Dim strMisionGanador As String
    Dim intMisionId As Integer
    Dim blnEsObjetivoComun As Boolean
    Dim intCantPaisesObjetivoComun As Integer
    Dim strSQL As String
    Dim intCantPaisesProxJugador As Integer
    
    strSQL = "SELECT Par_Activo_Ronda FROM Partidas WHERE Par_Id = " & GintPartidaActiva
    intJugadorActivo = CInt(EjecutarConsultaValor(strSQL))
    
    'Si no expiró el timer y no se salteo el turno realiza la verificación
    If Not blnExpiroTimer And Not blnTurnoSalteado Then
        'Verifica que el jugador que informó el fin de turno sea el activo
        If Not EsElJugadorActivo(IndiceAColor(GintOrigenMensaje)) Then
            sEnviarError "Acción incorrecta. Imposible finalizar el turno dado que el mismo está en posesión de otro jugador. " & _
                         "Intente resincronizarse.", errNoFinTurnoNoTurno, GintOrigenMensaje
            Exit Sub
        End If
        
    End If
        
    'Desactiva el timer del servidor
    DesactivarTimer
    
    'Toma de la BD el tipo de ronda
    intTipoRonda = GetTipoDeRondaActiva
    If intTipoRonda = trAccion Then
        'Al finalizar un turno de accion pone en cero la cantidad de tropas fijas del jugador
        If intTipoRonda = trAccion Then
            strSQL = "UPDATE Tropas SET Tro_Fijos = 0 " & _
                     "WHERE Par_Id = " & GintPartidaActiva & _
                     "  AND Col_Id = " & intJugadorActivo
            EjecutarComando strSQL
        End If
    
        'Si es ronda de accion verifica si se completo la Mision
        If CumplioMision(intJugadorActivo, blnEsObjetivoComun) Then
            'Informa a los clientes la misión cumplida
            'Toma de la BD la descripción de la misión ganadora
            If blnEsObjetivoComun Then
                intCantPaisesObjetivoComun = ValorOpcion(opMisionObjetivoComun)
                strMisionGanador = "Conquistar " & intCantPaisesObjetivoComun & " países correspondientes al objetivo común."
                intMisionId = -1
            Else
                'Puede ser que el jugador tenga una Mision asignada pero que
                'esten jugando a Conquistar el Mundo porque se cambio durante el juego.
                'Por este motivo, reviso la opcion.
                
                If ValorOpcion(opMisionTipo) = "C" Then
                    'Si se juga a Conquistar el Mundo
                    strSQL = "SELECT Mis.Mis_Desc " & _
                             "FROM Misiones Mis " & _
                             "WHERE Mis.Mis_Id = 0 "
                    strMisionGanador = EjecutarConsultaValor(strSQL)
                    intMisionId = 0
                Else
                    'Si se juga por Misiones
                    strSQL = "SELECT Mis.Mis_Desc " & _
                             "FROM Misiones Mis, Jugadores Jug " & _
                             "WHERE Mis.Mis_Id = Jug.Mis_Id " & _
                             "  AND Jug.Par_Id = " & GintPartidaActiva & _
                             "  AND Jug.Col_Id = " & intJugadorActivo
                    strMisionGanador = EjecutarConsultaValor(strSQL)
                
                    strSQL = "SELECT Mis.Mis_Id " & _
                             "FROM Misiones Mis, Jugadores Jug " & _
                             "WHERE Mis.Mis_Id = Jug.Mis_Id " & _
                             "  AND Jug.Par_Id = " & GintPartidaActiva & _
                             "  AND Jug.Col_Id = " & intJugadorActivo
                    intMisionId = CInt(EjecutarConsultaValor(strSQL))
                End If
            End If
            
            'Guarda en la BD el jugador ganador
            strSQL = "UPDATE Partidas SET Par_Jug_Ganador = " & intJugadorActivo & _
                     " WHERE Par_Id = " & GintPartidaActiva
            EjecutarComando strSQL
            
            'Envia el mensaje (broadcast)
            EnviarMensaje ArmarMensajeParam(msgMisionCumplida, CStr(intJugadorActivo), strMisionGanador, enuIndiceArchivoRecurso.pmsMisiones + intMisionId + 1)
            
            '###Log
            GuardarLog GvecNombreJugadorColor(intJugadorActivo) & " ha logrado cumplir su misiòn."
            sEnviarLog mscMisionCumplida, CstrTipoParametroResuelto & GvecNombreJugadorColor(intJugadorActivo)
            
            Exit Sub
        End If
            
        'Resetea la cantidad de conquistas del jugador activo
        strSQL = "UPDATE Partidas SET Par_Activo_Conquistas=0 " & _
                 "WHERE Par_Id = " & GintPartidaActiva
        EjecutarComando strSQL
    End If
    
    
    
    'Toma de la BD el próximo jugador para saber si es el primero (fin ronda)
    strSQL = "SELECT Jug_Prox_Ronda FROM Jugadores " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Col_Id = " & intJugadorActivo
    intProximoJugador = EjecutarConsultaValor(strSQL)
    
    'Toma de la BD el primer jugador de la ronda
    strSQL = "SELECT Par_Ronda_Primero FROM Partidas WHERE Par_Id = " & GintPartidaActiva
    intPrimerJugador = EjecutarConsultaValor(strSQL)
    
    'Verifica si terminó la ronda
    If intPrimerJugador = intProximoJugador Then
        'Inicia una nueva ronda
        sIniciarRonda
            
        'Pisa el valor del proximo jugador con el primero de la nueva ronda
        strSQL = "SELECT Par_Ronda_Primero FROM Partidas WHERE Par_Id = " & GintPartidaActiva
        intProximoJugador = EjecutarConsultaValor(strSQL)
    
        'Toma de la BD el tipo de ronda
        intTipoRonda = GetTipoDeRondaActiva
    
    End If
    
    'Ack de fin de turno, si no se saltea el turno
    If Not blnTurnoSalteado Then
        sConfirmarFinTurno blnExpiroTimer
    End If
    
    'Actualiza la BD para el nuevo turno
    
    strSQL = "UPDATE Partidas SET " & _
             "Par_Activo_Ronda = " & intProximoJugador & ", " & _
             "Par_Activo_Conquistas = 0 " & _
             " WHERE Par_Id = " & GintPartidaActiva
    '###
'''    strSQL = "UPDATE Partidas SET " & _
'''             "Par_Activo_Ronda = " & intProximoJugador & ", " & _
'''             "Par_Activo_Estado = " & enuEstadoActivo.estAgregandoAtacando & ", " & _
'''             "Par_Activo_Conquistas = 0 " & _
'''             " WHERE Par_Id = " & GintPartidaActiva
    EjecutarComando strSQL
    
    'Actualiza el estado del jugador activo
    If intTipoRonda = trAccion Then
        ActualizarEstadoServidor eveFinTurnoAccion
    Else
        ActualizarEstadoServidor eveFinTurnoRecuento
    End If
    
    'Informa incio de nuevo turno
    'Antes de informarle el comienzo del turno,
    'chequea que el jugador tenga al menos un pais.
    'Obtiene de la BD la cant. de paises que posee el prox jugador
    strSQL = "SELECT COUNT(*) FROM Tropas " & _
             " WHERE Par_Id = " & CStr(GintPartidaActiva) & _
             "   AND Col_Id = " & intProximoJugador
    intCantPaisesProxJugador = EjecutarConsultaValor(strSQL)
    If intCantPaisesProxJugador > 0 Then
        'Si tiene al menos un pais, le pasa el turno
        sInformarInicioTurno True
    Else
        'Salte al jugador que no tiene paises
        sFinTurno False, True
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "sFinTurno", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sConfirmarFinTurno(blnExpiroTimer As Boolean)
    On Error GoTo ErrorHandle
    Dim intJugadorActivo As Integer
    Dim strSQL As String
    
    'Desactiva el timer
    DesactivarTimer
    
    'Toma de la BD el jugador activo
    strSQL = "SELECT Par_Activo_Ronda FROM Partidas WHERE Par_Id = " & GintPartidaActiva
    intJugadorActivo = EjecutarConsultaValor(strSQL)
    
    'Informa al cliente que finalizó su turno
    EnviarMensaje ArmarMensajeParam(msgAckFinTurno, CStr(CInt(blnExpiroTimer))), ColorAIndice(intJugadorActivo)
    
    Exit Sub
ErrorHandle:
    ReportErr "sConfirmarFinTurno", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub ActivarTimer(intSegundos As Integer)
    On Error GoTo ErrorHandle
    
    GintValorTimerActual = intSegundos
    GintValorTimerTotal = intSegundos
    
    'Activa el timer
    frmServer.tmrTurno.Interval = 1000
    
    'Marca el inicio del turno
    GsngInicioTimerTurno = Timer
    
    Exit Sub
ErrorHandle:
    ReportErr "ActivarTimer", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub DesactivarTimer()
    On Error GoTo ErrorHandle
    
    frmServer.tmrTurno.Interval = 0
    
    Exit Sub
ErrorHandle:
    ReportErr "DesactivarTimer", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sIniciarRonda()
    'Suma 1 al número de ronda
    'Calcula el tipo de la ronda
    'Cambia el primero de la ronda si corresponde
    'Informa la nueva conformación de la ronda
    'Si la ronda es de recuento llama a sTropasDisponibles e sInformarTropasDisponibles
    'Limpia la cantidad de tropas fijas de los clientes
    On Error GoTo ErrorHandle
    Dim intNroRonda As Integer
    Dim intTipoRonda As enuTipoRonda
    Dim chrOpcionPrimero As String * 1
    Dim strSQL As String
    
    'Toma de la BD el nro de ronda
    strSQL = "SELECT Par_Ronda_Nro FROM Partidas WHERE Par_Id = " & GintPartidaActiva
    intNroRonda = EjecutarConsultaValor(strSQL)
    
    'Incrementa el número de ronda
    intNroRonda = intNroRonda + 1
    
    'Calcula el tipo de ronda
    If intNroRonda <= 2 Then
        intTipoRonda = trInicio
    Else
        If intNroRonda Mod 2 = 0 Then
            'si es par
            intTipoRonda = trRecuento
        Else
            'si es impar
            intTipoRonda = trAccion
        End If
    End If
    
    'Actualiza en la BD el nro y tipo de ronda
    strSQL = "UPDATE Partidas SET " & _
             "Par_Ronda_Nro = " & intNroRonda & ", " & _
             "Par_Ronda_Tipo = " & intTipoRonda & _
             " WHERE Par_Id = " & GintPartidaActiva
    EjecutarComando strSQL
    
    'Toma de la BD la opción para saber si el primero de la ronda es fijo o rotativo
    chrOpcionPrimero = ValorOpcion(opRondaTipo)
    
    'Cambia el primero de la ronda si corresponde (turno rotativo)
    If Trim(UCase(chrOpcionPrimero)) = "R" And intTipoRonda = trRecuento Then
        'Cambia el primero de la ronda
        strSQL = "UPDATE Partidas P, Jugadores J SET P.Par_Ronda_Primero = J.Jug_Prox_Ronda " & _
                 "WHERE P.Par_Ronda_Primero = J.Col_Id " & _
                 "  AND P.Par_Id = J.Par_Id " & _
                 "  AND P.Par_Id = " & GintPartidaActiva
        EjecutarComando strSQL
        
        'Informa a los clientes la nueva conformación de la ronda
        sActualizarRonda
    End If
    
    'Informa a los clientes el tipo de ronda
    sInformarTipoRonda
    
    'Si la ronda es de recuento llama a sTropasDisponibles e sInformarTropasDisponibles
    sTropasDisponibles
    sInformarTropasDisponiblesTodos
    
    Exit Sub
ErrorHandle:
    ReportErr "sIniciarRonda", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Function ValorOpcion(intOpcion As enuOpciones) As String
    'Dado un código de opción devuelve su valor
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim strValor As String
    
    strSQL = "SELECT Opc_Valor FROM Opciones " & _
             "WHERE Opc_Id = " & intOpcion & _
             "  AND Par_Id = " & GintPartidaActiva
    
    strValor = EjecutarConsultaValor(strSQL)
    
    If IsNull(strValor) Then
        ValorOpcion = ""
    Else
        ValorOpcion = strValor
    End If
    
    Exit Function
ErrorHandle:
    ReportErr "ValorOpcion", "mdlServidor", Err.Description, Err.Number, Err.Source
End Function

Public Sub sInformarTipoRonda(Optional intDestino As Integer)
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim intTipoRonda As enuTipoRonda
    
    'Obtiene de la BD el tipo de ronda y la envía a los clientes
    strSQL = "SELECT Par_Ronda_Tipo FROM Partidas WHERE Par_Id = " & GintPartidaActiva
    intTipoRonda = EjecutarConsultaValor(strSQL)
    
    'broadcast
    EnviarMensaje ArmarMensajeParam(msgTipoRonda, intTipoRonda), intDestino
    
    Exit Sub
ErrorHandle:
    ReportErr "sInformarTipoRonda", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sAgregarTropas(intPais As Integer, intCantidad As Integer)
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim intColorDuenio As Integer
    Dim intTropasDisponibles As Integer
    Dim intTropasDisponiblesContinente As Integer
    Dim intTropasDisponiblesLibres As Integer
    Dim intTropasAux As Integer
    Dim intTipoRonda As enuTipoRonda
    
    'Verifica que la acción provenga del activo
    If Not EsElJugadorActivo(IndiceAColor(GintOrigenMensaje)) Then
        sEnviarError "Acción incorrecta. Imposible agregar tropas, dado que no es su turno. " & _
                     "Intente resincronizarse.", errNoAgregarNoTurno, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que la ronda sea de recuento o inicial
    '### No es necesario con la matriz de estados
    intTipoRonda = GetTipoDeRondaActiva
    
    If intTipoRonda <> trRecuento And intTipoRonda <> trInicio Then
        sEnviarError "No es posible agregar tropas en la ronda de Acción.", errNoAgregarRondaAccion
        Exit Sub
    End If
    
    If Not ValidarProximoEstado(eveAgregarTropas) Then
        Exit Sub
    End If
    
    'Valida que el pais sea del jugador
    strSQL = "SELECT Col_Id FROM Tropas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Pas_Id = " & intPais
    intColorDuenio = EjecutarConsultaValor(strSQL)
    
    If intColorDuenio <> IndiceAColor(GintOrigenMensaje) Then
        sEnviarError "No es posible agregar tropas. El país seleccionado no le pertenece", errNoAgregarNoPais, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que no se supere la cantidad y el destino de tropas disponibles
    strSQL = "SELECT SUM(TDI.Tdi_Cantidad) " & _
             "FROM Tropas_Disponibles TDI, Paises PAS " & _
             "WHERE TDI.Con_Id Is Null " & _
             "  AND PAS.Pas_Id = " & intPais & _
             "  AND TDI.Col_Id = " & intColorDuenio & _
             "  AND TDI.Par_Id = " & GintPartidaActiva
    intTropasDisponiblesLibres = EjecutarConsultaValor(strSQL)
    
    strSQL = "SELECT TDI.Tdi_Cantidad " & _
             "FROM Tropas_Disponibles TDI, Paises PAS " & _
             "WHERE PAS.Con_Id = TDI.Con_Id " & _
             "  AND PAS.Pas_Id = " & intPais & _
             "  AND TDI.Col_Id = " & intColorDuenio & _
             "  AND TDI.Par_Id = " & GintPartidaActiva
    intTropasDisponiblesContinente = EjecutarConsultaValor(strSQL)
    
    intTropasDisponibles = intTropasDisponiblesContinente + intTropasDisponiblesLibres
    
    If intTropasDisponibles < intCantidad Then
        sEnviarError "No es posible agregar tropas. Tropas disponibles insuficientes. " & _
                     "Tiene " & intTropasDisponiblesLibres & " tropas disponibles libres y " & _
                     intTropasDisponiblesContinente & " tropas disponibles para el continente seleccionado. " & _
                     "Consulte la ventana de Detalle de Tropas Disponibles.", errNoAgregarNoTropas, GintOrigenMensaje
        Exit Sub
    End If

    'Valida que la cantidad de tropas a agregar sea mayor a cero
    If intCantidad <= 0 Then
        sEnviarError "No es posible agregar tropas. La cantidad de tropas a agregar no puede ser cero.", errNoAgregarNoCantidad, GintOrigenMensaje
        Exit Sub
    End If
    
    'Actualiza la BD
    'Actualiza las tropas disponibles
    intTropasAux = IIf(intCantidad > intTropasDisponiblesContinente, intTropasDisponiblesContinente, intCantidad)
    
    If intTropasDisponiblesContinente > 0 Then
        strSQL = "UPDATE Tropas_Disponibles TDI, Paises PAS " & _
                 "SET TDI.Tdi_Cantidad = TDI.Tdi_Cantidad - " & intTropasAux & " " & _
                 "WHERE PAS.Con_Id = TDI.Con_Id " & _
                 "  AND PAS.Pas_Id = " & intPais & _
                 "  AND TDI.Col_Id = " & intColorDuenio & _
                 "  AND TDI.Par_Id = " & GintPartidaActiva
        EjecutarComando strSQL
    End If
    
    If intCantidad > intTropasAux Then
        strSQL = "UPDATE Tropas_Disponibles TDI " & _
                 "SET TDI.Tdi_Cantidad = TDI.Tdi_Cantidad - " & intCantidad - intTropasAux & " " & _
                 "WHERE TDI.Con_Id IS NULL" & _
                 "  AND TDI.Col_Id = " & intColorDuenio & _
                 "  AND TDI.Par_Id = " & GintPartidaActiva
        EjecutarComando strSQL
    End If
    
    'Actualiza las tropas del pais
    strSQL = "UPDATE Tropas " & _
             "SET Tro_Cantidad = Tro_Cantidad + " & intCantidad & _
             " WHERE Pas_Id = " & intPais & _
             "   AND Par_Id = " & GintPartidaActiva
    EjecutarComando strSQL
    
    'Envia mensaje pais (broadcast)
    sActualizarPais intPais, orAgregado
    
    'Envia mensaje tropas disponibles (broadcast)
    sInformarTropasDisponibles intColorDuenio
    
    'Envia el Ack (Unicast) (Para el jugador virtual)
    EnviarMensaje ArmarMensajeParam(msgAckAgregarTropas, 0)
    
    'Si todo salió bien actualiza el estado
    ActualizarEstadoServidor eveAgregarTropas
    
    '###Log
    If intCantidad = 1 Then
        GuardarLog GvecNombreJugadorColor(intColorDuenio) & " ha agregado 1 tropa en " & GvecPaises(intPais) & "."
        sEnviarLog mscAgregado1Tropa, CstrTipoParametroResuelto & GvecNombreJugadorColor(intColorDuenio), CstrTipoParametroRecurso & CStr(intPais + enuIndiceArchivoRecurso.pmsPaises)
    Else
        GuardarLog GvecNombreJugadorColor(intColorDuenio) & " ha agregado " & intCantidad & " tropas en " & GvecPaises(intPais) & "."
        sEnviarLog mscAgregadoTropas, CstrTipoParametroResuelto & GvecNombreJugadorColor(intColorDuenio), CstrTipoParametroResuelto & CStr(intCantidad), CstrTipoParametroRecurso & CStr(intPais + enuIndiceArchivoRecurso.pmsPaises)
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "sAgregarTropas", "mdlInterfase", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sActualizarPais(intPais As Integer, intOrigen As enuOrigenMsgPais, Optional intDestino As Integer)
    'Envía a el/los cliente/s un mensaje pais
    On Error GoTo ErrorHandle
    
    Dim strSQL As String
    Dim rsPais As Recordset
    Dim intCantidad As Integer
    
    strSQL = "SELECT Col_Id, Pas_Id, Tro_Cantidad, Tro_Fijos FROM Tropas " & _
             "WHERE Pas_Id = " & intPais & _
             "  AND Par_Id = " & GintPartidaActiva & ""
    Set rsPais = EjecutarConsulta(strSQL)
    
    If Not rsPais.EOF Then
        EnviarMensaje ArmarMensajeParam(msgPais, CStr(rsPais!Pas_Id), CStr(rsPais!Col_Id), CStr(rsPais!Tro_Cantidad), CStr(intOrigen), CStr(rsPais!Tro_Fijos)), intDestino
    End If
    
    rsPais.Close
    Set rsPais = Nothing
    
    Exit Sub
ErrorHandle:
    ReportErr "sActualizarPais", "mdlServidor", Err.Description, Err.Number, Err.Source
    Set rsPais = Nothing
End Sub

Public Sub sAtacar(intPaisDesde As Integer, intPaisHasta As Integer)
    On Error GoTo ErrorHandle
    Dim intTipoRonda As enuTipoRonda
    Dim intColorDesde As Integer
    Dim intColorHasta As Integer
    Dim intCantTropasDesde As Integer
    Dim intCantTropasHasta As Integer
    Dim intCantDadosDesde As Integer
    Dim intCantDadosHasta As Integer
    Dim vecDadosDesde() As Integer
    Dim vecDadosHasta() As Integer
    Dim intTropasPerdidasDesde As Integer
    Dim intTropasPerdidasHasta As Integer
    Dim intCantLuchas As Integer
    Dim blnConquista As Boolean
    Dim vecMensaje(0 To 11) As String
    Dim intTropasResultadoDesde As Integer
    Dim intTropasResultadoHasta As Integer
    Dim intColorResultadoHasta As Integer
    Dim intCantPaisesVictima As Integer
    Dim strSQL As String
    Dim i As Integer
    Dim intCantTarjetasVictima As Integer
    Dim intCantTarjetasAsesino As Integer
    Dim intCantTarjetasHerencia As Integer
    Dim vecTarjetasPasadas() As String
    

    'Verifica que el jugador que informó el ataque sea el activo
    If Not EsElJugadorActivo(IndiceAColor(GintOrigenMensaje)) Then
        sEnviarError "Acción incorrecta. Imposible efectuar el ataque dado que no es su turno. " & _
                     "Intente resincronizarse.", errNoAtaqueNoTurno, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que la ronda sea de acción
    '### No es necesario con la matriz de estados
    intTipoRonda = GetTipoDeRondaActiva
    
    If intTipoRonda <> trAccion Then
        sEnviarError "No es posible efectuar un ataque en una ronda que no sea de Acción.", errNoAtaqueNoRondaAccion
        Exit Sub
    End If
    
    'Valida contra la matriz de estados
    If Not ValidarProximoEstado(eveAtaque) Then
        Exit Sub
    End If
    
    'Valida que el pais de Origen sea del jugador activo
    strSQL = "SELECT Col_Id FROM Tropas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Pas_Id = " & intPaisDesde
    intColorDesde = EjecutarConsultaValor(strSQL)
    
    If intColorDesde <> IndiceAColor(GintOrigenMensaje) Then
        sEnviarError "No es posible efectuar el ataque. El país de Origen no le pertenece.", errNoAtaqueNoOrigen, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que el pais Destino no sea del jugador activo
    strSQL = "SELECT Col_Id FROM Tropas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Pas_Id = " & intPaisHasta
    intColorHasta = EjecutarConsultaValor(strSQL)
    
    If intColorHasta = IndiceAColor(GintOrigenMensaje) Then
        sEnviarError "No es posible efectuar el ataque. El país de Destino no es enemigo.", errNoAtaqueNoDestino, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que el pais Desde tenga mas de una tropa
    strSQL = "SELECT Tro_Cantidad FROM Tropas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Pas_Id = " & intPaisDesde
    intCantTropasDesde = EjecutarConsultaValor(strSQL)
    
    If intCantTropasDesde <= 1 Then
        sEnviarError "No es posible efectuar el ataque. El país de Origen debe contener mas de una tropa.", errNoAtaqueNoTropas, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que los paises sean limitrofes
    strSQL = "SELECT COUNT(*) FROM Limites " & _
             "WHERE Pas_Id_Desde = " & intPaisDesde & _
             "  AND Pas_Id_Hasta = " & intPaisHasta
    
    If CInt(EjecutarConsultaValor(strSQL)) <= 0 Then
        sEnviarError "No es posible efectuar el ataque. Los países seleccionados no son limítrofes.", errNoAtaqueNoLimitrofes, GintOrigenMensaje
        Exit Sub
    End If
    
    'Toma de la BD la cantidad de tropas del pais Hasta
    strSQL = "SELECT Tro_Cantidad FROM Tropas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Pas_Id = " & intPaisHasta
    intCantTropasHasta = EjecutarConsultaValor(strSQL)
    
    'Calcula la cantidad de dados para Desde y Hasta
    intCantDadosDesde = intCantTropasDesde - 1
    If intCantDadosDesde > 3 Then
        intCantDadosDesde = 3
    End If
    
    intCantDadosHasta = intCantTropasHasta
    If intCantDadosHasta > 3 Then
        intCantDadosHasta = 3
    End If
    
    '###Log - Inicio Ataque
    GuardarLog GvecPaises(intPaisDesde) & "(" & intCantDadosDesde & ") ataca a " & _
               GvecPaises(intPaisHasta) & "(" & intCantDadosHasta & ")."
    sEnviarLog mscAtaque, CstrTipoParametroRecurso & CStr(intPaisDesde + enuIndiceArchivoRecurso.pmsPaises), _
                          CstrTipoParametroResuelto & CStr(intCantDadosDesde), _
                          CstrTipoParametroRecurso & CStr(intPaisHasta + enuIndiceArchivoRecurso.pmsPaises), _
                          CstrTipoParametroResuelto & CStr(intCantDadosHasta)
    
    ReDim vecDadosDesde(0 To intCantDadosDesde - 1)
    ReDim vecDadosHasta(0 To intCantDadosHasta - 1)
    
    'Tira los dados
    Randomize
    For i = LBound(vecDadosDesde) To UBound(vecDadosDesde)
        vecDadosDesde(i) = Dado
    Next i
    
    For i = LBound(vecDadosHasta) To UBound(vecDadosHasta)
        vecDadosHasta(i) = Dado
    Next i
    
    OrdenarVector vecDadosDesde
    OrdenarVector vecDadosHasta
    
    'Evalua los resultados
    intTropasPerdidasDesde = 0
    intTropasPerdidasHasta = 0
    
    If intCantDadosDesde >= intCantDadosHasta Then
        intCantLuchas = intCantDadosHasta
    Else
        intCantLuchas = intCantDadosDesde
    End If
    
    For i = LBound(vecDadosHasta) To intCantLuchas - 1
        If vecDadosDesde(i) > vecDadosHasta(i) Then
            'Pierde Hasta
            intTropasPerdidasHasta = intTropasPerdidasHasta + 1
        Else
            'Pierde Desde
            intTropasPerdidasDesde = intTropasPerdidasDesde + 1
        End If
    Next i
    
    'Evalua si hubo conquista
    If intCantTropasHasta <= intTropasPerdidasHasta Then
        'Hubo conquista
        blnConquista = True
        intTropasResultadoDesde = intCantTropasDesde - intTropasPerdidasDesde - 1
        intTropasResultadoHasta = 1
        intColorResultadoHasta = intColorDesde
    Else
        blnConquista = False
        intTropasResultadoDesde = intCantTropasDesde - intTropasPerdidasDesde
        intTropasResultadoHasta = intCantTropasHasta - intTropasPerdidasHasta
        intColorResultadoHasta = intColorHasta
    End If
    
       
    'Informa al cliente los resultados del ataque
    'Arma el mensaje
    For i = 0 To 2
        'Dados desde
        If i > UBound(vecDadosDesde) Then
            vecMensaje(i) = "0"
        Else
            vecMensaje(i) = CStr(vecDadosDesde(i))
        End If
        
        'Dados hasta
        If i > UBound(vecDadosHasta) Then
            vecMensaje(i + 3) = "0"
        Else
            vecMensaje(i + 3) = CStr(vecDadosHasta(i))
        End If
    Next i
    
    vecMensaje(6) = intPaisDesde
    vecMensaje(7) = intColorDesde
    vecMensaje(8) = intTropasResultadoDesde
    vecMensaje(9) = intPaisHasta
    vecMensaje(10) = intColorResultadoHasta
    vecMensaje(11) = intTropasResultadoHasta
    
    'Envia el mensaje a los clientes (broadcast)
    EnviarMensaje ArmarMensaje(msgAckAtaque, vecMensaje)
    
    
    '###Log - Resultado Ataque
    GuardarLog GvecPaises(intPaisDesde) & " perdió " & intTropasPerdidasDesde & " tropa/s y " & _
               GvecPaises(intPaisHasta) & " perdió " & intTropasPerdidasHasta & " tropa/s."
    sEnviarLog mscResultadoAtaque, CstrTipoParametroRecurso & CStr(intPaisDesde + enuIndiceArchivoRecurso.pmsPaises), _
                          CstrTipoParametroResuelto & CStr(intTropasPerdidasDesde), _
                          CstrTipoParametroRecurso & CStr(intPaisHasta + enuIndiceArchivoRecurso.pmsPaises), _
                          CstrTipoParametroResuelto & CStr(intTropasPerdidasHasta)
       
       
       
    'Actualiza la BD con los resultados del ataque
    If blnConquista Then
        
        'Actualiza la cantidad de conquistas del jugador activo
        strSQL = "UPDATE Partidas " & _
                 "SET Par_Activo_Conquistas = Par_Activo_Conquistas + 1 " & _
                 "WHERE Par_Id = " & GintPartidaActiva
        EjecutarComando strSQL
        
        'Verifica si es el último país del jugador atacado (para saber si fue eliminado)
        strSQL = "SELECT COUNT(*) - 1 FROM Tropas " & _
                 "WHERE Par_Id = " & GintPartidaActiva & _
                 "  AND Col_Id = " & intColorHasta
        intCantPaisesVictima = CInt(EjecutarConsultaValor(strSQL))
        
        '###Log - Conquista
        GuardarLog GvecNombreJugadorColor(intColorDesde) & " ha conquistado " & GvecPaises(intPaisHasta) & "."
        sEnviarLog mscConquista, CstrTipoParametroResuelto & GvecNombreJugadorColor(intColorDesde), _
                          CstrTipoParametroRecurso & CStr(intPaisHasta + enuIndiceArchivoRecurso.pmsPaises)
        GuardarLog "Se ha pasado una tropa de " & GvecPaises(intPaisDesde) & " a " & GvecPaises(intPaisHasta) & "."
        sEnviarLog mscMovimientoConquista1, CstrTipoParametroRecurso & CStr(intPaisDesde + enuIndiceArchivoRecurso.pmsPaises), _
                          CstrTipoParametroRecurso & CStr(intPaisHasta + enuIndiceArchivoRecurso.pmsPaises)
        
        If intCantPaisesVictima <= 0 Then
            'Si la victima fue eliminada por el jugador activo
            
            '###Log - Jugador eliminado
            GuardarLog GvecNombreJugadorColor(intColorHasta) & " ha sido eliminado por " & GvecNombreJugadorColor(intColorDesde) & "."
            sEnviarLog mscJugadorEliminado, CstrTipoParametroResuelto & GvecNombreJugadorColor(intColorHasta), _
                                            CstrTipoParametroResuelto & GvecNombreJugadorColor(intColorDesde)

            'Actualiza la BD
            strSQL = "UPDATE Jugadores SET Jug_Eliminado_Por = " & intColorDesde & _
                     " WHERE Par_Id = " & GintPartidaActiva & _
                     "   AND Col_Id = " & intColorHasta
            EjecutarComando strSQL
            
            'Toma de la BD la cantidad de tarjetas de la víctima.
            strSQL = "SELECT COUNT(*) FROM Jugadores_Tarjetas " & _
                     "WHERE Par_Id = " & GintPartidaActiva & _
                     "  AND Col_Id = " & intColorHasta
            intCantTarjetasVictima = EjecutarConsultaValor(strSQL)
            
            'Si la Víctima tenia Tarjetas, se las pasa al Asesino o las devuelve al mazo.
            If intCantTarjetasVictima > 0 Then
            
                'Toma de la BD la cantidad de tarjetas del asesino.
                strSQL = "SELECT COUNT(*) FROM Jugadores_Tarjetas " & _
                         "WHERE Par_Id = " & GintPartidaActiva & _
                         "  AND Col_Id = " & intColorDesde
                intCantTarjetasAsesino = EjecutarConsultaValor(strSQL)
                
                'Si el asesino tiene lugar, recibe todas las tarjetas posibles.
                If intCantTarjetasAsesino < 5 Then
                
                    'Calcula la cantidad de Tarjetas a Heredar
                    If (5 - intCantTarjetasAsesino) >= intCantTarjetasVictima Then
                        'Hereda todas las tarjetas de la Victima
                        intCantTarjetasHerencia = intCantTarjetasVictima
                    Else
                        'Hereda todas las que puede y las restantes van al mazo
                        intCantTarjetasHerencia = 5 - intCantTarjetasAsesino
                    End If
                    
                    'Antes de pasar las tarjetas guarda los Id de las mismas
                    strSQL = "SELECT TOP " & intCantTarjetasHerencia & " Tar_Id " & _
                             "FROM Jugadores_Tarjetas " & _
                             "WHERE Par_Id = " & GintPartidaActiva & _
                             "  AND Col_Id = " & intColorHasta
                    RecordsetAVector EjecutarConsulta(strSQL), 0, vecTarjetasPasadas
                    
                    'Se pasan todas las posibles (y se les quita la marca de "Cobrada") y las otras se vuelven al mazo
                    strSQL = "UPDATE Jugadores_Tarjetas " & _
                             "SET Col_Id = " & intColorDesde & ", " & _
                             "    Jut_Cobrada = 'N' " & _
                             "WHERE Par_Id = " & GintPartidaActiva & _
                             "  AND Col_Id = " & intColorHasta & _
                             "  AND Tar_Id IN(" & _
                                "SELECT TOP " & intCantTarjetasHerencia & " Tar_Id " & _
                                "FROM Jugadores_Tarjetas " & _
                                "WHERE Par_Id = " & GintPartidaActiva & _
                                "  AND Col_Id = " & intColorHasta & ")"
                    EjecutarComando strSQL
                    
                    '###Log - Herencia Tarjetas (todas las posibles)
                    GuardarLog "Se han pasado " & intCantTarjetasHerencia & " tarjeta/s de " & GvecNombreJugadorColor(intColorHasta) & " a " & GvecNombreJugadorColor(intColorDesde) & "."
                    sEnviarLog mscJugadorEliminadoTarjetasAsesino, CstrTipoParametroResuelto & CStr(intCantTarjetasHerencia), _
                                                                   CstrTipoParametroResuelto & GvecNombreJugadorColor(intColorHasta), _
                                                                   CstrTipoParametroResuelto & GvecNombreJugadorColor(intColorDesde)
                End If
                
                'Si sobran tarjetas de la victima, vuelven al mazo.
                If ((5 - intCantTarjetasAsesino) < intCantTarjetasVictima) Then
                
                    'Vuelven al mazo
                    strSQL = "DELETE FROM Jugadores_Tarjetas " & _
                             "WHERE Par_Id = " & GintPartidaActiva & _
                             "  AND Col_Id = " & intColorHasta
                    EjecutarComando strSQL
                
                    '###Log - Devolucion de Tarjetas al mazo
                    GuardarLog "Se han devuelto " & CStr(intCantTarjetasVictima - (5 - intCantTarjetasAsesino)) & " tarjeta/s de " & GvecNombreJugadorColor(intColorHasta) & " al mazo."
                    sEnviarLog mscJugadorEliminadoTarjetasMazo, CstrTipoParametroResuelto & CStr(intCantTarjetasVictima - (5 - intCantTarjetasAsesino)), _
                                                                 CstrTipoParametroResuelto & GvecNombreJugadorColor(intColorHasta)
                End If
                
                
                'Envia la cantidad de tarjetas de cada jugador (broadcast)
                sActualizarTarjetasTodos
                
                'Si el asesino recibió alguna tarjeta, se le envia el Detalle.
                If intCantTarjetasAsesino < 5 Then
                    'Detalle de las nuevas tarjetas del jugador asesino (unicast).
                    For i = 0 To UBound(vecTarjetasPasadas)
                        sEnviarTarjeta CInt(vecTarjetasPasadas(i))
                    Next
                    
                    'Envia el estado del turno del cliente
                    '(necesario porque al enviar las tarjetas al asesino, este pasa al estado
                    'tarjeta tomada o inconsistente)
                    sEnviarEstadoTurno ColorAIndice(intColorDesde)
                End If
            
            End If
        End If
    
    End If
    
    strSQL = "UPDATE Tropas SET " & _
             "Tro_Cantidad = " & intTropasResultadoDesde & _
             " WHERE Pas_Id = " & intPaisDesde & _
             "   AND Par_Id = " & GintPartidaActiva
    EjecutarComando strSQL
    
    strSQL = "UPDATE Tropas SET " & _
             "Col_Id = " & intColorResultadoHasta & ", " & _
             "Tro_Cantidad = " & intTropasResultadoHasta & _
             " WHERE Pas_Id = " & intPaisHasta & _
             "   AND Par_Id = " & GintPartidaActiva
    EjecutarComando strSQL
        
    'Si todo salio bien actualiza la matriz de estados
    ActualizarEstadoServidor eveAtaque
    
    Exit Sub
ErrorHandle:
    ReportErr "sAtacar", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Function EsElJugadorActivo(intColor As Integer) As Boolean
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim intJugadorActivo As Integer
    
    'Toma de la BD el jugador activo
    strSQL = "SELECT Par_Activo_Ronda FROM Partidas WHERE Par_Id = " & GintPartidaActiva
    intJugadorActivo = EjecutarConsultaValor(strSQL)
    
    'Verifica que el jugador sea el activo
    If intColor <> intJugadorActivo Then
        EsElJugadorActivo = False
    Else
        EsElJugadorActivo = True
    End If

    Exit Function
ErrorHandle:
    ReportErr "EsElJugadorActivo", "mdlServidor", Err.Description, Err.Number, Err.Source
End Function

Public Function GetTipoDeRondaActiva() As enuTipoRonda
    On Error GoTo ErrorHandle
    Dim strSQL As String
    
    strSQL = "SELECT Par_Ronda_Tipo FROM Partidas WHERE Par_Id = " & GintPartidaActiva
    GetTipoDeRondaActiva = CInt(EjecutarConsultaValor(strSQL))
    
    Exit Function
ErrorHandle:
    ReportErr "GetTipoDeRondaActiva", "mdlServidor", Err.Description, Err.Number, Err.Source
End Function

Public Sub OrdenarVector(ByRef vectorAordenar() As Integer, Optional blnAscendente As Boolean = True)
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim intAux As Integer
    
    For i = LBound(vectorAordenar) To UBound(vectorAordenar) - 1
        If vectorAordenar(i) < vectorAordenar(i + 1) Then
            intAux = vectorAordenar(i)
            vectorAordenar(i) = vectorAordenar(i + 1)
            vectorAordenar(i + 1) = intAux
            'Si hubo un cambio vuelve a empezar
            i = LBound(vectorAordenar) - 1
        End If
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "OrdenarVector", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sMover(intPaisDesde As Integer, intPaisHasta As Integer, intCantidadTropas As Integer, intTipoMovimiento As enuTipoMovimiento)
    On Error GoTo ErrorHandle
    
    Dim intTipoRonda As enuTipoRonda
    Dim intCantTropasDesde As Integer
    Dim intCantTropasDesdeLibres As Integer
    Dim intCantTropasHasta As Integer
    Dim intColorDesde As Integer
    Dim intColorHasta As Integer
    
    Dim intTropasResultadoDesde As Integer
    Dim intTropasResultadoHasta As Integer
    Dim strSQL As String
    Dim i As Integer
    
    'Verifica que el jugador que informó el movimiento sea el activo
    If Not EsElJugadorActivo(IndiceAColor(GintOrigenMensaje)) Then
        sEnviarError "Acción incorrecta. Imposible efectuar el movimiento dado que no es su turno. " & _
                     "Intente resincronizarse.", errNoMovimientoNoTurno, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que la ronda sea de acción
    '### No es necesario con la matriz de estados
    intTipoRonda = GetTipoDeRondaActiva
    
    If intTipoRonda <> trAccion Then
        sEnviarError "No es posible efectuar un movimiento en una ronda que no sea de Acción.", errNoMovimientoNoRondaAccion, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida contra la Matriz de Estados
    If Not ValidarProximoEstado(eveMovimiento) Then
        Exit Sub
    End If
    
    'Valida que el pais de Origen sea del jugador activo
    strSQL = "SELECT Col_Id FROM Tropas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Pas_Id = " & intPaisDesde
    intColorDesde = EjecutarConsultaValor(strSQL)
    
    If intColorDesde <> IndiceAColor(GintOrigenMensaje) Then
        sEnviarError "No es posible efectuar el movimiento. El país de Origen no le pertenece.", errNoMovimientoNoOrigen, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que el pais Destino sea tambien del jugador activo
    strSQL = "SELECT Col_Id FROM Tropas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Pas_Id = " & intPaisHasta
    intColorHasta = EjecutarConsultaValor(strSQL)
    
    If intColorHasta <> IndiceAColor(GintOrigenMensaje) Then
        sEnviarError "No es posible efectuar el movimiento. El país de Destino no le pertenece.", errNoMovimientoNoDestino, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que el pais Desde tenga mas tropas (libres) que las informadas
    strSQL = "SELECT Tro_Cantidad - Tro_Fijos FROM Tropas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Pas_Id = " & intPaisDesde
    intCantTropasDesdeLibres = CInt(EjecutarConsultaValor(strSQL))
    
    If intCantTropasDesdeLibres <= intCantidadTropas Then
        sEnviarError "No es posible efectuar el movimiento. El país de Origen debe contener mas tropas libres (que no hayan sido movidas en esta ronda) de las que desea mover.", errNoMovimientoNoTropas, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que los paises sean limitrofes
    strSQL = "SELECT COUNT(*) FROM Limites " & _
             "WHERE Pas_Id_Desde = " & intPaisDesde & _
             "  AND Pas_Id_Hasta = " & intPaisHasta
    
    If CInt(EjecutarConsultaValor(strSQL)) <= 0 Then
        sEnviarError "No es posible efectuar el movimiento. Los países seleccionados no son limítrofes.", errNoMovimientoNoLimitrofes, GintOrigenMensaje
        Exit Sub
    End If
    
    'Toma de la BD la cantidad de tropas del pais Desde
    strSQL = "SELECT Tro_Cantidad FROM Tropas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Pas_Id = " & intPaisDesde
    intCantTropasDesde = CInt(EjecutarConsultaValor(strSQL))
    
    'Toma de la BD la cantidad de tropas del pais Hasta
    strSQL = "SELECT Tro_Cantidad FROM Tropas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Pas_Id = " & intPaisHasta
    intCantTropasHasta = CInt(EjecutarConsultaValor(strSQL))
    
    
    'Actualiza la BD con los resultados del movimiento
    intTropasResultadoDesde = intCantTropasDesde - intCantidadTropas
    intTropasResultadoHasta = intCantTropasHasta + intCantidadTropas
    
    'Desde
    strSQL = "UPDATE Tropas SET " & _
             "Tro_Cantidad = " & intTropasResultadoDesde & _
             " WHERE Pas_Id = " & intPaisDesde & _
             "  AND Par_Id = " & GintPartidaActiva
    EjecutarComando strSQL
    
    'Hasta
    If intTipoMovimiento = tmConquista Then
        strSQL = "UPDATE Tropas SET " & _
                 "Tro_Cantidad = " & intTropasResultadoHasta & _
                 " WHERE Pas_Id = " & intPaisHasta & _
                 "  AND Par_Id = " & GintPartidaActiva
    Else
        strSQL = "UPDATE Tropas SET " & _
                 "Tro_Cantidad = " & intTropasResultadoHasta & ", " & _
                 "Tro_Fijos = Tro_Fijos + " & intCantidadTropas & _
                 " WHERE Pas_Id = " & intPaisHasta & _
                 "  AND Par_Id = " & GintPartidaActiva
    End If
    EjecutarComando strSQL
        
    'Informa al cliente los resultados del movimiento
    
    'Envia el mensaje a los clientes (broadcast)
    EnviarMensaje ArmarMensajeParam(msgAckMovimiento, intPaisDesde, intColorDesde, intTropasResultadoDesde, _
                                                      intPaisHasta, intColorHasta, intTropasResultadoHasta, _
                                                      intTipoMovimiento, intCantidadTropas)
    
    'Si todo salio bien actualiza el estado del jugador activo (solo si no es una conquista)
    If intTipoMovimiento = tmMovimiento Then
        ActualizarEstadoServidor eveMovimiento
    End If
    
    '###Log
    If intTipoMovimiento = tmConquista Then
        If intCantidadTropas = 1 Then
            GuardarLog "Se ha pasado 1 tropa mas de " & GvecPaises(intPaisDesde) & " a " & GvecPaises(intPaisHasta) & "."
            sEnviarLog mscMovimientoConquistaN1, _
                      CstrTipoParametroRecurso & CStr(intPaisDesde + enuIndiceArchivoRecurso.pmsPaises), _
                      CstrTipoParametroRecurso & CStr(intPaisHasta + enuIndiceArchivoRecurso.pmsPaises)
        Else
            GuardarLog "Se han pasado " & intCantidadTropas & " tropas mas de " & GvecPaises(intPaisDesde) & " a " & GvecPaises(intPaisHasta) & "."
            sEnviarLog mscMovimientoConquistaN, CstrTipoParametroResuelto & CStr(intCantidadTropas), _
                      CstrTipoParametroRecurso & CStr(intPaisDesde + enuIndiceArchivoRecurso.pmsPaises), _
                      CstrTipoParametroRecurso & CStr(intPaisHasta + enuIndiceArchivoRecurso.pmsPaises)
        End If

    Else
        If intCantidadTropas = 1 Then
            GuardarLog "Se ha movido 1 tropa de " & GvecPaises(intPaisDesde) & " a " & GvecPaises(intPaisHasta) & "."
            sEnviarLog mscMovimiento1Tropa, _
                      CstrTipoParametroRecurso & CStr(intPaisDesde + enuIndiceArchivoRecurso.pmsPaises), _
                      CstrTipoParametroRecurso & CStr(intPaisHasta + enuIndiceArchivoRecurso.pmsPaises)
        Else
            GuardarLog "Se han movido " & intCantidadTropas & " tropas de " & GvecPaises(intPaisDesde) & " a " & GvecPaises(intPaisHasta) & "."
            sEnviarLog mscMovimientoTropas, CstrTipoParametroResuelto & CStr(intCantidadTropas), _
                      CstrTipoParametroRecurso & CStr(intPaisDesde + enuIndiceArchivoRecurso.pmsPaises), _
                      CstrTipoParametroRecurso & CStr(intPaisHasta + enuIndiceArchivoRecurso.pmsPaises)
        End If
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "sMover", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sTomarTarjeta()
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim intTipoRonda As Integer
    Dim intCantTarjetas As Integer
    Dim intCantConquistas As Integer
    Dim intCantCanjes As Integer
    Dim vecTarjetasDisponibles() As String
    Dim intTarjetaSorteada As Integer
    
    'Verifica que el jugador que pidió la tarjeta sea el activo
    If Not EsElJugadorActivo(IndiceAColor(GintOrigenMensaje)) Then
        sEnviarError "Acción incorrecta. Imposible Tomar Tarjeta dado que no es su turno. " & _
                     "Intente resincronizarse.", errNoTomarNoTurno, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que la ronda sea de acción
    '### No es necesario con la matriz de estados
    intTipoRonda = GetTipoDeRondaActiva
    
    If intTipoRonda <> trAccion Then
        sEnviarError "No es posible Tomar Tarjeta en una ronda que no sea de Acción.", errNoTomarNoRondaAccion, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida contra la matriz de estados
    If Not ValidarProximoEstado(evePedidoTarjeta) Then
        Exit Sub
    End If
    
    'Valida que haya efectuado las conquistas necesarias
    'Obtiene de la BD la cantidad de conquistas
    strSQL = "SELECT Par_Activo_Conquistas FROM Partidas WHERE Par_Id = " & GintPartidaActiva
    intCantConquistas = CInt(EjecutarConsultaValor(strSQL))
    
    'Toma de la BD la cantidad de canjes
    strSQL = "SELECT Jug_Nro_Canje FROM Jugadores " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Col_Id = " & IndiceAColor(GintOrigenMensaje)
    intCantCanjes = CInt(EjecutarConsultaValor(strSQL))
             
    If intCantCanjes <= 3 Then
        'Si la cantidad de canjes es menor o igual a 3 solo necesita una conquista
        If intCantConquistas < 1 Then
            sEnviarError "No es posible Tomar Tarjeta sin haber conquistado al menos un país.", errNoTomarNoConquista1, GintOrigenMensaje
            Exit Sub
        End If
    Else
        'Si la cantidad de canjes es mayor a 3 necesita 2 conquistas
        If intCantConquistas < 2 Then
            sEnviarError "No es posible Tomar Tarjeta sin haber conquistado al menos dos países.", errNoTomarNoConquista2, GintOrigenMensaje
            Exit Sub
        End If
    End If
    
    'Verifica que tenga menos de cinco tarjetas
    strSQL = "SELECT COUNT(*) FROM Jugadores_Tarjetas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Col_Id = " & IndiceAColor(GintOrigenMensaje)
    intCantTarjetas = EjecutarConsultaValor(strSQL)
    
    If intCantTarjetas >= CintCantMaxTarjetas Then
        sEnviarError "No es posible Tomar Tarjeta. No puede tener mas de " & CintCantMaxTarjetas & " tarjetas.", errNoTomarNoMasTarjetas, GintOrigenMensaje
        Exit Sub
    End If
    
    'Toma de la BD las tarjetas disponibles
    strSQL = "SELECT Tar_Id FROM Tarjetas " & _
             "WHERE Tar_Id NOT IN (SELECT Tar_Id FROM Jugadores_Tarjetas " & _
                                   "WHERE Par_Id = " & GintPartidaActiva & ")"
    RecordsetAVector EjecutarConsulta(strSQL), 0, vecTarjetasDisponibles
    
    'Toma una tarjeta al azar de las disponibles
    intTarjetaSorteada = CInt(vecTarjetasDisponibles(Aleatorio(LBound(vecTarjetasDisponibles), UBound(vecTarjetasDisponibles))))
    
    'Actualiza la BD
    strSQL = "INSERT INTO Jugadores_Tarjetas(Par_Id, Col_Id, Tar_Id, Jut_Cobrada) VALUES(" & _
             GintPartidaActiva & ", " & _
             IndiceAColor(GintOrigenMensaje) & ", " & _
             intTarjetaSorteada & ", " & _
             "'N')"
    EjecutarComando strSQL
    
    'Informa a todos los clientes que se tomó una tarjeta (broadcast)
    sActualizarTarjetasJugador (IndiceAColor(GintOrigenMensaje))
    
    'Informa al jugador la nueva tarjeta
    sEnviarTarjeta intTarjetaSorteada
    
    '###Log
    GuardarLog GvecNombreJugadorColor(IndiceAColor(GintOrigenMensaje)) & " tomó tarjeta."
    sEnviarLog mscTarjetaTomada, CstrTipoParametroResuelto & GvecNombreJugadorColor(IndiceAColor(GintOrigenMensaje))
    
    'Si todo salio bien actualiza el estado del jugador activo
    ActualizarEstadoServidor evePedidoTarjeta
    
    Exit Sub
ErrorHandle:
    ReportErr "sTomarTarjeta", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sActualizarTarjetasJugador(intColor As Integer, Optional intDestino As Integer)
    'Informa la cantidad de tarjetas de un jugador
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim intCantTarjetas As Integer
    
    strSQL = "SELECT COUNT(*) FROM Jugadores_Tarjetas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Col_Id = " & intColor
    intCantTarjetas = CInt(EjecutarConsultaValor(strSQL))
    
    'Envia el mensaje
    EnviarMensaje ArmarMensajeParam(msgTarjetasJugador, intColor, intCantTarjetas), intDestino
             
    Exit Sub
ErrorHandle:
    ReportErr "sActualizarTarjetasJugador", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sEnviarTarjeta(intTarjeta As Integer)
    'Informa al dueño los valores de la tarjeta pasada por parametro
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim rsTarjeta As Recordset
    
    strSQL = "SELECT Jut.Col_Id, Jut.Jut_Cobrada, Tar.Fig_Id, Tar.Pas_Id " & _
             "FROM Jugadores_Tarjetas Jut, Tarjetas Tar " & _
             "WHERE Jut.Tar_Id = Tar.Tar_Id" & _
             "  AND Jut.Par_Id = " & GintPartidaActiva & _
             "  AND Jut.Tar_Id = " & intTarjeta
    Set rsTarjeta = EjecutarConsulta(strSQL)
    
    If Not rsTarjeta.EOF Then
        EnviarMensaje ArmarMensajeParam(msgTarjeta, rsTarjeta!Pas_Id, rsTarjeta!Fig_Id, rsTarjeta!Jut_Cobrada), ColorAIndice(CInt(rsTarjeta!Col_Id))
    End If
    
    rsTarjeta.Close
    Set rsTarjeta = Nothing
    
    Exit Sub
ErrorHandle:
    Set rsTarjeta = Nothing
    ReportErr "sEnviarTarjeta", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sCobrarTarjeta(intPais As Integer)
    
    On Error GoTo ErrorHandle
    Dim intTarjeta As Integer
    Dim intDuenioPais As Integer
    Dim intColorJugadorActivo As Integer
    Dim blnCobrada As Boolean
    Dim intTipoRonda As enuTipoRonda
    Dim intNuevaCantidadTropas As Integer
    Dim intTropasPorCobro As Integer
    Dim strSQL As String
    
    intColorJugadorActivo = IndiceAColor(GintOrigenMensaje)
    
    'Verifica que el jugador que pidió la tarjeta sea el activo
    If Not EsElJugadorActivo(intColorJugadorActivo) Then
        sEnviarError "Acción incorrecta. Imposible Cobrar Tarjeta dado que no es su turno. " & _
                     "Intente resincronizarse.", errNoCobrarNoTurno, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que la ronda sea de acción
    '### No es necesario con la matriz de estados
    intTipoRonda = GetTipoDeRondaActiva
    
    If intTipoRonda <> trAccion Then
        sEnviarError "No es posible Cobrar una Tarjeta en una ronda que no sea de Acción.", errNoCobrarNoRondaAccion, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida contra la Matriz de Estados
    If Not ValidarProximoEstado(eveCobroTarjeta) Then
        Exit Sub
    End If
    
    'Obtiene de la BD el Id de la tarjeta que corresponde al pais
    strSQL = "SELECT Tar_Id FROM Tarjetas WHERE Pas_Id = " & intPais
    intTarjeta = EjecutarConsultaValor(strSQL)
    
    'Valida que el pais que se corresponde con la tarjeta sea del jugador actual
    strSQL = "SELECT Col_Id FROM Jugadores_Tarjetas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Tar_Id = " & intTarjeta
    intDuenioPais = CInt(EjecutarConsultaValor(strSQL))
    
    If intDuenioPais <> intColorJugadorActivo Then
        sEnviarError "No es posible Cobrar la Tarjeta ya que el país al cual se refiere la tarjeta no le pertenece.", errNoCobrarNoPais, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que la tarjeta no haya sido cobrada
    strSQL = "SELECT Jut_Cobrada FROM Jugadores_Tarjetas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Tar_Id = " & intTarjeta
    
    blnCobrada = IIf(Trim(UCase(EjecutarConsultaValor(strSQL))) = "S", True, False)
    
    If blnCobrada Then
        sEnviarError "No es posible Cobrar la Tarjeta ya que la misma ya fue cobrada.", errNoCobrarYaCobrada, GintOrigenMensaje
        Exit Sub
    End If
    
    'Obtiene de la BD la cantidad de tropas que le corresponden al cobro de la tarjeta
    intTropasPorCobro = ValorOpcion(opBonusTarjetaPropia)
    
    'Actualiza la BD
    'Marca la tarjeta como cobrada
    strSQL = "UPDATE Jugadores_Tarjetas SET Jut_Cobrada = 'S' " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Tar_Id = " & intTarjeta
    EjecutarComando strSQL
    
    'Agrega 2 tropas al pais en cuestion
    strSQL = "SELECT Tro_Cantidad FROM Tropas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Pas_Id = " & intPais
    intNuevaCantidadTropas = CInt(EjecutarConsultaValor(strSQL)) + intTropasPorCobro
    
    '###Las tropas que se cobran son fijas??
    strSQL = "UPDATE Tropas SET " & _
             "Tro_Cantidad = " & intNuevaCantidadTropas & ", " & _
             "Tro_Fijos = Tro_Fijos + " & intTropasPorCobro & _
             " WHERE Par_Id = " & GintPartidaActiva & _
             "   AND Pas_Id = " & intPais
    EjecutarComando strSQL
    
    'Envía la confirmación (broadcast)
    EnviarMensaje ArmarMensajeParam(msgAckCobroTarjeta, intPais, intColorJugadorActivo, intNuevaCantidadTropas)
    
    '###Log
    GuardarLog GvecNombreJugadorColor(intColorJugadorActivo) & " cobró la tarjeta del pais " & GvecPaises(intPais)
    sEnviarLog mscTarjetaCobrada, CstrTipoParametroResuelto & GvecNombreJugadorColor(intColorJugadorActivo), _
                                  CstrTipoParametroRecurso & CStr(intPais + enuIndiceArchivoRecurso.pmsPaises)
    
    'Si todo salio bien actualiza el estado del jugador activo
    ActualizarEstadoServidor eveCobroTarjeta
    
    Exit Sub
ErrorHandle:
    ReportErr "sCobrarTarjeta", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub






'FUNCIONES RELACIONADAS CON EL ESTADO
'------------------------------------

Private Sub CargarRegistroEstado(estadoOrigen As enuEstadoActivo, eventoOrigen As enuEventosServidor, estadoDestino As enuEstadoActivo)
    'Subrutina utilizada para facilitar la carga de la matriz de estados
    On Error GoTo ErrorHandle
    
    GMatrizEstados(estadoOrigen, eventoOrigen) = estadoDestino
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarRegistroEstado", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CargarMatrizEstados()
    'Carga la matriz de estados del servidor
    On Error GoTo ErrorHandle
    
    CargarRegistroEstado estAgregando, eveAgregarTropas, estAgregando
    CargarRegistroEstado estAgregando, eveAtaque, 100
    CargarRegistroEstado estAgregando, eveMovimiento, 101
    CargarRegistroEstado estAgregando, evePedidoTarjeta, 102
    CargarRegistroEstado estAgregando, eveCobroTarjeta, 103
    CargarRegistroEstado estAgregando, eveCanjeTarjeta, estAgregando
'    CargarRegistroEstado estAgregando, eveFinTurnoInicial, estAgregando
    CargarRegistroEstado estAgregando, eveFinTurnoRecuento, estAgregando
    CargarRegistroEstado estAgregando, eveFinTurnoAccion, estAtacando
    
    CargarRegistroEstado estAtacando, eveAgregarTropas, 104
    CargarRegistroEstado estAtacando, eveAtaque, estAtacando
    CargarRegistroEstado estAtacando, eveMovimiento, estMoviendo
    CargarRegistroEstado estAtacando, evePedidoTarjeta, estTarjetaTomada
    CargarRegistroEstado estAtacando, eveCobroTarjeta, estTarjetaCobrada
    CargarRegistroEstado estAtacando, eveCanjeTarjeta, 105
'    CargarRegistroEstado estAtacando, eveFinTurnoInicial, estAgregando
    CargarRegistroEstado estAtacando, eveFinTurnoRecuento, estAgregando
    CargarRegistroEstado estAtacando, eveFinTurnoAccion, estAtacando
    
    CargarRegistroEstado estMoviendo, eveAgregarTropas, 106
    CargarRegistroEstado estMoviendo, eveAtaque, 107
    CargarRegistroEstado estMoviendo, eveMovimiento, estMoviendo
    CargarRegistroEstado estMoviendo, evePedidoTarjeta, estTarjetaTomada
    CargarRegistroEstado estMoviendo, eveCobroTarjeta, estTarjetaCobrada
    CargarRegistroEstado estMoviendo, eveCanjeTarjeta, 108
'    CargarRegistroEstado estMoviendo, eveFinTurnoInicial, estAgregando
    CargarRegistroEstado estMoviendo, eveFinTurnoRecuento, estAgregando
    CargarRegistroEstado estMoviendo, eveFinTurnoAccion, estAtacando
    
    CargarRegistroEstado estTarjetaCobrada, eveAgregarTropas, 109
    CargarRegistroEstado estTarjetaCobrada, eveAtaque, 110
    CargarRegistroEstado estTarjetaCobrada, eveMovimiento, 111
    CargarRegistroEstado estTarjetaCobrada, evePedidoTarjeta, estTarjetaTomadaCobrada
    CargarRegistroEstado estTarjetaCobrada, eveCobroTarjeta, estTarjetaCobrada
    CargarRegistroEstado estTarjetaCobrada, eveCanjeTarjeta, 113
'    CargarRegistroEstado estTarjetaCobrada, eveFinTurnoInicial, estAgregando
    CargarRegistroEstado estTarjetaCobrada, eveFinTurnoRecuento, estAgregando
    CargarRegistroEstado estTarjetaCobrada, eveFinTurnoAccion, estAtacando
    
    CargarRegistroEstado estTarjetaTomada, eveAgregarTropas, 114
    CargarRegistroEstado estTarjetaTomada, eveAtaque, 115
    CargarRegistroEstado estTarjetaTomada, eveMovimiento, 116
    CargarRegistroEstado estTarjetaTomada, evePedidoTarjeta, 117
    CargarRegistroEstado estTarjetaTomada, eveCobroTarjeta, estTarjetaTomadaCobrada
    CargarRegistroEstado estTarjetaTomada, eveCanjeTarjeta, 118
'    CargarRegistroEstado estTarjetaTomada, eveFinTurnoInicial, estAgregando
    CargarRegistroEstado estTarjetaTomada, eveFinTurnoRecuento, estAgregando
    CargarRegistroEstado estTarjetaTomada, eveFinTurnoAccion, estAtacando
    
    CargarRegistroEstado estTarjetaTomadaCobrada, eveAgregarTropas, 119
    CargarRegistroEstado estTarjetaTomadaCobrada, eveAtaque, 120
    CargarRegistroEstado estTarjetaTomadaCobrada, eveMovimiento, 121
    CargarRegistroEstado estTarjetaTomadaCobrada, evePedidoTarjeta, 122
    CargarRegistroEstado estTarjetaTomadaCobrada, eveCobroTarjeta, estTarjetaTomadaCobrada
    CargarRegistroEstado estTarjetaTomadaCobrada, eveCanjeTarjeta, 123
'    CargarRegistroEstado estTarjetaTomadaCobrada, eveFinTurnoInicial, estAgregando
    CargarRegistroEstado estTarjetaTomadaCobrada, eveFinTurnoRecuento, estAgregando
    CargarRegistroEstado estTarjetaTomadaCobrada, eveFinTurnoAccion, estAtacando
    
    '### Borrar
    vecEstadosActivo(estTarjetaTomadaCobrada) = "Tarjeta Cobrada y Tomada"
    vecEstadosActivo(estTarjetaTomada) = "Tarjeta Tomada"
    vecEstadosActivo(estTarjetaCobrada) = "Tarjeta Cobrada"
    vecEstadosActivo(estMoviendo) = "Moviendo"
    vecEstadosActivo(enuEstadoActivo.estAgregando) = "Agregando"
    vecEstadosActivo(enuEstadoActivo.estAtacando) = "Atacando"
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarMatrizEstados", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Function ValidarProximoEstado(intNuevoEvento As enuEventosServidor) As Boolean
    'Comprueba que se pueda pasar al proximo estado sin cambiar el estado
    On Error GoTo ErrorHandle
    Dim intPosibleEstado As Integer
    
    'Si el estado no es jugando
    If GEstadoServidor < estEjecutandoPartida Then
        sEnviarError "Imposible ejecutar la acción dado que aún no se ha iniciado la partida.", errNoAccionNoPartida, GintOrigenMensaje
        ValidarProximoEstado = False
    Else
        'Obtiene el estado correspondiente al evento
        intPosibleEstado = GMatrizEstados(GEstadoActivo, intNuevoEvento)
        If intPosibleEstado < 100 Then
            ValidarProximoEstado = True
        Else
            '### Tomar el error de la base de datos
            sEnviarError "Error", errNoAccionNoEstado, GintOrigenMensaje
            
            ValidarProximoEstado = False
        End If
    End If
    
    Exit Function
ErrorHandle:
    ReportErr "ValidarProximoEstado", "mdlServidor", Err.Description, Err.Number, Err.Source
End Function

Public Sub ActualizarEstadoServidor(intNuevoEvento As enuEventosServidor)
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim intProximoEstado As Integer
    
    intProximoEstado = GMatrizEstados(GEstadoActivo, intNuevoEvento)
    
    'Actualiza el estado del cliente solo si el nuevo estado corresponde
    'a un estado valido y actualiza la base solo si se modifico
    If intProximoEstado < 100 And intProximoEstado <> GEstadoActivo Then
        GEstadoActivo = intProximoEstado
        'Actualiza la base de datos con el estado actual
        strSQL = "UPDATE Partidas SET Par_Activo_Estado = " & GEstadoActivo & _
                 " WHERE Par_Id = " & GintPartidaActiva
        EjecutarComando strSQL
        
        '### Borrar
        frmServer.lblEstadoJugadorActivo.Caption = vecEstadosActivo(GEstadoActivo)
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "ValidarProximoEstado", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sCanjearTarjeta(intPais1 As Integer, intPais2 As Integer, intPais3 As Integer)
    On Error GoTo ErrorHandle
    Dim intColorJugadorActivo As Integer
    Dim vecFigurasTarjetas() As String
    Dim blnCanjeValido As Boolean
    Dim intNroCanje As Integer
    Dim intBonus As Integer
    Dim strSQL As String
    
    intColorJugadorActivo = IndiceAColor(GintOrigenMensaje)
    
    'Verifica que el jugador que pidió la tarjeta sea el activo
    If Not EsElJugadorActivo(intColorJugadorActivo) Then
        sEnviarError "Acción incorrecta. Imposible Canjear Tarjeta dado que no es su turno. " & _
                     "Intente resincronizarse.", errNoCanjeNoTurno, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida contra la matriz de estados
    If Not ValidarProximoEstado(eveCanjeTarjeta) Then
        Exit Sub
    End If
    
    'Obtiene de la BD las figuras de los paises pasados por parametro
    strSQL = "SELECT Tar.Fig_Id FROM Tarjetas Tar, Jugadores_Tarjetas Jut " & _
             "WHERE Tar.Tar_Id = Jut.Tar_Id " & _
             "  AND Jut.Par_Id = " & GintPartidaActiva & _
             "  AND Jut.Col_Id = " & IndiceAColor(GintOrigenMensaje) & _
             "  AND Tar.Pas_Id IN (" & intPais1 & ", " & intPais2 & ", " & intPais3 & ")"
             
    RecordsetAVector EjecutarConsulta(strSQL), 0, vecFigurasTarjetas
    
    'Valida que todas las tarjetas sean del jugador
    If UBound(vecFigurasTarjetas) < 2 Then
        sEnviarError "Imposible canjear las tarjetas. Debe seleccionar tres tarjetas que sean suyas", errNoCanjeNoTarjetas, GintOrigenMensaje
        Exit Sub
    End If
    
    'Valida que las tarjetas sean válidas
    blnCanjeValido = False
    If vecFigurasTarjetas(0) = figComodin _
    Or vecFigurasTarjetas(1) = figComodin _
    Or vecFigurasTarjetas(2) = figComodin Then
        'Si hay algun comodin
        blnCanjeValido = True
    Else
        'Si no hay ningun comodin
        If vecFigurasTarjetas(0) = vecFigurasTarjetas(1) And vecFigurasTarjetas(1) = vecFigurasTarjetas(2) Then
            'Si son las tres iguales
            blnCanjeValido = True
        ElseIf vecFigurasTarjetas(0) <> vecFigurasTarjetas(1) _
           And vecFigurasTarjetas(1) <> vecFigurasTarjetas(2) _
           And vecFigurasTarjetas(0) <> vecFigurasTarjetas(2) Then
            blnCanjeValido = True
        End If
    End If
    
    If Not blnCanjeValido Then
        sEnviarError "Imposible canjear las tarjetas. Debe seleccionar tres tarjetas con la misma figura o bien tres tarjetas con figuras distintas.", errNoCanjeNoFiguras, _
                     GintOrigenMensaje
        Exit Sub
    End If
    
    'Calcula el bonus
    
    'Obtiene de la BD el numero de canje
    strSQL = "SELECT Jug_Nro_Canje FROM Jugadores " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Col_Id = " & intColorJugadorActivo
    intNroCanje = CInt(EjecutarConsultaValor(strSQL)) + 1
    
    Select Case intNroCanje
        Case 1 'Es el primer canje
            intBonus = CInt(ValorOpcion(opCanjeNro1))
        Case 2 'Es el segundo
            intBonus = CInt(ValorOpcion(opCanjeNro2))
        Case 3 'Es el tercero
            intBonus = CInt(ValorOpcion(opCanjeNro3))
        Case Is >= 4
            intBonus = CInt(ValorOpcion(opCanjeNro3)) + ((intNroCanje - 3) * CInt(ValorOpcion(opCanjeIncremento)))
    End Select
    
    'Actualiza la BD
    
    'Suma 1 al número de canje
    strSQL = "UPDATE Jugadores SET Jug_Nro_Canje = " & intNroCanje & _
             " WHERE Par_Id = " & GintPartidaActiva & _
             "   AND Col_Id = " & intColorJugadorActivo
    EjecutarComando strSQL
    
    'Suma el bonus a las tropas disponibles
    strSQL = "UPDATE Tropas_Disponibles SET Tdi_Cantidad = Tdi_Cantidad + " & intBonus & _
             " WHERE Par_Id = " & GintPartidaActiva & _
             "   AND Col_Id = " & intColorJugadorActivo & _
             "   AND Con_Id IS NULL"
    EjecutarComando strSQL
    
    'Borra de la BD las tarjetas involucradas
    strSQL = "DELETE FROM Jugadores_Tarjetas " & _
             "WHERE Par_Id = " & GintPartidaActiva & _
             "  AND Col_Id = " & intColorJugadorActivo & _
             "  AND Tar_Id IN (SELECT Tar_Id FROM Tarjetas " & _
                              "WHERE Pas_Id IN (" & intPais1 & ", " & intPais2 & ", " & intPais3 & ")" & _
                              ")"
    EjecutarComando strSQL
    
    'Envia los mensajes
    'Ack de canjear tarjeta (unicast)
    EnviarMensaje ArmarMensajeParam(msgAckCanjeTarjeta, intPais1, intPais2, intPais3), GintOrigenMensaje
    
    'Tropas disponibles (broadcast)
    sInformarTropasDisponibles intColorJugadorActivo
    
    'Tarjetas disponibles (broadcast)
    sActualizarTarjetasJugador intColorJugadorActivo
    
    'Número de canje (broadcast)
    EnviarMensaje ArmarMensajeParam(msgCanjesJugador, intColorJugadorActivo, intNroCanje)
    
    '###Log
    GuardarLog GvecNombreJugadorColor(intColorJugadorActivo) & " canjeó las tarjetas de " & GvecPaises(intPais1) & ", " & GvecPaises(intPais2) & ", " & GvecPaises(intPais3)
    sEnviarLog mscCanje, CstrTipoParametroResuelto & GvecNombreJugadorColor(intColorJugadorActivo), _
                         CstrTipoParametroRecurso & CStr(intPais1 + enuIndiceArchivoRecurso.pmsPaises), _
                         CstrTipoParametroRecurso & CStr(intPais2 + enuIndiceArchivoRecurso.pmsPaises), _
                         CstrTipoParametroRecurso & CStr(intPais3 + enuIndiceArchivoRecurso.pmsPaises)
    
    Exit Sub
ErrorHandle:
    ReportErr "sCanjearTarjeta", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Private Function CumplioMision(intColor As Integer, ByRef blnObjetivoComun As Boolean) As Boolean
    'Verifica si el jugador pasado por parámetro completó la misión
    'Informa si se cumplió el objetivo común
    On Error GoTo ErrorHandle
    Dim intCantPaisesObjetivoComun As Integer
    Dim intCantPaises As Integer
    Dim varAsesino As Variant
    Dim strSQL As String
    Dim rsObjetivos As Recordset
    Dim blnReturn As Boolean
    
    blnObjetivoComun = False
    
    'Primero se fija se juega a Conquistar el Mundo (C) o por Misiones (M).
    'OJO!: Cuando esta Opcion es modificada durante el juego, la Mision de los jugadores
    'no se mofica y siguen con la mision asignada al comienzo, tal como si juagran con Misiones
    'Por este motivo es conveniente revisar directamente Opcion en lugar de ver si la Mision=0
    If ValorOpcion(opMisionTipo) = "C" Then
        'Si se juega a Conquistar el Mundo, controla que el jugador tenga los 50 paises.
        
        'Toma de la BD la cantidad de paises del jugador activo
        strSQL = "SELECT COUNT(*) FROM Tropas " & _
                 "WHERE Par_Id = " & GintPartidaActiva & _
                 "  AND Col_Id = " & intColor
        intCantPaises = CInt(EjecutarConsultaValor(strSQL))
        
        If intCantPaises = 50 Then
            CumplioMision = True
            blnObjetivoComun = False
            Exit Function
        Else
            CumplioMision = False
            blnObjetivoComun = False
            Exit Function
        End If
    
    Else
        'Si no se juega a conquistar el mundo, controla el Objetivo Comun
    
        'Toma de la BD la cantidad de paises del objetivo comun
        intCantPaisesObjetivoComun = ValorOpcion(opMisionObjetivoComun)
        
        'Verifica si se completó el objetivo común
        strSQL = "SELECT COUNT(*) FROM Tropas " & _
                 "WHERE Par_Id = " & GintPartidaActiva & _
                 "  AND Col_Id = " & intColor
        intCantPaises = CInt(EjecutarConsultaValor(strSQL))
        
        If intCantPaises >= intCantPaisesObjetivoComun Then
            CumplioMision = True
            blnObjetivoComun = True
            Exit Function
        End If
    
    End If

    'Si se juega por Misiones y no se cumplio el Objetivo Comun, controla los Objetivos de la Mision

    'Toma de la BD los objetivos de la misión
    strSQL = "SELECT Mob.* FROM Misiones_Objetivos Mob, Jugadores Jug " & _
             "WHERE Mob.Mis_Id = Jug.Mis_Id" & _
             "  AND Jug.Col_Id = " & intColor & _
             "  AND Jug.Par_Id = " & GintPartidaActiva
    Set rsObjetivos = EjecutarConsulta(strSQL)
    
    blnReturn = True
    While Not (rsObjetivos.EOF) And blnReturn = True
        
        'OBJETIVO DESTRUIR
        If Not IsNull(rsObjetivos!Col_Id) Then
            'Toma de la BD el asesino del color a eliminar
            strSQL = "SELECT Jug_Eliminado_Por FROM Jugadores " & _
                     "WHERE Par_Id = " & GintPartidaActiva & _
                     "  AND Col_Id = " & CInt(rsObjetivos!Col_Id)
            
            varAsesino = EjecutarConsultaValor(strSQL)
            'Si el jugador fue asesinado
            If Not IsNull(varAsesino) Then
                'Si el asesino es el color pasado por parametro
                If CInt(varAsesino) = intColor Then
                    blnReturn = blnReturn And True
                Else
                    blnReturn = False
                End If
            End If
        End If
        
        
        'OBJETIVO PAISES CONTINENTE
        If Not IsNull(rsObjetivos!Con_Id) Then
            'Toma de la BD la cantidad de paises que son del jugador y
            'pertenecen al continente
            strSQL = "SELECT COUNT(*) FROM Paises Pas, Tropas Tro " & _
                     "WHERE Pas.Pas_Id = Tro.Pas_Id" & _
                     "  AND Tro.Par_Id = " & GintPartidaActiva & _
                     "  AND Tro.Col_Id = " & intColor & _
                     "  AND Pas.Con_Id = " & rsObjetivos!Con_Id
            'Si la cantidad de paises es la necesaria
            If CInt(EjecutarConsultaValor(strSQL)) >= CInt(rsObjetivos!Mob_Cant_Paises) Then
                blnReturn = blnReturn And True
            Else
                blnReturn = False
            End If
        End If
        
        'OBJETIVO CONQUISTAR MUNDO
        If IsNull(rsObjetivos!Mob_Limitrofes) _
        And IsNull(rsObjetivos!Con_Id) _
        And IsNull(rsObjetivos!Col_Id) _
        And Not IsNull(rsObjetivos!Mob_Cant_Paises) _
        Then
            'Toma de la BD la cantidad de paises que son del jugador activo
            strSQL = "SELECT COUNT(*) FROM Tropas " & _
                     "WHERE Par_Id = " & GintPartidaActiva & _
                     "  AND Col_Id = " & intColor
            intCantPaises = CInt(EjecutarConsultaValor(strSQL))
            
            If intCantPaises >= rsObjetivos!Mob_Cant_Paises Then
                blnReturn = blnReturn And True
            Else
                blnReturn = False
            End If
        
        End If
        
        'OBJETIVO PAISES LIMITROFES
        If Not IsNull(rsObjetivos!Mob_Limitrofes) Then
            'Version Hardcodeada
            Dim intPaisesDeEuropa As Integer
            Dim rsTrios As Recordset
            Dim blnExisteTrioValido As Boolean
            
            'Toma de la BD la cantidad de paises que el jugador posee en Europa
            strSQL = "SELECT COUNT(*) " & _
                     "FROM Tropas Tro, Paises Pas " & _
                     "WHERE Tro.Pas_Id = Pas.Pas_Id " & _
                     "  AND Pas.Con_Id = 5 " & _
                     "  AND Tro.Col_Id = " & intColor & _
                     "  AND Tro.Par_Id = " & GintPartidaActiva
            intPaisesDeEuropa = CInt(EjecutarConsultaValor(strSQL))
            
            'Solo verifica limítrofes si la cantidad de paises de europa es >= 7
            '(como pide la mision)
            If intPaisesDeEuropa < 7 Then
                blnReturn = False
            Else
                'Toma de la BD los Trios de paises limitrofes
                'que posee el jugador y que no pertenecen a America del Sur (Con_Id=3)
                strSQL = "SELECT Pas1.Con_Id AS Con1, Pas2.Con_Id AS Con2, Pas3.Con_Id AS Con3 " & _
                         "FROM Paises Pas1, Tropas Tro1, Limites Lim1, Paises Pas2, Tropas Tro2, Limites Lim2, Tropas Tro3, Paises Pas3 " & _
                         "WHERE Tro1.Pas_Id = Lim1.Pas_Id_Desde " & _
                         "  AND Lim1.Pas_Id_Hasta = Tro2.Pas_Id " & _
                         "  AND Tro2.Pas_Id = Lim2.Pas_Id_Desde " & _
                         "  AND Lim2.Pas_Id_Hasta = Tro3.Pas_Id " & _
                         "  AND Tro1.Pas_Id <> Tro2.Pas_Id " & _
                         "  AND Tro1.Pas_Id <> Tro3.Pas_Id " & _
                         "  AND Tro2.Pas_Id <> Tro3.Pas_Id " & _
                         "  AND Tro1.Col_Id = Tro2.Col_Id " & _
                         "  AND Tro2.Col_Id = Tro3.Col_Id " & _
                         "  AND Tro1.Pas_Id = Pas1.Pas_Id " & _
                         "  AND Tro2.Pas_Id = Pas2.Pas_Id " & _
                         "  AND Tro3.Pas_Id = Pas3.Pas_Id " & _
                         "  AND Pas1.Con_Id <> 3 " & _
                         "  AND Pas2.Con_Id <> 3 " & _
                         "  AND Pas3.Con_Id <> 3 " & _
                         "  AND Tro1.Par_Id = Tro2.Par_Id " & _
                         "  AND Tro2.Par_Id = Tro3.Par_Id " & _
                         "  AND Tro1.Col_Id = " & intColor & _
                         "  AND Tro1.Par_Id = " & GintPartidaActiva
                Set rsTrios = EjecutarConsulta(strSQL)
                
                'Toma trio por trio y verifica que no pertenezcan a los 7 de Europa
                blnExisteTrioValido = False
                While Not rsTrios.EOF
                    'Valida segun la cantidad de paises que hay en Europa
                    Select Case intPaisesDeEuropa
                        Case 7
                            'Ninguno del trio puede pertenecer a Europa (Con_Id=5)
                            If rsTrios!Con1 <> 5 And rsTrios!Con2 <> 5 And rsTrios!Con3 <> 5 Then
                                blnExisteTrioValido = True
                            End If
                        Case 8
                            'Solo 1 del trio puede pertenecer a Europa (Con_Id=5)
                            If (rsTrios!Con1 <> 5 And rsTrios!Con2 <> 5) _
                            Or (rsTrios!Con2 <> 5 And rsTrios!Con3 <> 5) _
                            Or (rsTrios!Con1 <> 5 And rsTrios!Con3 <> 5) Then
                                blnExisteTrioValido = True
                            End If
                        Case 9
                            'Solo 1 del trio no puede pertenecer a Europa (Con_id=7)
                            If rsTrios!Con1 <> 5 Or rsTrios!Con2 <> 5 Or rsTrios!Con3 <> 5 Then
                                blnExisteTrioValido = True
                            End If
                    End Select
                            
                    rsTrios.MoveNext
                Wend
                
                If blnExisteTrioValido Then
                    blnReturn = blnReturn And True
                Else
                    blnReturn = False
                End If
                
                rsTrios.Close
                Set rsTrios = Nothing
            End If
        End If
            
        rsObjetivos.MoveNext
    Wend
        
    rsObjetivos.Close
    Set rsObjetivos = Nothing
    
    CumplioMision = blnReturn
    
    Exit Function
ErrorHandle:
    Set rsObjetivos = Nothing
    ReportErr "CumplioMision", "mdlServidor", Err.Description, Err.Number, Err.Source
End Function

Public Sub sResincronizar(intIndiceJugador As Integer)
    'Envía los mensajes correspondientes a la resincronización al
    'Jugador pasado por parametro
    On Error GoTo ErrorHandle
    Dim intColorJugador As Integer
    Dim strSQL As String
    Dim varGanador As Variant
    Dim strMisionGanador As String
    Dim intMisionId As Integer
    
    intColorJugador = IndiceAColor(intIndiceJugador)
    
    'Jugadores conectados
    sConexionesActuales intIndiceJugador
    
    'Manda mensajes pais
    sActualizarMapa intIndiceJugador
    
    'Obtiene la misión del jugador y la envía
    sEnviarMision intColorJugador
    
    'Orden de la ronda
    sActualizarRonda intIndiceJugador
    
    'Tipo de ronda
    sInformarTipoRonda intIndiceJugador
    
    'Cantidad de tropas disponibles de cada jugador
    sInformarTropasDisponiblesTodos intIndiceJugador
    
    'Detalle de las tropas disponibles del jugador a resincronizar
    sInformarTropasDisponibles intColorJugador
    
    'Tarjetas
    'Envia la cantidad de tarjetas de cada jugador
    sActualizarTarjetasTodos intIndiceJugador
    
    'Detalle de las tarjetas del jugador a resincronizar.
    sEnviarTodasLasTarjetas intColorJugador
    
    'Cantidad de canjes de todos los jugadores.
    sActualizarCanjesTodos intIndiceJugador
    
    'Envia las opciones
    sEnviarOpciones False, intIndiceJugador
    
    '###Version
    If frmServer.wskServer(intIndiceJugador).Tag = "" Or Mid(frmServer.wskServer(intIndiceJugador).Tag, 2) <= "10000" Then
        'Inicio de turno (tiene que ser lo último)
        sInformarInicioTurno False, intIndiceJugador
        
        'Envia el estado del turno del cliente
        sEnviarEstadoTurno intIndiceJugador
    Else
        'Envia el estado del turno del cliente
        sEnviarEstadoTurno intIndiceJugador
        
        'Inicio de turno (tiene que ser lo último)
        sInformarInicioTurno False, intIndiceJugador
    End If
    
    'Informa el jugador ganador (si hay un ganador de la partida)
    strSQL = "SELECT Par_Jug_Ganador FROM Partidas " & _
             " WHERE Par_Id = " & GintPartidaActiva
    varGanador = EjecutarConsultaValor(strSQL)
    
    If Not IsNull(varGanador) Then
        'Informa a los clientes la misión cumplida
        'Toma de la BD la descripción de la misión ganadora
        strSQL = "SELECT Mis.Mis_Desc " & _
                 "FROM Misiones Mis, Jugadores Jug " & _
                 "WHERE Mis.Mis_Id = Jug.Mis_Id " & _
                 "  AND Jug.Par_Id = " & GintPartidaActiva & _
                 "  AND Jug.Col_Id = " & varGanador
        strMisionGanador = EjecutarConsultaValor(strSQL)
        
        'Toma de la BD el código de la misión ganadora
        strSQL = "SELECT Mis.Mis_Id " & _
                 "FROM Misiones Mis, Jugadores Jug " & _
                 "WHERE Mis.Mis_Id = Jug.Mis_Id " & _
                 "  AND Jug.Par_Id = " & GintPartidaActiva & _
                 "  AND Jug.Col_Id = " & varGanador
        intMisionId = CInt(EjecutarConsultaValor(strSQL))
        
        'Envia el mensaje (unicast)
        EnviarMensaje ArmarMensajeParam(msgMisionCumplida, CStr(varGanador), strMisionGanador, enuIndiceArchivoRecurso.pmsMisiones + intMisionId + 1), intIndiceJugador
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "sResincronizar", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sActualizarCanjesTodos(Optional intDestino As Integer)
    'Informa la cantidad de canjes de cada jugador
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim rsColores As Recordset
    Dim intNroCanje As Integer
    
    'Obtiene de la BD los colores de los jugadores de la partida
    strSQL = "SELECT Col_Id FROM Jugadores WHERE Par_Id = " & GintPartidaActiva
    Set rsColores = EjecutarConsulta(strSQL)
    
    While Not rsColores.EOF
        'Informa la cantidad de tarjetas del jugador
        
        strSQL = "SELECT Jug_Nro_Canje FROM Jugadores " & _
                 "WHERE Par_Id = " & GintPartidaActiva & _
                 "  AND Col_Id = " & rsColores!Col_Id
        intNroCanje = CInt(EjecutarConsultaValor(strSQL))
        EnviarMensaje ArmarMensajeParam(msgCanjesJugador, rsColores!Col_Id, intNroCanje), intDestino
                
        rsColores.MoveNext
    Wend
    
    rsColores.Close
    Set rsColores = Nothing
    
    Exit Sub
ErrorHandle:
    Set rsColores = Nothing
    ReportErr "sActualizarCanjesTodos", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sActualizarTarjetasTodos(Optional intDestino As Integer)
    'Informa la cantidad de tarjetas de cada jugador
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim rsColores As Recordset
    
    'Obtiene de la BD los colores de los jugadores de la partida
    strSQL = "SELECT Col_Id FROM Jugadores WHERE Par_Id = " & GintPartidaActiva
    Set rsColores = EjecutarConsulta(strSQL)
    
    While Not rsColores.EOF
        'Informa la cantidad de tarjetas del jugador
        sActualizarTarjetasJugador rsColores!Col_Id, intDestino
        rsColores.MoveNext
    Wend
    
    rsColores.Close
    Set rsColores = Nothing
        
    Exit Sub
ErrorHandle:
    Set rsColores = Nothing
    ReportErr "sActualizarTarjetasTodos", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sEnviarTodasLasTarjetas(intColor As Integer)
    'Envia el detalle de las tarjetas del jugador determinado
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim rsTarjetas As Recordset
    
    strSQL = "SELECT Tar_Id FROM Jugadores_Tarjetas " & _
             "WHERE Col_Id = " & intColor & _
             "  AND Par_Id = " & GintPartidaActiva
    Set rsTarjetas = EjecutarConsulta(strSQL)
    
    While Not rsTarjetas.EOF
        sEnviarTarjeta rsTarjetas!Tar_Id
        rsTarjetas.MoveNext
    Wend
    
    rsTarjetas.Close
    Set rsTarjetas = Nothing
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarTodasLasTarjetas", "mdlServidor", Err.Description, Err.Number, Err.Source
    Set rsTarjetas = Nothing
End Sub

Public Sub sEnviarEstadoTurno(intDestino As Integer)
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim intEstadoServidor As enuEstadoActivo
    Dim intEstadoCliente As enuEstadoCli
    Dim strEstadoCliente As String
    Dim intColorActivo As Integer
    
    'Verifica si el jugador es el activo
    strSQL = "SELECT Par_Activo_Ronda " & _
             "FROM Partidas " & _
             "WHERE Par_Id = " & GintPartidaActiva
    intColorActivo = CInt(EjecutarConsultaValor(strSQL))
    
    If intColorActivo = IndiceAColor(intDestino) Then
        'Toma de la BD el estado del turno para el servidor
        strSQL = "SELECT Par_Activo_Estado " & _
                 "FROM Partidas " & _
                 "WHERE Par_Id = " & GintPartidaActiva
        intEstadoServidor = CInt(EjecutarConsultaValor(strSQL))
        
        'Traduce del estado del servidor al estado del cliente
        Select Case intEstadoServidor
            Case enuEstadoActivo.estAgregando
                intEstadoCliente = enuEstadoCli.estAgregando
            Case enuEstadoActivo.estAtacando
                intEstadoCliente = enuEstadoCli.estAtacando
            Case enuEstadoActivo.estMoviendo
                intEstadoCliente = enuEstadoCli.estMoviendo
            Case enuEstadoActivo.estTarjetaCobrada
                intEstadoCliente = enuEstadoCli.estTarjetaCobrada
            Case enuEstadoActivo.estTarjetaTomada
                intEstadoCliente = enuEstadoCli.estTarjetaTomada
            Case enuEstadoActivo.estTarjetaTomadaCobrada
                intEstadoCliente = enuEstadoCli.estTarjetaCobradaTomada
        End Select
    Else
        'Si no es el activo está esperando turno
        intEstadoCliente = estEsperandoTurno
    End If
    
    'Si el Servidor esta Pausado,
    'agrega dicho estado en segundo lugar dentro del mensaje
    strEstadoCliente = ""
    If GEstadoServidor = estPartidaDetenida Then
        'Estado del Cliente Detenido
        strEstadoCliente = CStr(enuEstadoCli.estPartidaPausada)
    End If
    
    'Envia el mensaje
    EnviarMensaje ArmarMensajeParam(msgEstadoTurnoCliente, intEstadoCliente, strEstadoCliente), intDestino
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarEstadoTurno", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sEliminarPartida(strNombre As String)
    On Error GoTo ErrorHandle
    'Elimina la partida con el nombre especificado
    Dim strSQL As String
    
    strSQL = "DELETE FROM Partidas WHERE trim(Par_Nombre) = '" & Trim(strNombre) & "'"
    
    EjecutarComando strSQL
    
    'Envia el mensaje con las partidas guardadas
    sEnviarPartidasGuardadas tpgGuardarPartida
    
    Exit Sub
ErrorHandle:
    ReportErr "sEliminarPartida", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CopiarPartida(strNombreOrigen As String, strNombreDestino As String)
    On Error GoTo ErrorHandle
    'Copia una partida y la guarda con el nombre especificado
    Dim strSQL As String
    Dim varMaxPartida As Variant
    Dim intCodigoPartidaDestino As Integer
    Dim intCodigoPartidaOrigen As Integer
    
    'Obtiene de la BD el Id de la partida de origen
    strSQL = "SELECT Par_Id FROM Partidas WHERE trim(Par_Nombre) = '" & Trim(strNombreOrigen) & "'"
    intCodigoPartidaOrigen = EjecutarConsultaValor(strSQL)
    
    'Borra la partida destino(si existe)
    strSQL = "DELETE FROM Partidas WHERE trim(Par_Nombre) = '" & Trim(strNombreDestino) & "'"
    EjecutarComando strSQL

    'CREA LA PARTIDA DESTINO
    'Obtiene el número de la última partida
    strSQL = "SELECT MAX(Par_Id) FROM Partidas"
    varMaxPartida = EjecutarConsultaValor(strSQL)
    If IsNull(varMaxPartida) Then
        varMaxPartida = 0
    End If
    
    intCodigoPartidaDestino = CInt(varMaxPartida) + 1
    
    'Inserta en la BD la nueva partida
    strSQL = "INSERT INTO Partidas (Par_Id, Par_Nombre, Par_Fecha_Creacion, Par_Fecha_Actu, Par_Ronda_Nro, Par_Ronda_Primero, Par_Ronda_Tipo, Par_Activo_Estado, Par_Activo_Ronda, Par_Activo_Conquistas) " & _
             "SELECT " & CStr(intCodigoPartidaDestino) & ", '" & strNombreDestino & "', Par_Fecha_Creacion , Date(), Par_Ronda_Nro, Par_Ronda_Primero, Par_Ronda_Tipo, Par_Activo_Estado, Par_Activo_Ronda, Par_Activo_Conquistas " & _
             "FROM Partidas WHERE Par_Id = " & intCodigoPartidaOrigen
    EjecutarComando strSQL
    
    'JUGADORES
    strSQL = "INSERT INTO Jugadores (Par_Id, Col_Id, Mis_Id, Jug_Nombre, Jug_Tipo, Jug_Nro_Canje, Jug_Prox_Ronda) " & _
             "SELECT " & CStr(intCodigoPartidaDestino) & ", Col_Id, Mis_Id, Jug_Nombre, Jug_Tipo, Jug_Nro_Canje, Jug_Prox_Ronda " & _
             "FROM Jugadores WHERE Par_Id = " & intCodigoPartidaOrigen
    EjecutarComando strSQL
    
    'TROPAS DISPONIBLES
    strSQL = "INSERT INTO Tropas_Disponibles (Par_Id, Con_Id, Col_Id, Tdi_Cantidad) " & _
             "SELECT " & CStr(intCodigoPartidaDestino) & ", Con_Id, Col_Id, Tdi_Cantidad " & _
             "FROM Tropas_Disponibles WHERE Par_Id = " & intCodigoPartidaOrigen
    EjecutarComando strSQL
    
    'TROPAS
    strSQL = "INSERT INTO Tropas (Par_Id, Col_Id, Pas_Id, Tro_Cantidad, Tro_Fijos) " & _
             "SELECT " & CStr(intCodigoPartidaDestino) & ", Col_Id, Pas_Id, Tro_Cantidad, Tro_Fijos " & _
             "FROM Tropas WHERE Par_Id = " & intCodigoPartidaOrigen
    EjecutarComando strSQL
    
    'JUGADORES TARJETAS
    strSQL = "INSERT INTO Jugadores_Tarjetas (Par_Id, Col_Id, Tar_Id, Jut_Cobrada) " & _
             "SELECT " & CStr(intCodigoPartidaDestino) & ", Col_Id, Tar_Id, Jut_Cobrada " & _
             "FROM Jugadores_Tarjetas WHERE Par_Id = " & intCodigoPartidaOrigen
    EjecutarComando strSQL
    
    'OPCIONES
    strSQL = "INSERT INTO Opciones (Par_Id, Opc_Id, Opc_Valor, Opc_Desc) " & _
             "SELECT " & CStr(intCodigoPartidaDestino) & ", Opc_Id, Opc_Valor, Opc_Desc " & _
             "FROM Opciones WHERE Par_Id = " & intCodigoPartidaOrigen
    EjecutarComando strSQL
    
    
    Exit Sub
ErrorHandle:
    ReportErr "CopiarPartida", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sGuardarPartida(strNombre As String)
    On Error GoTo ErrorHandle
    'Copia la partida actual con el nombre especificado
    Dim strSQL As String
    Dim varMaxPartida As Variant
    Dim intCodigoNuevaPartida As Integer
    Dim intOrigenMensaje As Integer
    
    intOrigenMensaje = GintOrigenMensaje
    
    'Valida que la acción sea solicitada por el administrador
    If intOrigenMensaje <> GintIndiceAdm Then
        EnviarMensaje ArmarMensajeParam(msgAckGuardarPartida, enuAckGuardarPartida.AckNoAdministrador), intOrigenMensaje
        Exit Sub
    End If
    
    CopiarPartida strNombrePartidaActual, strNombre
    
    'Responde con un ACK de OK
    EnviarMensaje ArmarMensajeParam(msgAckGuardarPartida, enuAckGuardarPartida.AckOk), intOrigenMensaje

    Exit Sub
ErrorHandle:
    ReportErr "sGuardarPartida", "mdlServidor", Err.Description, Err.Number, Err.Source
    'Responde con un ACK desconocido
    EnviarMensaje ArmarMensajeParam(msgAckGuardarPartida, enuAckGuardarPartida.AckErrorDesconocido), intOrigenMensaje

End Sub

Public Sub GuardarLog(strLog As String)
    On Error GoTo ErrorHandle
    
    'frmServer.lstLog.AddItem Date & " " & Time & " - " & strLog
    frmServer.lstLog.AddItem Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss") & " - " & strLog
    'Se posiciona en el último elemento
    frmServer.lstLog.ListIndex = frmServer.lstLog.ListCount - 1
    
    Exit Sub
ErrorHandle:
    ReportErr "GuardarLog", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sEnviarLog(intMascara As enuMascara, ParamArray vecParametros())
    'Envìa a los clientes un mensaje de log formado por una màscara y sus
    'correspondientes parámetros
    On Error GoTo ErrorHandle
    
    Dim vecLog() As String
    Dim intParametro As Integer
    
    'Arma un unico vector con la mascara y sus parametros
    ReDim vecLog(UBound(vecParametros) + 1)
    vecLog(0) = intMascara
    
    For intParametro = 1 To UBound(vecParametros) + 1
        vecLog(intParametro) = vecParametros(intParametro - 1)
    Next intParametro
    
    '###
    'EnviarMensaje ArmarMensajeVector(msgLog, vecLog)
    EnviarMensaje ArmarMensaje(msgLog, vecLog)

    Exit Sub
ErrorHandle:
    ReportErr "sEnviarLog", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CargarVectorColores()
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim vecColoresAux() As String
    Dim i As Integer
    
    strSQL = "SELECT Col_Nombre FROM Colores ORDER BY Col_Id"
    RecordsetAVector EjecutarConsulta(strSQL), 0, vecColoresAux
    
    'Carga el vector global con el auxiliar
    For i = 1 To UBound(GvecColores)
        GvecColores(i) = vecColoresAux(i - 1)
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarVectorColores", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CargarVectorPaises()
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim i As Integer
    
    strSQL = "SELECT Pas_Nombre FROM Paises ORDER BY Pas_Id"
    RecordsetAVector EjecutarConsulta(strSQL), 0, GvecPaises
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarVectorPaises", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub CargarVectorContinentes()
    On Error GoTo ErrorHandle
    Dim strSQL As String
    Dim vecContinentesAux() As String
    Dim i As Integer
    
    strSQL = "SELECT Con_Nombre FROM Continentes ORDER BY Con_Id"
    RecordsetAVector EjecutarConsulta(strSQL), 0, vecContinentesAux
    
    'Carga el vector global con el auxiliar
    For i = 1 To UBound(GvecContinentes)
        GvecContinentes(i) = vecContinentesAux(i - 1)
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarVectorContinentes", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub ReportErr(ByVal strFuncion As String, ByVal strModulo As String, ByVal strDesc As String, _
                    ByVal intErr As Long, ByVal strSource As String, _
                    Optional styIcono As VbMsgBoxStyle = vbCritical)
    'Reportes de Errores
    On Error GoTo ErrorHandle
    Dim strMsg As String
    
    'Cambia el icono
    SysTrayChangeIcon frmServer.hwnd, frmServer.imgSystray(1)
    'frmServer.imgSystray(0).Visible = False
    'frmServer.imgSystray(1).Visible = True
    
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

Public Sub sEnviarVersionServidor(intIndiceDestinatario As Integer)
    'Envia la versión del servidor al cliente que acaba de conectarse
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgVersionServidor, App.Major, App.Minor, App.Revision), intIndiceDestinatario
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarVersionServidor", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sRecibirVersionCliente(intIndiceOrigenMensaje As Integer, byTipoJugador As enuInteligenciaJugador, intMajor As Integer, intMinor As Integer, intRevision As Integer)
    'Valida que la versión del cliente sea compatible con la del servidor
    On Error GoTo ErrorHandle
    
    'Almacena la versión del cliente para consultarla posteriormente y saber si el cliente la
    'envió o no (la versión 1.0.0 no envía su versión)
    frmServer.wskServer(intIndiceOrigenMensaje).Tag = byTipoJugador & intMajor * 10000 + intMinor * 100 + intRevision
    
    GuardarLog frmServer.wskServer(intIndiceOrigenMensaje).RemoteHostIP & " - Versión del Cliente: " & intMajor & "." & intMinor & "." & intRevision & " (" & IIf(byTipoJugador = hrHumano, "Humano", IIf(byTipoJugador = hrRobot, "Robot", "Desconocido")) & ")"
    
    Exit Sub
ErrorHandle:
    ReportErr "sRecibirVersionCliente", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Private Sub sEnviarLimitrofes(Optional intDestino As Integer)
    On Error GoTo ErrorHandle
    Dim vecLimitrofes() As String
    Dim strSQL As String
    'Envia al cliente un mensaje con todos los paises limitrofes
    
    strSQL = "SELECT Pas_Id_Desde & ',' & Pas_Id_Hasta & '' as Campo FROM Limites ORDER BY Pas_Id_Desde"
    
    RecordsetAVector EjecutarConsulta(strSQL), 0, vecLimitrofes
    
    'broadcast o unicast
    EnviarMensaje ArmarMensaje(msgLimitrofes, vecLimitrofes), intDestino
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarLimitrofes", "mdlInterfase", Err.Description, Err.Number, Err.Source
End Sub

Private Sub sEnviarPaisContinente(Optional intDestino As Integer)
    On Error GoTo ErrorHandle
    Dim vecPaises() As String
    Dim strSQL As String
    'Envia al cliente un mensaje con los paises y sus respectivos continentes
    'Solo lo usan los JV
    
    strSQL = "SELECT Pas_Id & ',' & Con_Id & '' as Campo FROM Paises ORDER BY Pas_Id"
    
    RecordsetAVector EjecutarConsulta(strSQL), 0, vecPaises
    
    'broadcast o unicast
    EnviarMensaje ArmarMensaje(msgPaisContinente, vecPaises), intDestino
    
    Exit Sub
ErrorHandle:
    ReportErr "sEnviarPaisContinente", "mdlInterfase", Err.Description, Err.Number, Err.Source
End Sub
Public Sub sPausarPartida()
    'Realiza una Pausa en el Servidor y le avisa al
    'Robot de turno para que tambien se detenga.
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    'Valida que la acción sea solicitada por el Administrador,
    'y que sea el único Humano.
    If GintOrigenMensaje <> GintIndiceAdm Then
        sEnviarError CstrErrorNoAdm, errNoAccionNoAdm, GintOrigenMensaje
        Exit Sub
    End If
    'Chequea que el ADM se el unico Humano
    For i = 1 To UBound(GvecColoresSock)
        If GvecColoresSock(i) > 0 Then
            If i <> GintIndiceAdm And Mid$(frmServer.wskServer(i).Tag, 1, 1) <> enuInteligenciaJugador.hrRobot Then
                sEnviarError "No se puede Pausar la Partida, porque existen otros jugadores Humanos.", errNoPausaHayHumanos, GintIndiceAdm
                Exit Sub
            End If
        End If
    Next i
    
    'Envia la Pausa al Robot de turno
    sInformarPausaPartida
    
   'Congela el Servidor
    '###E
    'Cambia el estado del servidor
    CambiarEstadoServidor estPartidaDetenida
    'Detiene timer
    DesactivarTimer
   
    'Confirma la Pausa al ADM
    sConfirmarPausarPartida

    Exit Sub
ErrorHandle:
    ReportErr "sPausarPartida", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub
Public Sub sConfirmarPausarPartida()
    'Confirma al Adm que la partida se ha Pausado
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgAckPausarPartida), GintIndiceAdm

    Exit Sub
ErrorHandle:
    ReportErr "sConfirmarPausarPartida", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sContinuarPartida()
    'Continua la Partida, luego de la Pausa
    On Error GoTo ErrorHandle
    
    'Valida que la acción sea solicitada por el Administrador.
    If GintOrigenMensaje <> GintIndiceAdm Then
        sEnviarError CstrErrorNoAdm, errNoAccionNoAdm, GintOrigenMensaje
        Exit Sub
    End If
    'Valida que el Servidor se encuentre Pausado.
    If GEstadoServidor <> estPartidaDetenida Then
        sEnviarError "No se puede Continuar la Partida. El Servidor no se encuentra Pausado.", errNoContinuarNoPausa, GintIndiceAdm
        Exit Sub
    End If

    'Reactiva los Robots
    sInformarContinuacionPartida
    
   'Descongela el Servidor
    '###E
    'Cambia el estado del servidor
    CambiarEstadoServidor estEjecutandoPartida
    'Reactiva el Timer con el tiempo restante del Turno.
    ActivarTimer GintValorTimerActual
    
    'Confirma la Salida de la Pausa al ADM
    sConfirmarContinuarPartida

    Exit Sub
ErrorHandle:
    ReportErr "sContinuarPartida", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sConfirmarContinuarPartida()
    'Confirma al Adm que la partida se ha Continuado
    On Error GoTo ErrorHandle
    
    EnviarMensaje ArmarMensajeParam(msgAckContinuarPartida), GintIndiceAdm

    Exit Sub
ErrorHandle:
    ReportErr "sConfirmarContinuarPartida", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sInformarPausaPartida()
    'Informa al Robot de Turno que la partida se ha Pausado
    On Error GoTo ErrorHandle
    
    Dim strSQL As String
    Dim intIndiceActivo As Integer
    
    'Toma de la BD el jugador activo
    strSQL = "SELECT Par_Activo_Ronda " & _
             "FROM Partidas " & _
             "WHERE Par_Id = " & GintPartidaActiva
    intIndiceActivo = ColorAIndice(CInt(EjecutarConsultaValor(strSQL)))
    
    'Si no es el turno del Adm, le avisa al Robot de turno
    If intIndiceActivo <> GintIndiceAdm Then
        EnviarMensaje ArmarMensajeParam(msgPartidaPausada), intIndiceActivo
    End If

    Exit Sub
ErrorHandle:
    ReportErr "sInformarPausaPartida", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub

Public Sub sInformarContinuacionPartida()
    'Informa al Robot de Turno que la partida se ha Continuado
    On Error GoTo ErrorHandle
    
    Dim strSQL As String
    Dim intIndiceActivo As Integer
    
    'Toma de la BD el jugador activo
    strSQL = "SELECT Par_Activo_Ronda " & _
             "FROM Partidas " & _
             "WHERE Par_Id = " & GintPartidaActiva
    intIndiceActivo = ColorAIndice(CInt(EjecutarConsultaValor(strSQL)))
    
    'Si no es el turno del Adm, le avisa al Robot de turno
    If intIndiceActivo <> GintIndiceAdm Then
        EnviarMensaje ArmarMensajeParam(msgPartidaContinuada), intIndiceActivo
    End If

    Exit Sub
ErrorHandle:
    ReportErr "sInformarPausaPartida", "mdlServidor", Err.Description, Err.Number, Err.Source
End Sub
Public Function Tiempo() As String
    Tiempo = Format(Date, "dd/mm/yy ") & Format(Time, "hh:mm:ss.") & Format(Round((Timer - Fix(Timer)) * 100), "00")
End Function
