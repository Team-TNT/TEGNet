Attribute VB_Name = "mdlRegional"
'Este módulo contiene las declaraciones de constantes necesarias
'para el manejo del archivo de recursos que tiene la aplicacion
Option Explicit

'Idiomas
Public Const CintBaseSpanish As Integer = 0
Public Const CintBaseEnglish As Integer = 4000

'La ubicación de un objeto en un arhivo de recursos viene dado
'por una base (el idioma) y un desplazamiento (el objeto dentro
'del idioma)
Public GintBaseIdioma As Integer

'Los desplazamientos dentro de un idioma para cada opción vienen
'dados por las siguientes constantes

'PARAMETROS


'FORMULARIOS
Public Const CintAppNombre As Integer = 1001
Public Const CintAppSlogan As Integer = 1002
Public Const CintAppCaption As Integer = 1003
Public Const CintInicioCaption As Integer = 1010
Public Const CintInicioVersion As Integer = 1011
Public Const CintInicioCopyright As Integer = 1012
Public Const CintInicioArgentina As Integer = 1013

Public Const CintComienzoAtras As Integer = 1101
Public Const CintComienzoSiguiente As Integer = 1102
Public Const CintComienzoConectar As Integer = 1103
Public Const CintComienzoCancelar As Integer = 1104
Public Const CintComienzo1Principal As Integer = 1105
Public Const CintComienzo1Organizar As Integer = 1106
Public Const CintComienzo1Unirse As Integer = 1107
Public Const CintComienzo2Principal As Integer = 1108
Public Const CintComienzo2Servidor As Integer = 1109
Public Const CintComienzo2Puerto As Integer = 1110
Public Const CintComienzo3Nueva As Integer = 1111
Public Const CintComienzo3Guardada As Integer = 1112
Public Const CintComienzo3Avanzado As Integer = 1113
Public Const CintComienzo3Ocultar As Integer = 1114
Public Const CintComienzo3Local As Integer = 1115
Public Const CintComienzo3Remoto As Integer = 1116
Public Const CintComienzo3Servidor As Integer = 1117
Public Const CintComienzo3Puerto As Integer = 1118
Public Const CintComienzo3Activar As Integer = 1119
Public Const CintComienzo4Ultima As Integer = 1120
Public Const CintComienzo4Guardada As Integer = 1121
Public Const CintComienzoCaption As Integer = 1122
Public Const CintComienzo1Caption As Integer = 1123
Public Const CintComienzo2Caption As Integer = 1124
Public Const CintComienzo3Caption As Integer = 1125
Public Const CintComienzo4Caption As Integer = 1126
Public Const CintComienzo3Ubicacion As Integer = 1127
Public Const CintComienzoMsgDesconectar As Integer = 1128
Public Const CintComienzoMsgDesconectarCaption As Integer = 1129
Public Const CintComienzoAceptar As Integer = 1130
Public Const CintComienzoConectando As Integer = 1131
Public Const CintComienzoErrorConectando As Integer = 1132

'Formulario Principal
Public Const CintPrincipalCaption As Integer = 1200
Public Const CintPrincipalMnuPartida = 1201
Public Const CintPrincipalMnuVer As Integer = 1202
Public Const CintPrincipalMnuJuego As Integer = 1203
Public Const CintPrincipalMnuVentana As Integer = 1204
Public Const CintPrincipalMnuAdministracion As Integer = 1205
Public Const CintPrincipalMnuAyuda As Integer = 1206
Public Const CintPrincipalMnuPartidaConectar As Integer = 1207
Public Const CintPrincipalMnuPartidaDesconectar As Integer = 1208
Public Const CintPrincipalMnuPartidaResincronizar As Integer = 1209
Public Const CintPrincipalMnuPartidaGuardar As Integer = 1210
Public Const CintPrincipalMnuPartidaOpciones As Integer = 1211
Public Const CintPrincipalMnuPartidaSalir As Integer = 1212
Public Const CintPrincipalMnuVerMapa As Integer = 1213
Public Const CintPrincipalMnuVerJugadores As Integer = 1214
Public Const CintPrincipalMnuVerChat As Integer = 1215
Public Const CintPrincipalMnuVerSeleccion As Integer = 1216
Public Const CintPrincipalMnuVerDados As Integer = 1217
Public Const CintPrincipalMnuVerMision As Integer = 1218
Public Const CintPrincipalMnuVerTarjetas As Integer = 1219
Public Const CintPrincipalMnuVerTropas As Integer = 1220
Public Const CintPrincipalMnuVerInformacion As Integer = 1221
Public Const CintPrincipalMnuJuegoAgregar As Integer = 1222
Public Const CintPrincipalMnuJuegoAtacar As Integer = 1223
Public Const CintPrincipalMnuJuegoMover As Integer = 1224
Public Const CintPrincipalMnuJuegoTomar As Integer = 1225
Public Const CintPrincipalMnuJuegoFinTurno As Integer = 1226
Public Const CintPrincipalMnuVentanaOrganizar As Integer = 1227
Public Const CintPrincipalMnuAdministracionCambiar As Integer = 1228
Public Const CintPrincipalMnuAdministracionJV As Integer = 1229
Public Const CintPrincipalMnuAdministracionBajar As Integer = 1230
Public Const CintPrincipalMnuAyudaAcercaDe As Integer = 1231
Public Const CintPrincipalTipConectar As Integer = 1232
Public Const CintPrincipalTipResincronizar As Integer = 1233
Public Const CintPrincipalTipOpciones As Integer = 1234
Public Const CintPrincipalTipGuardar As Integer = 1235
Public Const CintPrincipalTipPausar As Integer = 1339
Public Const CintPrincipalTipContinuar As Integer = 1340
Public Const CintPrincipalTipVerMision As Integer = 1236
Public Const CintPrincipalTipVerTarjetas As Integer = 1237
Public Const CintPrincipalTipAgregar As Integer = 1238
Public Const CintPrincipalTipAtacar As Integer = 1239
Public Const CintPrincipalTipMover As Integer = 1240
Public Const CintPrincipalTipTomar As Integer = 1241
Public Const CintPrincipalTipFinTurno As Integer = 1242
Public Const CintPrincipalStatusIntercambio As Integer = 1243
Public Const CintPrincipalStatusMiColor As Integer = 1244
Public Const CintPrincipalStatusTipoRonda As Integer = 1245
Public Const CintPrincipalStatusEstadoJuego As Integer = 1246
Public Const CintPrincipalStatusAdministracion As Integer = 1247
Public Const CintPrincipalStatusTiempoTurno As Integer = 1248
Public Const CintPrincipalEstadoDesconectado As Integer = 1249
Public Const CintPrincipalEstadoConectado As Integer = 1250
Public Const CintPrincipalEstadoValidado As Integer = 1251
Public Const CintPrincipalEstadoEsperandoTurno As Integer = 1252
Public Const CintPrincipalEstadoAgregando As Integer = 1253
Public Const CintPrincipalEstadoAtacando As Integer = 1254
Public Const CintPrincipalEstadoMoviendo As Integer = 1255
Public Const CintPrincipalEstadoTarTomada As Integer = 1256
Public Const CintPrincipalEstadoTarCobrada As Integer = 1257
Public Const CintPrincipalEstadoTarTomadaCobrada As Integer = 1258
Public Const CintPrincipalEstadoPausada As Integer = 1341
Public Const CintPrincipalEstadoFinalizada As Integer = 1259
Public Const CintPrincipalEstadoInconsistente As Integer = 1260
Public Const CintPrincipalMsgAbandonar As Integer = 1261
Public Const CintPrincipalMsgAbandonarCaption As Integer = 1262
Public Const CintPrincipalMsgDesconectar As Integer = 1263
Public Const CintPrincipalMsgDesconectarCaption As Integer = 1264
Public Const CintPrincipalMsgServidor As Integer = 1265
Public Const CintPrincipalMsgServidorCaption As Integer = 1266
Public Const CintPrincipalMsgConexion As Integer = 1267
Public Const CintPrincipalMsgConexionCaption As Integer = 1268
Public Const CintPrincipalMsgConexion2 As Integer = 1269
Public Const CintPrincipalMsgConexion2Caption As Integer = 1270
Public Const CintPrincipalMsgOtros As Integer = 1271
Public Const CintPrincipalMsgOtrosCaption As Integer = 1272
Public Const CintPrincipalRondaInicio As Integer = 1273
Public Const CintPrincipalRondaAccion As Integer = 1274
Public Const CintPrincipalRondaRecuento As Integer = 1275
Public Const CintPrincipalMsgBajarServidor As Integer = 1276
Public Const CintPrincipalMsgBajarServidorCaption As Integer = 1277
Public Const CintPrincipalMnuVerLog As Integer = 1278
Public Const CintPrincipalMnuVerListaMisiones As Integer = 1279
Public Const CintPrincipalMnuPartidaIdioma As Integer = 1280
Public Const CintPrincipalMnuPopupAtacar As Integer = 1281
Public Const CintPrincipalMnuPopupMover1 As Integer = 1282
Public Const CintPrincipalMnuPopupMoverTodas As Integer = 1283
Public Const CintPrincipalMnuPopupAgregar1 As Integer = 1284
Public Const CintPrincipalMnuPopupAgregar5 As Integer = 1285
Public Const CintPrincipalMnuPopupAgregar10 As Integer = 1286
Public Const CintPrincipalMnuPopupAgregarTodas As Integer = 1287
Public Const CintPrincipalMsgServidorCerrado As Integer = 1288
Public Const CintPrincipalMsgCerrarServidor As Integer = 1289
Public Const CintPrincipalMsgCerrarServidorCaption As Integer = 1290


Public Const CintOpcionesCaption As Integer = 1300
Public Const CintOpcionesRecuperarDefecto As Integer = 1301
Public Const CintOpcionesGuardarDefecto As Integer = 1302
Public Const CintOpcionesAceptar As Integer = 1303
Public Const CintOpcionesCancelar As Integer = 1304
Public Const CintOpcionesTurno As Integer = 1305
Public Const CintOpcionesTurnoDuracion As Integer = 1306
Public Const CintOpcionesTurnoTolerancia As Integer = 1307
Public Const CintOpcionesRonda As Integer = 1308
Public Const CintOpcionesRondaTropasPrimera As Integer = 1309
Public Const CintOpcionesRondaTropasSegunda As Integer = 1310
Public Const CintOpcionesRondaTipo As Integer = 1311
Public Const CintOpcionesRondaFija As Integer = 1312
Public Const CintOpcionesRondaRotativa As Integer = 1313
Public Const CintOpcionesBonus As Integer = 1314
Public Const CintOpcionesBonusPaisPropio As Integer = 1315
Public Const CintOpcionesBonusContinente As Integer = 1316
Public Const CintOpcionesBonusAfrica As Integer = 1317
Public Const CintOpcionesBonusANorte As Integer = 1318
Public Const CintOpcionesBonusASur As Integer = 1319
Public Const CintOpcionesBonusAsia As Integer = 1320
Public Const CintOpcionesBonusEuropa As Integer = 1321
Public Const CintOpcionesBonusOceania As Integer = 1322
Public Const CintOpcionesCanje As Integer = 1323
Public Const CintOpcionesCanjePrimero As Integer = 1324
Public Const CintOpcionesCanjeSegundo As Integer = 1325
Public Const CintOpcionesCanjeTercero As Integer = 1326
Public Const CintOpcionesCanjeIncremento As Integer = 1327
Public Const CintOpcionesMision As Integer = 1328
Public Const CintOpcionesMisionConquistarMundo As Integer = 1329
Public Const CintOpcionesMisionMisiones As Integer = 1330
Public Const CintOpcionesMisionDestruir As Integer = 1331
Public Const CintOpcionesMisionObjetivoComun As Integer = 1332
Public Const CintOpcionesMisionPaises As Integer = 1333
Public Const CintOpcionesOtras As Integer = 1334
Public Const CintOpcionesOtrasRepartoInicial As Integer = 1335
Public Const CintOpcionesMsgDesconectar As Integer = 1336
Public Const CintOpcionesMsgDesconectarCaption As Integer = 1337
Public Const CintOpcionesMsgErrorCaption As Integer = 1338

Public Const CintGralMsgCaracterInvalido As Integer = 1400
Public Const CintGralMsgNumeroInvalido As Integer = 1401
Public Const CintGralTextoAdministrador As Integer = 1402
Public Const CintGralMsgColorAsignado As Integer = 1403
Public Const CintGralMsgColorAsignadoCaption As Integer = 1404
Public Const CintGralMsgNombreAsignado As Integer = 1405
Public Const CintGralMsgNombreAsignadoCaption As Integer = 1406
Public Const CintGralMsgNombreColorAsignados As Integer = 1407
Public Const CintGralMsgNombreColorAsignadosCaption As Integer = 1408
Public Const CintGralMsgColorInvalido As Integer = 1409
Public Const CintGralMsgColorInvalidoCaption As Integer = 1410
Public Const CintGralMsgColorUtilizado As Integer = 1411
Public Const CintGralMsgColorUtilizadoCaption As Integer = 1412
Public Const CintGralMsgColorNombreOtroCaption As Integer = 1413
Public Const CintGralMsgAdmDesignado As Integer = 1414
Public Const CintGralMsgAdmDesignadoCaption As Integer = 1415
Public Const CintGralMsgErrServidorCaption As Integer = 1416
Public Const CintGralMsgAdmNoDesignado As Integer = 1417
Public Const CintGralMsgAdmNoDesignadoCaption As Integer = 1418
Public Const CintGralMsgTurnoExpirado As Integer = 1419
Public Const CintGralMsgTurnoExpiradoCaption As Integer = 1420
Public Const CintGralMsgTropasMover As Integer = 1421
Public Const CintGralMsgTropasMoverCaption As Integer = 1422
Public Const CintGralMsgMisionCumplida As Integer = 1423
Public Const CintGralMsgErrGuardarPartidaNoAdm As Integer = 1424
Public Const CintGralMsgErrGuardarPartidaNoAdmCaption As Integer = 1425
Public Const CintGralMsgErrGuardarPartidaDesconocido As Integer = 1426
Public Const CintGralMsgErrGuardarPartidaDesconocidoCaption As Integer = 1427
Public Const CintGralMsgErrRutina As Integer = 1428
Public Const CintGralMsgErrModulo As Integer = 1429
Public Const CintGralMsgErrDescripcion As Integer = 1430
Public Const CintGralMsgErrOrigen As Integer = 1431
Public Const CintGralMsgErrNumero As Integer = 1432
Public Const CintGralMsgErrRutinaError As Integer = 1433
Public Const CintGralMsgCaracterInvalidoCaption As Integer = 1434
Public Const CintGralMSgNumeroInvalidoCaption As Integer = 1435
Public Const CintGralMsgServidorPausado As Integer = 1436
Public Const CintGralMsgServidorPausadoCaption As Integer = 1437

Public Const CintSelColorCaption As Integer = 1500
Public Const CintSelColorAceptar As Integer = 1501
Public Const CintSelColorCancelar As Integer = 1502
Public Const CintSelColorJV As Integer = 1503
Public Const CintSelColorIniciarPartida As Integer = 1504
Public Const CintSelColorlblNombre As Integer = 1505
Public Const CintSelColorValidando As Integer = 1506
Public Const CintSelColorColorNoSeleccionado As Integer = 1507
Public Const CintSelColorNombreNoIngresado As Integer = 1508
Public Const CintSelColorConfirmaCancelar As Integer = 1509
Public Const CintSelColorConfirmaCancelarCaption As Integer = 1510
Public Const CintSelColorConfirmaIniciar As Integer = 1511
Public Const CintSelColorConfirmaIniciarCaption As Integer = 1512
Public Const CintSelColorCaptionEst0 As Integer = 1513
Public Const CintSelColorCaptionEst1 As Integer = 1514
Public Const CintSelColorStatusEst1Adm As Integer = 1515
Public Const CintSelColorStatusEst1NoAdm As Integer = 1516
Public Const CintSelColorCaptionEst2 As Integer = 1517
Public Const CintSelColorStatusEst2 As Integer = 1518
Public Const CintSelColorDisponible As Integer = 1519
Public Const CintSelColorNoDisponible As Integer = 1520

Public Const CintCreditosCaption As Integer = 1550
Public Const CintCreditosAceptar As Integer = 1551
Public Const CintCreditosSlogan As Integer = 1552
Public Const CintCreditosTitCreditos As Integer = 1553
Public Const CintCreditosDesarrolladoPor As Integer = 1554
Public Const CintCreditosVisitenos As Integer = 1555
Public Const CintCreditosEMail As Integer = 1556
Public Const CintCreditosTitAgradecimientos As Integer = 1557
Public Const CintCreditosAgradecimiento1 As Integer = 1558
Public Const CintCreditosAgradecimiento2 As Integer = 1559
Public Const CintCreditosAgradecimiento3 As Integer = 1560
Public Const CintCreditosAgradecimiento4 As Integer = 1561
Public Const CintCreditosAgradecimiento5 As Integer = 1562
Public Const CintCreditosAgradecimiento6 As Integer = 1563
Public Const CintCreditosAgradecimiento7 As Integer = 1564
Public Const CintCreditosAgradecimiento8 As Integer = 1565
Public Const CintCreditosAgradecimiento9 As Integer = 1569
Public Const CintCreditosAgradecimiento10 As Integer = 1570
Public Const CintCreditosAgradecimiento11 As Integer = 1571
Public Const CintCreditosAgradecimiento12 As Integer = 1572
Public Const CintCreditosAgradecimientoFinal As Integer = 1580
Public Const CintCreditosTitCopyright As Integer = 1566
Public Const CintCreditosCopyright As Integer = 1567
Public Const CintCreditosDisclaimer As Integer = 1568

Public Const CintJVCaption As Integer = 1600
Public Const CintJVbtnAceptar As Integer = 1601
Public Const CintJVbtnCancelar As Integer = 1602
Public Const CintJVGeneral As Integer = 1603
Public Const CintJVPerfil As Integer = 1604
Public Const CintJVNombre As Integer = 1605
Public Const CintJVAleatorio As Integer = 1606
Public Const CintJVNivelAgresividad As Integer = 1607
Public Const CintJVMuyAgresivo As Integer = 1608
Public Const CintJVPocoAgresivo As Integer = 1609
Public Const CintJVActitud As Integer = 1610
Public Const CintJVArriesgado As Integer = 1611
Public Const CintJVConservador As Integer = 1612
Public Const CintJVCaptionEst0 As Integer = 1613
Public Const CintJVCaptionEst1 As Integer = 1614
Public Const CintJVDisponible As Integer = 1615
Public Const CintJVNoDisponible As Integer = 1616

Public Const CintCambioAdmCaption As Integer = 1650
Public Const CintCambioAdmFrame As Integer = 1651
Public Const CintCambioAdmBtnAceptar As Integer = 1652
Public Const CintCambioAdmBtnCancelar As Integer = 1653
Public Const CintCambioAdmErrSeleccionJugador As Integer = 1654
Public Const CintCambioAdmErrSeleccionJugadorCaption As Integer = 1655
Public Const CintCambioAdmNoDisponible As Integer = 1656

Public Const CintChatCaption As Integer = 1660
Public Const CintChatEnviar As Integer = 1661
Public Const CintChatMostrarLog As Integer = 1662

Public Const CintDadosCaption As Integer = 1670
Public Const CintDadosAtaque As Integer = 1671
Public Const CintDadosDefensa As Integer = 1672

Public Const CintGuardarCaption As Integer = 1680
Public Const CintGuardarTitulo As Integer = 1681
Public Const CintGuardarEliminar As Integer = 1682
Public Const CintGuardarNombrePartida As Integer = 1683
Public Const CintGuardarGuardar As Integer = 1684
Public Const CintGuardarCancelar As Integer = 1685

Public Const CintJugadoresCaption As Integer = 1690
Public Const CintJugadoresRonda As Integer = 1691
Public Const CintJugadoresDetalle As Integer = 1692
Public Const CintJugadoresPaises As Integer = 1693
Public Const CintJugadoresTropas As Integer = 1694
Public Const CintJugadoresTropasDisponibles As Integer = 1695
Public Const CintJugadoresTarjetas As Integer = 1696
Public Const CintJugadoresCanjes As Integer = 1697
Public Const CintJugadoresDetalleDe As Integer = 1698

Public Const CintLogCaption As Integer = 1700

Public Const CintMensajeAceptar As Integer = 1710

Public Const CintMisionCaption As Integer = 1720

Public Const CintMisionCumplidaCaption As Integer = 1730
Public Const CintMisionCumplidaAceptar As Integer = 1731

Public Const CintInformacionCaption As Integer = 1740
Public Const CintInformacionDuenio As Integer = 1741
Public Const CintInformacionTropas As Integer = 1742
Public Const CintInformacionFijas As Integer = 1743
Public Const CintInformacionCaptionDinamica As Integer = 1744

Public Const CintSeleccionCaption As Integer = 1750
Public Const CintSeleccionDesde As Integer = 1751
Public Const CintSeleccionHasta As Integer = 1752

Public Const CintTarjetasCaption As Integer = 1760
Public Const CintTarjetasCobrar As Integer = 1761
Public Const CintTarjetasCanjear As Integer = 1762
Public Const CintTarjetasCobrada As Integer = 1763

Public Const CintTropasCaption As Integer = 1770
Public Const CintTropasDetalle As Integer = 1771
Public Const CintTropasLibres As Integer = 1772
Public Const CintTropasAfrica As Integer = 1773
Public Const CintTropasANorte As Integer = 1774
Public Const CintTropasASur As Integer = 1775
Public Const CintTropasAsia As Integer = 1776
Public Const CintTropasEuropa As Integer = 1777
Public Const CintTropasOceania As Integer = 1778

Public Const CintMisionesCaption As Integer = 1780

Public Const CintConquistaUnaTropa As Integer = 1790
Public Const CintConquistaDosTropas As Integer = 1791
Public Const CintConquistaTresTropas As Integer = 1792

Public Const CintMapaCaption As Integer = 1800

Public Const CintIdiomaMsgCambio As Integer = 1810
Public Const CintIdiomaMsgCambioCaption As Integer = 1811

Public Function ObtenerTextoRecurso(intIdRecurso As Integer) As String
    On Error GoTo ErrorHandle
    'Devuelve el texto asociado al objeto de archivo de recurso
    'especificado
    ObtenerTextoRecurso = LoadResString(intIdRecurso + GintBaseIdioma)
    Exit Function
ErrorHandle:
    ObtenerTextoRecurso = ""
End Function
