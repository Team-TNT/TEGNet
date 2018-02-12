VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSeleccionColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione un color"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSeleccionColor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdIniciarPartida 
      Cancel          =   -1  'True
      Caption         =   "&Iniciar Partida"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2490
      TabIndex        =   8
      Top             =   2250
      Width           =   1470
   End
   Begin VB.CommandButton cmdCargarJV 
      Caption         =   "&Jugador Virtual..."
      Height          =   345
      Left            =   645
      TabIndex        =   7
      Top             =   2250
      Width           =   1470
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   13
      Top             =   2670
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraSeleccionColor 
      Height          =   2085
      Left            =   45
      TabIndex        =   10
      Top             =   -15
      Width           =   4605
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   6
         Left            =   2925
         TabIndex        =   5
         Top             =   1230
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   660
         TabIndex        =   0
         Top             =   330
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   660
         TabIndex        =   1
         Top             =   780
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   660
         TabIndex        =   2
         Top             =   1230
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   4
         Left            =   2925
         TabIndex        =   3
         Top             =   330
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   5
         Left            =   2925
         TabIndex        =   4
         Top             =   780
         Width           =   1635
      End
      Begin VB.TextBox txtNickName 
         Height          =   300
         Left            =   1215
         MaxLength       =   20
         TabIndex        =   6
         ToolTipText     =   "Ingrese su nombre de jugador"
         Top             =   1650
         Width           =   2715
      End
      Begin VB.Label lblNickName 
         Caption         =   "Nombre:"
         Height          =   270
         Left            =   480
         TabIndex        =   12
         Top             =   1680
         Width           =   690
      End
      Begin VB.Image imgFicha 
         Height          =   420
         Index           =   2
         Left            =   165
         Stretch         =   -1  'True
         Top             =   675
         Width           =   420
      End
      Begin VB.Image imgFicha 
         Height          =   420
         Index           =   3
         Left            =   165
         Stretch         =   -1  'True
         Top             =   1125
         Width           =   420
      End
      Begin VB.Image imgFicha 
         Height          =   420
         Index           =   5
         Left            =   2460
         Stretch         =   -1  'True
         Top             =   675
         Width           =   420
      End
      Begin VB.Image imgFicha 
         Height          =   420
         Index           =   6
         Left            =   2460
         Stretch         =   -1  'True
         Top             =   1125
         Width           =   420
      End
      Begin VB.Image imgFicha 
         Height          =   420
         Index           =   4
         Left            =   2460
         Stretch         =   -1  'True
         Top             =   225
         Width           =   420
      End
      Begin VB.Image imgFicha 
         Height          =   420
         Index           =   1
         Left            =   165
         Stretch         =   -1  'True
         Top             =   225
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   990
      TabIndex        =   9
      Top             =   2205
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   2400
      TabIndex        =   11
      Top             =   2205
      Width           =   1230
   End
End
Attribute VB_Name = "frmSeleccionColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnSalioPorIniciarPartida As Boolean
Dim intTipoJugadoresConectados As enuTipoJugadoresConectados

Private Sub cmdAceptar_Click()
    On Error GoTo ErrorHandle
    
    Dim intColorSeleccionado As Integer
    Dim i As Integer
    
    'Valida el texto ingresado
    If Not ValidaTexto(txtNickName.Text, 20) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgCaracterInvalido), vbInformation, ObtenerTextoRecurso(CintGralMsgCaracterInvalidoCaption)
        Exit Sub
    End If
    
    StatusBar1.SimpleText = ObtenerTextoRecurso(CintSelColorValidando) '"Validando selección..."
    cmdAceptar.Enabled = False
    
    For i = 1 To optColor.Count
        If optColor(i).Enabled = True Then
            If optColor(i).Value = True Then intColorSeleccionado = i
        End If
    Next i
    
    If intColorSeleccionado = 0 Then
        MsgBox ObtenerTextoRecurso(CintSelColorColorNoSeleccionado) '"Debe seleccionar un color"
        Exit Sub
    End If
    
    If Trim(txtNickName.Text) = "" Then
        MsgBox ObtenerTextoRecurso(CintSelColorNombreNoIngresado) '"Debe ingresar un Nombre"
        If txtNickName.Enabled Then txtNickName.SetFocus
        Exit Sub
    End If
    
    If intTipoJugadoresConectados = tjcIngreso Then
        cAltaJugador intColorSeleccionado, Trim(txtNickName.Text)
    Else
        cReconectarJugador intColorSeleccionado
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdAceptar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCargarJV_Click()
    On Error GoTo ErrorHandle
    
    MostrarFormulario frmJugadorVirtual, vbModal
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdIniciarPartida_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdIniciarPartida_Click()
    On Error GoTo ErrorHandle
    
    blnSalioPorIniciarPartida = True
    cIniciarPartida
    Unload Me

    Exit Sub
ErrorHandle:
    ReportErr "cmdIniciarPartida_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    'Regional
    Me.Caption = ObtenerTextoRecurso(CintSelColorCaption)
    Me.cmdAceptar.Caption = ObtenerTextoRecurso(CintSelColorAceptar)
    Me.cmdCancelar.Caption = ObtenerTextoRecurso(CintSelColorCancelar)
    Me.cmdCargarJV.Caption = ObtenerTextoRecurso(CintSelColorJV)
    Me.cmdIniciarPartida.Caption = ObtenerTextoRecurso(CintSelColorIniciarPartida)
    Me.lblNickName.Caption = ObtenerTextoRecurso(CintSelColorlblNombre)
    
    cmdAceptar.Enabled = False
    blnSalioPorIniciarPartida = False
    
    'Si ya hay un color selccionado no deja elegir otro
'    If GintMiColor <> 0 Then
'        CambiarEstadoPantalla 1
'    Else
'        CambiarEstadoPantalla 0
'    End If
    
    For i = 1 To optColor.Count
        imgFicha(i).Picture = mdifrmPrincipal.imgLstFichas.ListImages(i).Picture
'        optColor(i).BackColor = GvecColores(i)
'        optColor(i).ForeColor = GvecColoresInv(i)
    Next
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Load", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    
    If blnSalioPorIniciarPartida Then
        Exit Sub
    End If
    
    If GblnSeCierra Then
        Exit Sub
    End If
    
    'Si todavía no seleccionó el color se desconecta
    If GEstadoCliente < estValidado Then
        '¿Está seguro que desea cancelar la partida?
        If MsgBox(ObtenerTextoRecurso(CintSelColorConfirmaCancelar), vbQuestion + vbYesNo + vbDefaultButton2, ObtenerTextoRecurso(CintSelColorConfirmaCancelarCaption)) = vbNo Then
            Cancel = 1
        Else
            If GsoyAdministrador Then
                'Baja el servidor
                cBajarServidor
            Else
                cDesconectar
            End If
        End If
    Else
        '### Para evitar que pregunte si se volvió a cargar accidentalmente
        If GEstadoCliente < estEsperandoTurno Then
            If GsoyAdministrador Then
                '¿Desea iniciar la partida?
                If MsgBox(ObtenerTextoRecurso(CintSelColorConfirmaIniciar), vbQuestion + vbYesNo + vbDefaultButton1, ObtenerTextoRecurso(CintSelColorConfirmaIniciarCaption)) = vbYes Then
                    'Inicia la partida
                    cIniciarPartida
                Else
                    '¿Está seguro que desea cancelar la partida?
                    If MsgBox(ObtenerTextoRecurso(CintSelColorConfirmaCancelar), vbQuestion + vbYesNo + vbDefaultButton2, ObtenerTextoRecurso(CintSelColorConfirmaCancelarCaption)) = vbNo Then
                        Cancel = 1
                    Else
                        'Baja el servidor
                        cBajarServidor
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Unload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub imgFicha_Click(Index As Integer)
    If optColor(Index).Enabled Then optColor(Index).Value = True
End Sub

Public Sub CambiarEstadoPantalla(intNuevoEstado As Integer)
    'Cambia el estado de la pantalla de acuerdo a si está o no seleccionado el color local
    On Error GoTo ErrorHandle
    Select Case intNuevoEstado
        Case 0 'No esta el color seleccionado
            Me.Caption = ObtenerTextoRecurso(CintSelColorCaptionEst0) '"Seleccione un Color y un Nombre"
            StatusBar1.SimpleText = ""
            Me.txtNickName.Visible = True
            Me.lblNickName.Visible = True
            Me.fraSeleccionColor.Enabled = True
            Me.cmdAceptar.Visible = True
            Me.cmdCancelar.Visible = True
            cmdIniciarPartida.Visible = False
            cmdCargarJV.Visible = False
            intTipoJugadoresConectados = tjcIngreso
        Case 1 'Existe un color seleccionado
            Me.Caption = ObtenerTextoRecurso(CintSelColorCaptionEst1) '"Jugadores Conectados"
            If GsoyAdministrador Then
                StatusBar1.SimpleText = ObtenerTextoRecurso(CintSelColorStatusEst1Adm) '"Esperando Nuevos Jugadores... [Servidor:"
                cmdIniciarPartida.Visible = True
                cmdCargarJV.Visible = True
            Else
                StatusBar1.SimpleText = ObtenerTextoRecurso(CintSelColorStatusEst1NoAdm) '"Esperando Inicio de la Partida..."
                cmdIniciarPartida.Visible = False
                cmdCargarJV.Visible = False
            End If
            Me.txtNickName.Visible = False
            Me.lblNickName.Visible = False
            Me.fraSeleccionColor.Enabled = False
            Me.cmdAceptar.Visible = False
            Me.cmdCancelar.Visible = False
            intTipoJugadoresConectados = tjcIngreso
        Case 2 'Reconexion
            Me.Caption = ObtenerTextoRecurso(CintSelColorCaptionEst2) '"Seleccione un Jugador para reconectarse"
            StatusBar1.SimpleText = ObtenerTextoRecurso(CintSelColorStatusEst2) '"Reconexión..."
            Me.txtNickName.Visible = False
            Me.lblNickName.Visible = False
            Me.fraSeleccionColor.Enabled = True
            Me.cmdAceptar.Visible = True
            Me.cmdCancelar.Visible = True
            cmdIniciarPartida.Visible = False
            cmdCargarJV.Visible = False
            intTipoJugadoresConectados = tjcReconexion
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "CambiarEstadoPantalla", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub optColor_Click(Index As Integer)
    On Error GoTo ErrorHandle
    
    'Si se esta reconectando, copia el nombre de la opcion al textbox
    'sino le pasa el foco para que ingrese el nombre
    If intTipoJugadoresConectados = tjcReconexion Then
        txtNickName.Text = IIf(Left(optColor(Index).Caption, 1) = chrPREFIJOADM, Mid(optColor(Index).Caption, 2), optColor(Index).Caption)
    Else
        txtNickName.SetFocus
    End If
    
    If Trim(txtNickName.Text) <> "" Then
        cmdAceptar.Enabled = True
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "optColor_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtNickName_Change()
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    Dim aux As Integer
    
    aux = 0
    
    For i = 1 To optColor.Count
        If optColor(i).Enabled = True Then
            aux = aux + optColor(i).Value
        End If
    Next i
    
    If Trim(txtNickName.Text) <> "" And aux <> 0 Then
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "txtNickName_Change", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtNickName_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    
    If Not ValidaTexto(Chr$(KeyAscii), 0) Then
        
        KeyAscii = 0
        Beep
    
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "txtNickName_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub Actualizar(intTipoJugadoresConectados As enuTipoJugadoresConectados)
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    'Reccorre el vector de jugadores
    For i = 1 To UBound(GvecJugadores)
        
        If intTipoJugadoresConectados = tjcIngreso Then
            'INGRESO
            If GvecJugadores(i).intEstado = conConectado And GvecJugadores(i).strNombre <> "" Then
                'Ingreso (deshabilita los que ya ingresaron)
                optColor(i).Caption = GvecJugadores(i).strNombre
                optColor(i).Enabled = False
            Else
                'Ingreso (habilita los disponibles)
                optColor(i).Caption = ObtenerTextoRecurso(CintSelColorDisponible) '"Disponible"
                optColor(i).Enabled = True
            End If
        Else
            'RECONEXION
            If GvecJugadores(i).intEstado = conNoJuega Then
                'Reconexion (deshabilita los que no juegan)
                optColor(i).Caption = ObtenerTextoRecurso(CintSelColorNoDisponible) '"No disponible"
                optColor(i).Enabled = False
            Else
                'Reconexion (habilita los que estan jugando)
                optColor(i).Caption = GvecJugadores(i).strNombre
                optColor(i).Enabled = True
            End If
        End If
    Next i
        
    'Si no está jugando muestra los colores o la pantalla de reconexion
    If GEstadoCliente < estEsperandoTurno Then
        If Not Me.Visible Then
            If GintMiColor <> 0 Then
                'Si ya estoy validado
                CambiarEstadoPantalla 1
            Else
                'Si todavía no estoy validado
                If intTipoJugadoresConectados = tjcIngreso Then
                    'Estoy ingresando por primera vez
                    CambiarEstadoPantalla 0
                Else
                    'Estoy reconectandome
                    CambiarEstadoPantalla 2
                End If
            End If
            MostrarFormulario Me, vbModal
        End If
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "Actualizar", Me.Name, Err.Description, Err.Number, Err.Source
End Sub
