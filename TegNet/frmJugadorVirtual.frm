VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJugadorVirtual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activar Jugador Virtual"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmJugadorVirtual.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   3975
      TabIndex        =   11
      Top             =   3000
      Width           =   1230
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   330
      Left            =   2565
      TabIndex        =   10
      Top             =   3000
      Width           =   1230
   End
   Begin VB.Frame fraSolapas 
      Height          =   2220
      Index           =   0
      Left            =   195
      TabIndex        =   13
      Top             =   480
      Width           =   4845
      Begin VB.TextBox txtNickName 
         Height          =   300
         Left            =   1215
         MaxLength       =   20
         TabIndex        =   6
         ToolTipText     =   "Ingrese su nombre de jugador"
         Top             =   1755
         Width           =   2715
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   5
         Left            =   3045
         TabIndex        =   4
         Top             =   780
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   735
         TabIndex        =   2
         Top             =   1230
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   735
         TabIndex        =   1
         Top             =   780
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   6
         Left            =   3045
         TabIndex        =   5
         Top             =   1230
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   735
         TabIndex        =   0
         Top             =   330
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   4
         Left            =   3045
         TabIndex        =   3
         Top             =   330
         Width           =   1635
      End
      Begin VB.Image imgFicha 
         Height          =   420
         Index           =   1
         Left            =   255
         Stretch         =   -1  'True
         Top             =   225
         Width           =   420
      End
      Begin VB.Image imgFicha 
         Height          =   420
         Index           =   4
         Left            =   2580
         Stretch         =   -1  'True
         Top             =   225
         Width           =   420
      End
      Begin VB.Image imgFicha 
         Height          =   420
         Index           =   6
         Left            =   2580
         Stretch         =   -1  'True
         Top             =   1125
         Width           =   420
      End
      Begin VB.Image imgFicha 
         Height          =   420
         Index           =   5
         Left            =   2580
         Stretch         =   -1  'True
         Top             =   675
         Width           =   420
      End
      Begin VB.Image imgFicha 
         Height          =   420
         Index           =   3
         Left            =   255
         Stretch         =   -1  'True
         Top             =   1125
         Width           =   420
      End
      Begin VB.Image imgFicha 
         Height          =   420
         Index           =   2
         Left            =   255
         Stretch         =   -1  'True
         Top             =   675
         Width           =   420
      End
      Begin VB.Label lblNickName 
         Caption         =   "Nombre:"
         Height          =   270
         Left            =   480
         TabIndex        =   14
         Top             =   1785
         Width           =   690
      End
   End
   Begin VB.Frame fraSolapas 
      Height          =   2220
      Index           =   1
      Left            =   195
      TabIndex        =   15
      Top             =   480
      Width           =   4845
      Begin MSComctlLib.Slider SldActitud 
         Height          =   1350
         Left            =   2640
         TabIndex        =   9
         Top             =   795
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2381
         _Version        =   393216
         Orientation     =   1
         Max             =   6
         SelStart        =   3
         Value           =   3
      End
      Begin MSComctlLib.Slider SldAgresividad 
         Height          =   1350
         Left            =   480
         TabIndex        =   8
         Top             =   795
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2381
         _Version        =   393216
         Orientation     =   1
         Max             =   6
         SelStart        =   3
         Value           =   3
      End
      Begin VB.CheckBox chkAleatorio 
         Caption         =   "Perfil Aleatorio"
         Height          =   225
         Left            =   630
         TabIndex        =   7
         Top             =   210
         Width           =   2040
      End
      Begin VB.Label lblConservador 
         Caption         =   "Conservador"
         Height          =   255
         Left            =   3090
         TabIndex        =   21
         Top             =   1890
         Width           =   1365
      End
      Begin VB.Label lblArriesgado 
         Caption         =   "Arriesgado"
         Height          =   255
         Left            =   3090
         TabIndex        =   20
         Top             =   870
         Width           =   1365
      End
      Begin VB.Label lblActitud 
         Caption         =   "Actitud:"
         Height          =   270
         Left            =   2640
         TabIndex        =   19
         Top             =   570
         Width           =   1590
      End
      Begin VB.Label lblPocoAgresivo 
         Caption         =   "Poco Agresivo"
         Height          =   255
         Left            =   930
         TabIndex        =   18
         Top             =   1890
         Width           =   1365
      End
      Begin VB.Label lblMuyAgresivo 
         Caption         =   "Muy Agresivo"
         Height          =   255
         Left            =   930
         TabIndex        =   17
         Top             =   870
         Width           =   1365
      End
      Begin VB.Label lblAgresividad 
         Caption         =   "Nivel de Agresividad:"
         Height          =   270
         Left            =   495
         TabIndex        =   16
         Top             =   570
         Width           =   1590
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2730
      Left            =   90
      TabIndex        =   12
      Top             =   105
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   4815
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedWidth   =   1764
      TabMinWidth     =   354
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Object.ToolTipText     =   "Selección de Nombre y Color"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Perfil"
            Object.ToolTipText     =   "Perfil del Jugador Virtual"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmJugadorVirtual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CambiarEstadoPantalla(intNuevoEstado As Integer)
    'Cambia el estado de la pantalla de acuerdo a si
    'se está dando de alta un JV o se está reconectando/reemplazando
    On Error GoTo ErrorHandle
    Select Case intNuevoEstado
        Case 0 'Alta
            Me.Caption = ObtenerTextoRecurso(CintJVCaptionEst0) '"Activar Jugador Virtual"
            Me.txtNickName.Visible = True
            Me.lblNickName.Visible = True
        Case 1 'Reconexion/Reemplazo
            Me.Caption = ObtenerTextoRecurso(CintJVCaptionEst1) '"Asignar Jugador Virtual"
            Me.txtNickName.Visible = False
            Me.lblNickName.Visible = False
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "CambiarEstadoPantalla", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub chkAleatorio_Click()
    On Error GoTo ErrorHandle
    
    HabilitarPerfil IIf(chkAleatorio.Value = 1, False, True)

    Exit Sub
ErrorHandle:
    ReportErr "chkAleatorio_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub HabilitarPerfil(blnValor As Boolean)
    On Error GoTo ErrorHandle

    'Habilita o deshabilita los controles del perfil
    lblAgresividad.Enabled = blnValor
    lblActitud.Enabled = blnValor
    SldAgresividad.Enabled = blnValor
    SldActitud.Enabled = blnValor
    lblMuyAgresivo.Enabled = blnValor
    lblPocoAgresivo.Enabled = blnValor
    lblConservador.Enabled = blnValor
    lblArriesgado.Enabled = blnValor

    Exit Sub
ErrorHandle:
    ReportErr "HabilitarPerfil", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdAceptar_Click()
    On Error GoTo ErrorHandle
    
    Dim intColorSeleccionado As Integer
    Dim i As Integer
    
    Dim intAgresividad As Integer
    Dim intActitud As Integer
    
    'Obtiene los parametros del perfil del JV
    If chkAleatorio.Value = 1 Then
        intAgresividad = Aleatorio(0, 90)
        intActitud = Aleatorio(0, 100)
    Else
        intAgresividad = SldAgresividad.Value * (100 / SldAgresividad.Max)
        intActitud = 100 - SldActitud.Value * (100 / SldActitud.Max)
    End If
    
    'Busca el color seleccionado
    For i = 1 To optColor.Count
        If optColor(i).Enabled = True Then
            If optColor(i).Value = True Then intColorSeleccionado = i
        End If
    Next i
    
    'Ejecuta el jugador virtual
    Shell App.Path & "\..\JV\TEGNet_JV.exe |" & GintPuerto & "|" & GstrServidor & _
            "|" & txtNickName.Text & "|" & CStr(intColorSeleccionado) & "|" & intAgresividad & "|" & intActitud
    
    Unload Me
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdAceptar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ErrorHandle
    
    Unload Me
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdCancelar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub TabStrip1_Click()
    On Error GoTo ErrorHandle
    
    SeleccionarSolapa TabStrip1.SelectedItem.Index
    
    Exit Sub
ErrorHandle:
    ReportErr "TabStrip1_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub SeleccionarSolapa(intSolapa As Integer)
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    For i = 1 To TabStrip1.Tabs.Count
        If i = intSolapa Then
            fraSolapas(i - 1).Visible = True
        Else
            fraSolapas(i - 1).Visible = False
        End If
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "SeleccionarSolapa", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtNickName_Change()
    On Error GoTo ErrorHandle
    
    'Una vez cargado el nombre y seleccionado el color
    'habilita el botón aceptar
    
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

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    'Regional
    Me.Caption = ObtenerTextoRecurso(CintJVCaption)
    Me.cmdAceptar.Caption = ObtenerTextoRecurso(CintJVbtnAceptar)
    Me.cmdCancelar.Caption = ObtenerTextoRecurso(CintJVbtnCancelar)
    Me.TabStrip1.Tabs(1).Caption = ObtenerTextoRecurso(CintJVGeneral)
    Me.TabStrip1.Tabs(2).Caption = ObtenerTextoRecurso(CintJVPerfil)
    Me.lblNickName.Caption = ObtenerTextoRecurso(CintJVNombre)
    Me.chkAleatorio.Caption = ObtenerTextoRecurso(CintJVAleatorio)
    Me.lblAgresividad.Caption = ObtenerTextoRecurso(CintJVNivelAgresividad)
    Me.lblMuyAgresivo.Caption = ObtenerTextoRecurso(CintJVMuyAgresivo)
    Me.lblPocoAgresivo.Caption = ObtenerTextoRecurso(CintJVPocoAgresivo)
    Me.lblActitud.Caption = ObtenerTextoRecurso(CintJVActitud)
    Me.lblArriesgado.Caption = ObtenerTextoRecurso(CintJVArriesgado)
    Me.lblConservador.Caption = ObtenerTextoRecurso(CintJVConservador)
    
    cmdAceptar.Enabled = False
    
    'Carga las imagenes de los colores
    For i = 1 To optColor.Count
        imgFicha(i).Picture = mdifrmPrincipal.imgLstFichas.ListImages(i).Picture
    Next
    
    'Por defecto, asigna un perfil aleatorio
    chkAleatorio.Value = 1
    HabilitarPerfil False
    
    Actualizar
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Load", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub Actualizar()
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    'Reccorre el vector de jugadores
    For i = 1 To UBound(GvecJugadores)
        If GEstadoCliente = estValidado Then
            'INGRESO
            If GvecJugadores(i).intEstado = conConectado And GvecJugadores(i).strNombre <> "" Then
                'Ingreso (deshabilita los que ya ingresaron)
                optColor(i).Caption = GvecJugadores(i).strNombre
                optColor(i).Enabled = False
            Else
                'Ingreso (habilita los disponibles)
                optColor(i).Caption = ObtenerTextoRecurso(CintJVDisponible) '"Disponible"
                optColor(i).Enabled = True
            End If
            CambiarEstadoPantalla 0
        Else
            'REEMPLAZO o REINGRESO
            If GvecJugadores(i).intEstado = conNoJuega Then
                'Reconexion (deshabilita los que no juegan)
                optColor(i).Caption = ObtenerTextoRecurso(CintJVNoDisponible) '"No disponible"
                optColor(i).Enabled = False
            Else
                'Reconexion (habilita los que estan jugando)
                optColor(i).Caption = GvecJugadores(i).strNombre
                optColor(i).Enabled = True
            End If
            CambiarEstadoPantalla 1
        End If
    Next i
        
    Exit Sub
ErrorHandle:
    ReportErr "Actualizar", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub optColor_Click(Index As Integer)
    On Error GoTo ErrorHandle
    
    'Si se esta reconectando, copia el nombre de la opcion al textbox
    'sino le pasa el foco para que ingrese el nombre
    If GEstadoCliente <> estValidado Then
        txtNickName.Text = IIf(Left(optColor(Index).Caption, 1) = chrPREFIJOADM, Mid(optColor(Index).Caption, 2), optColor(Index).Caption)
    Else
        txtNickName.SetFocus
    End If
    
    'Una vez cargado el nombre y seleccionado el color
    'habilita el botón aceptar
    If Trim(txtNickName.Text) <> "" Then
        cmdAceptar.Enabled = True
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "optColor_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub


