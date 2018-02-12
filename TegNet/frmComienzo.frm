VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmComienzo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asistente de Conexión"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   Icon            =   "frmComienzo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEtapas 
      Height          =   3120
      Index           =   3
      Left            =   4650
      TabIndex        =   41
      Top             =   3120
      Width           =   4545
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   330
         Index           =   3
         Left            =   3510
         TabIndex        =   29
         Top             =   2700
         Width           =   945
      End
      Begin VB.CommandButton cmdAtras 
         Caption         =   "< &Atrás"
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   150
         TabIndex        =   26
         Top             =   2700
         Width           =   945
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "&Siguiente >"
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   1155
         TabIndex        =   27
         Top             =   2700
         Width           =   945
      End
      Begin VB.CommandButton cmdFinalizar 
         Caption         =   "C&onectar"
         Height          =   330
         Index           =   3
         Left            =   2520
         TabIndex        =   28
         Top             =   2700
         Width           =   945
      End
      Begin VB.Frame fraInterior 
         Height          =   2520
         Index           =   3
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   4545
         Begin VB.OptionButton optVieja 
            Caption         =   "Partida Guardada"
            Height          =   225
            Left            =   360
            TabIndex        =   24
            Top             =   870
            Width           =   2910
         End
         Begin VB.OptionButton optUltima 
            Caption         =   "Última Partida"
            Height          =   225
            Left            =   330
            TabIndex        =   23
            Top             =   375
            Value           =   -1  'True
            Width           =   2910
         End
         Begin VB.ListBox lstPartidasGuardadas 
            Enabled         =   0   'False
            Height          =   1035
            Left            =   585
            TabIndex        =   25
            Top             =   1185
            Width           =   3525
         End
         Begin VB.Label lblPartidasGuardadas 
            Caption         =   "Partidas Guardadas:"
            Height          =   405
            Left            =   360
            TabIndex        =   43
            Top             =   390
            Width           =   3600
         End
      End
   End
   Begin VB.Frame fraEtapas 
      Height          =   3120
      Index           =   2
      Left            =   4665
      TabIndex        =   37
      Top             =   0
      Width           =   4545
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   330
         Index           =   2
         Left            =   3510
         TabIndex        =   12
         Top             =   2700
         Width           =   945
      End
      Begin VB.CommandButton cmdAtras 
         Caption         =   "< &Atrás"
         Height          =   330
         Index           =   2
         Left            =   150
         TabIndex        =   9
         Top             =   2700
         Width           =   945
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "&Siguiente >"
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   1155
         TabIndex        =   10
         Top             =   2700
         Width           =   945
      End
      Begin VB.CommandButton cmdFinalizar 
         Caption         =   "C&onectar"
         Height          =   330
         Index           =   2
         Left            =   2520
         TabIndex        =   11
         Top             =   2700
         Width           =   945
      End
      Begin VB.Frame fraInterior 
         Height          =   2520
         Index           =   2
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   4545
         Begin VB.TextBox txtServidorRemoto 
            Height          =   315
            Left            =   1740
            TabIndex        =   6
            ToolTipText     =   "Nombre o Direccion IP del Servidor"
            Top             =   1065
            Width           =   1365
         End
         Begin VB.TextBox txtPuertoRemoto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1755
            MaxLength       =   4
            TabIndex        =   7
            Text            =   "5479"
            Top             =   1770
            Width           =   525
         End
         Begin MSComCtl2.UpDown UpdPuertoRemoto 
            Height          =   315
            Left            =   2265
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1770
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1001
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtPuertoRemoto"
            BuddyDispid     =   196620
            OrigLeft        =   1785
            OrigTop         =   1485
            OrigRight       =   1980
            OrigBottom      =   1800
            Max             =   9000
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Conectándose con el Servidor..."
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   120
            TabIndex        =   50
            Top             =   2160
            Width           =   4305
         End
         Begin VB.Image imgServidor 
            Height          =   480
            Left            =   270
            Picture         =   "frmComienzo.frx":014A
            Stretch         =   -1  'True
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label6 
            Caption         =   "Ingrese los datos del Servidor al cual desea conectarse."
            Height          =   510
            Left            =   1065
            TabIndex        =   45
            Top             =   345
            Width           =   3120
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Servidor:"
            Height          =   255
            Left            =   1050
            TabIndex        =   40
            Top             =   1095
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Puerto:"
            Height          =   210
            Left            =   1080
            TabIndex        =   39
            Top             =   1800
            Width           =   585
         End
      End
   End
   Begin VB.Frame fraEtapas 
      Height          =   3120
      Index           =   1
      Left            =   45
      TabIndex        =   32
      Top             =   3120
      Width           =   4545
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   330
         Index           =   1
         Left            =   3525
         TabIndex        =   22
         Top             =   2700
         Width           =   945
      End
      Begin VB.CommandButton cmdAtras 
         Caption         =   "< &Atrás"
         Height          =   330
         Index           =   1
         Left            =   150
         TabIndex        =   19
         Top             =   2700
         Width           =   945
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "&Siguiente >"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1155
         TabIndex        =   20
         Top             =   2700
         Width           =   945
      End
      Begin VB.CommandButton cmdFinalizar 
         Caption         =   "C&onectar"
         Height          =   330
         Index           =   1
         Left            =   2520
         TabIndex        =   21
         Top             =   2700
         Width           =   945
      End
      Begin VB.Frame fraInterior 
         Height          =   2520
         Index           =   1
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   4545
         Begin VB.CommandButton cmdAvanzado 
            Caption         =   "&Avanzado..."
            Height          =   300
            Left            =   3405
            TabIndex        =   15
            Top             =   2115
            Width           =   1035
         End
         Begin VB.Frame fraAvanzado 
            Height          =   1380
            Left            =   60
            TabIndex        =   34
            Top             =   1080
            Visible         =   0   'False
            Width           =   4425
            Begin VB.OptionButton optRemoto 
               Caption         =   "Remoto"
               Height          =   255
               Left            =   3120
               TabIndex        =   48
               Top             =   255
               Width           =   930
            End
            Begin VB.OptionButton optLocal 
               Caption         =   "Local"
               Height          =   255
               Left            =   2070
               TabIndex        =   47
               Top             =   255
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.CheckBox chkActivarServidor 
               Caption         =   "Activar Servidor"
               Height          =   240
               Left            =   2340
               TabIndex        =   46
               Top             =   690
               Value           =   1  'Checked
               Width           =   1860
            End
            Begin VB.TextBox txtServidor 
               Enabled         =   0   'False
               Height          =   315
               Left            =   765
               TabIndex        =   18
               Text            =   "localhost"
               ToolTipText     =   "Nombre o Dirección IP del Servidor"
               Top             =   660
               Width           =   1365
            End
            Begin VB.TextBox txtPuerto 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   765
               MaxLength       =   4
               TabIndex        =   16
               Text            =   "5479"
               Top             =   1005
               Width           =   525
            End
            Begin MSComCtl2.UpDown UpDPuerto 
               Height          =   315
               Left            =   1290
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   1005
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   556
               _Version        =   393216
               Value           =   1001
               BuddyControl    =   "txtPuerto"
               BuddyDispid     =   196632
               OrigLeft        =   1395
               OrigTop         =   255
               OrigRight       =   1590
               OrigBottom      =   570
               Max             =   9000
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label1 
               Caption         =   "Ubicación del Servidor:"
               Height          =   255
               Left            =   105
               TabIndex        =   49
               Top             =   285
               Width           =   1785
            End
            Begin VB.Label lblServidor 
               Alignment       =   1  'Right Justify
               Caption         =   "Servidor:"
               Enabled         =   0   'False
               Height          =   255
               Left            =   100
               TabIndex        =   36
               Top             =   690
               Width           =   630
            End
            Begin VB.Label lblPuerto 
               Alignment       =   1  'Right Justify
               Caption         =   "Puerto:"
               Height          =   255
               Left            =   100
               TabIndex        =   35
               Top             =   1035
               Width           =   630
            End
         End
         Begin VB.OptionButton optGuardada 
            Caption         =   "Cargar Partida Guardada"
            Height          =   300
            Left            =   1275
            TabIndex        =   14
            Top             =   1470
            Width           =   3015
         End
         Begin VB.OptionButton optNueva 
            Caption         =   "Iniciar Nueva Partida"
            Height          =   300
            Left            =   1275
            TabIndex        =   13
            Top             =   750
            Value           =   -1  'True
            Width           =   3015
         End
         Begin VB.Image imgGuardada 
            Height          =   480
            Left            =   675
            Picture         =   "frmComienzo.frx":0454
            Stretch         =   -1  'True
            Top             =   1365
            Width           =   480
         End
         Begin VB.Image imgNueva 
            Height          =   480
            Left            =   690
            Picture         =   "frmComienzo.frx":075E
            Stretch         =   -1  'True
            Top             =   660
            Width           =   480
         End
      End
   End
   Begin VB.Frame fraEtapas 
      Height          =   3120
      Index           =   0
      Left            =   60
      TabIndex        =   30
      Top             =   0
      Width           =   4545
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   330
         Index           =   0
         Left            =   3525
         TabIndex        =   5
         Top             =   2700
         Width           =   945
      End
      Begin VB.Frame fraInterior 
         Height          =   2520
         Index           =   0
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   4545
         Begin VB.OptionButton optOrganizar 
            Caption         =   "Organizar una Partida"
            Height          =   300
            Left            =   1260
            TabIndex        =   0
            Top             =   1035
            Value           =   -1  'True
            Width           =   3135
         End
         Begin VB.OptionButton optUnirse 
            Caption         =   "Unirse a Partida en red"
            Height          =   300
            Left            =   1260
            TabIndex        =   1
            Top             =   1740
            Width           =   3135
         End
         Begin VB.Image imgUnirse 
            Height          =   480
            Left            =   645
            Picture         =   "frmComienzo.frx":0A68
            Stretch         =   -1  'True
            Top             =   1635
            Width           =   480
         End
         Begin VB.Image imgOrganizar 
            Height          =   480
            Left            =   660
            Picture         =   "frmComienzo.frx":1332
            Stretch         =   -1  'True
            Top             =   930
            Width           =   480
         End
         Begin VB.Label lblComentario 
            Caption         =   "Bienvenido al Asistente de Conexión de TEGNet que lo guiará en el inicio de una Partida."
            Height          =   525
            Left            =   225
            TabIndex        =   44
            Top             =   330
            Width           =   4035
         End
      End
      Begin VB.CommandButton cmdFinalizar 
         Caption         =   "Ac&eptar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   2520
         TabIndex        =   4
         Top             =   2700
         Width           =   945
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "&Siguiente >"
         Height          =   330
         Index           =   0
         Left            =   1155
         TabIndex        =   3
         Top             =   2700
         Width           =   945
      End
      Begin VB.CommandButton cmdAtras 
         Caption         =   "< &Atrás"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   2700
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmComienzo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum enuEtapas
    etaInicio = 0
    etaOrganizar = 1
    etaUnirse = 2
    etaGuardada = 3
End Enum

Dim blnSalioPorAceptar As Boolean
Dim blnConfirmarSalir As Boolean

'Regional - Variables que contienen el caption del formulario
'para los distintos idiomas
Dim strCaption1 As String
Dim strCaption2 As String
Dim strCaption3 As String
Dim strCaption4 As String
Dim strAvanzado As String
Dim strOcultar As String

Private Sub cmdAtras_Click(Index As Integer)
    On Error GoTo ErrorHandle
    
    Select Case Index
        Case enuEtapas.etaOrganizar
            'Retrocede a la etapa de inicio
            MostrarEtapa etaInicio
        Case enuEtapas.etaUnirse
            'Retrocede a la etapa de inicio
            MostrarEtapa etaInicio
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdAtras_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdAvanzado_Click()
    On Error GoTo ErrorHandle
    
    If cmdAvanzado.Caption = strAvanzado Then
        cmdAvanzado.Caption = strOcultar
        fraAvanzado.Visible = True
        
        optNueva.Top = 305
        imgNueva.Top = 305 - 100
        optGuardada.Top = 755
        imgGuardada.Top = 755 - 90
    Else
        cmdAvanzado.Caption = strAvanzado
        fraAvanzado.Visible = False
        
        optNueva.Top = 750
        imgNueva.Top = 750 - 90
        optGuardada.Top = 1470
        imgGuardada.Top = 1470 - 90
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdAzanzado_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    On Error GoTo ErrorHandle
    
    Unload Me
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdCancelar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdFinalizar_Click(Index As Integer)
    On Error GoTo ErrorHandle
    Dim strNombre As String
    
    blnSalioPorAceptar = True
    Screen.MousePointer = vbHourglass
            
    Select Case Index
        Case enuEtapas.etaOrganizar
            
            'Valida el número ingresado
            If Not ValidaEntero(txtPuerto.Text, UpDPuerto.Min, UpDPuerto.Max) Then
                MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            'Valida que haya ingresado un nombre de servidor
            If Trim(txtServidor.Text) = "" Then
                MsgBox "Debe ingresar el nombre o la dirección IP del Servidor donde desea organizar la partida.", vbExclamation, "Atención"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            'Nueva Partida (levanta el servidor y se conecta como administrador)
            cConectarAdm txtServidor.Text, CInt(Me.txtPuerto.Text), IIf(chkActivarServidor.Value = Checked, True, False)
            TipoPartida = tpNueva
            
            Screen.MousePointer = vbDefault
            Unload Me
            
        Case enuEtapas.etaUnirse
            'Valida el número ingresado
            If Not ValidaEntero(txtPuertoRemoto.Text, UpdPuertoRemoto.Min, UpdPuertoRemoto.Max) Then
                MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
                Exit Sub
            End If
            
            Label4.Caption = ObtenerTextoRecurso(CintComienzoConectando)
            
            'Unirse a una partida existente
            cConectar Me.txtServidorRemoto.Text, CInt(txtPuertoRemoto.Text)
            'Por si se conecta como administrador
            TipoPartida = tpNueva
            
        Case enuEtapas.etaGuardada
        
            If optUltima.Value = True Then
                strNombre = strNombrePartidaActual
            Else
                strNombre = GvecNombresPartidas(lstPartidasGuardadas.ListIndex)
            End If
            EnviarMensaje ArmarMensajeParam(msgNombrePartida, strNombre)
            
            Screen.MousePointer = vbDefault
            Unload Me
    
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdFinalizar_Click", Me.Name, Err.Description, Err.Number, Err.Source
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSiguiente_Click(Index As Integer)
    On Error GoTo ErrorHandle
    
    Select Case Index
        Case enuEtapas.etaInicio
            If optOrganizar.Value = True Then
                'Eligio Organizar: Pasa a la etapa 1
                MostrarEtapa etaOrganizar
            Else
                'Eligio Unirse: Pasa a la etapa 2
                MostrarEtapa etaUnirse
            End If
        Case enuEtapas.etaOrganizar
            If optGuardada.Value = True Then
                Screen.MousePointer = vbHourglass
                
                'Eligio guardada: Se conecta al servidor y espera la
                'lista de partidas guardadas
                'Valida el puerto ingresado
                If Not ValidaEntero(txtPuerto.Text, UpDPuerto.Min, UpDPuerto.Max) Then
                    MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                'Valida que haya ingresado un nombre de servidor
                If Trim(txtServidor.Text) = "" Then
                    MsgBox "Debe ingresar el nombre o la dirección IP del Servidor donde desea organizar la partida.", vbExclamation, "Atención"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                'Nueva Partida (levanta el servidor y se conecta como administrador)
                cConectarAdm txtServidor.Text, CInt(Me.txtPuerto.Text), IIf(chkActivarServidor.Value = Checked, True, False)
                TipoPartida = enuTipoPartida.tpGuardada
                
                'Al cancelar solo pide confirmacion en la pantalla
                'de partidas guardadas
                blnConfirmarSalir = True
                
                'Deshabilita el botón Siguiente para que no pueda
                'levantar otro Servidor mientras realiza la conexión
                cmdSiguiente(enuEtapas.etaOrganizar).Enabled = False

            End If
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdSiguiente_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    blnSalioPorAceptar = False
    blnConfirmarSalir = False
    
    cmdFinalizar(2).Enabled = False
    
    'Acomoda los frames
    fraEtapas(enuEtapas.etaInicio).BorderStyle = 0
    For i = 0 To fraEtapas.Count - 1
        fraEtapas(i).Top = fraEtapas(enuEtapas.etaInicio).Top
        fraEtapas(i).Left = fraEtapas(enuEtapas.etaInicio).Left
        fraEtapas(i).Width = fraEtapas(enuEtapas.etaInicio).Width
        fraEtapas(i).Height = fraEtapas(enuEtapas.etaInicio).Height
        fraEtapas(i).BorderStyle = fraEtapas(enuEtapas.etaInicio).BorderStyle
    
        'Regional
        cmdAtras(i).Caption = ObtenerTextoRecurso(CintComienzoAtras)
        cmdSiguiente(i).Caption = ObtenerTextoRecurso(CintComienzoSiguiente)
        'En los frames 0, 1 y 3, el botón "Conectar" tiene el nombre "Aceptar"
        If i = 0 Or i = 1 Or i = 3 Then
            cmdFinalizar(i).Caption = ObtenerTextoRecurso(CintComienzoAceptar)
        Else
            cmdFinalizar(i).Caption = ObtenerTextoRecurso(CintComienzoConectar)
        End If
        cmdCancelar(i).Caption = ObtenerTextoRecurso(CintComienzoCancelar)
    Next i
    
    'Regional
    lblComentario.Caption = ObtenerTextoRecurso(CintComienzo1Principal)
    optOrganizar.Caption = ObtenerTextoRecurso(CintComienzo1Organizar)
    optUnirse.Caption = ObtenerTextoRecurso(CintComienzo1Unirse)
    Label6.Caption = ObtenerTextoRecurso(CintComienzo2Principal)
    Label3.Caption = ObtenerTextoRecurso(CintComienzo2Servidor)
    Label2.Caption = ObtenerTextoRecurso(CintComienzo2Puerto)
    Label4.Caption = ""
    optNueva.Caption = ObtenerTextoRecurso(CintComienzo3Nueva)
    optGuardada.Caption = ObtenerTextoRecurso(CintComienzo3Guardada)
    cmdAvanzado.Caption = ObtenerTextoRecurso(CintComienzo3Avanzado)
    strAvanzado = ObtenerTextoRecurso(CintComienzo3Avanzado)
    strOcultar = ObtenerTextoRecurso(CintComienzo3Ocultar)
    Label1.Caption = ObtenerTextoRecurso(CintComienzo3Ubicacion)
    optLocal.Caption = ObtenerTextoRecurso(CintComienzo3Local)
    optRemoto.Caption = ObtenerTextoRecurso(CintComienzo3Remoto)
    chkActivarServidor.Caption = ObtenerTextoRecurso(CintComienzo3Activar)
    lblServidor.Caption = ObtenerTextoRecurso(CintComienzo3Servidor)
    lblPuerto.Caption = ObtenerTextoRecurso(CintComienzo3Puerto)
    optUltima.Caption = ObtenerTextoRecurso(CintComienzo4Ultima)
    optVieja.Caption = ObtenerTextoRecurso(CintComienzo4Guardada)
    Me.Caption = ObtenerTextoRecurso(CintComienzoCaption)
    strCaption1 = ObtenerTextoRecurso(CintComienzo1Caption)
    strCaption2 = ObtenerTextoRecurso(CintComienzo2Caption)
    strCaption3 = ObtenerTextoRecurso(CintComienzo3Caption)
    strCaption4 = ObtenerTextoRecurso(CintComienzo4Caption)
    
    'Para que funcione bien en XP y en Clasico...
    Me.Width = 4665 + (Me.Width - Me.ScaleWidth)
    Me.Height = 3165 + (Me.Height - Me.ScaleHeight)
    
    'Pone al primero arriba
    MostrarEtapa etaInicio
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Load", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    
    If blnSalioPorAceptar Then
        Exit Sub
    End If
    
    If Not blnConfirmarSalir Then
        Exit Sub
    End If
    
    'Si todavía no seleccionó el color se desconecta
    If MsgBox(ObtenerTextoRecurso(CintComienzoMsgDesconectar), vbQuestion + vbYesNo + vbDefaultButton2, ObtenerTextoRecurso(CintComienzoMsgDesconectarCaption)) = vbNo Then
        Cancel = 1
    Else
        If GEstadoCliente >= estConectado Then
            'Cuando se desconecta el último jugador se baja el servidor
            cDesconectar
        End If
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Unload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub optGuardada_Click()
    On Error GoTo ErrorHandle
    
    cmdFinalizar(enuEtapas.etaOrganizar).Enabled = optNueva.Value
    cmdSiguiente(enuEtapas.etaOrganizar).Enabled = optGuardada.Value
    cmdSiguiente(enuEtapas.etaOrganizar).Default = True
    
    Exit Sub
ErrorHandle:
    ReportErr "optGuardada_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub optLocal_Click()
    On Error GoTo ErrorHandle
    
    lblServidor.Enabled = False
    txtServidor.Text = "localhost"
    txtServidor.Enabled = False
    
    chkActivarServidor.Enabled = True
    chkActivarServidor.Value = Checked
    
    Exit Sub
ErrorHandle:
    ReportErr "optLocal_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub optRemoto_Click()
    On Error GoTo ErrorHandle
    
    lblServidor.Enabled = True
    txtServidor.Enabled = True
    txtServidor.Text = ""
    
    chkActivarServidor.Value = Unchecked
    chkActivarServidor.Enabled = False
    
    Exit Sub
ErrorHandle:
    ReportErr "optRemoto_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub optNueva_Click()
    On Error GoTo ErrorHandle
    
    cmdFinalizar(enuEtapas.etaOrganizar).Enabled = optNueva.Value
    cmdSiguiente(enuEtapas.etaOrganizar).Enabled = optGuardada.Value
    cmdFinalizar(enuEtapas.etaOrganizar).Default = True
    
    Exit Sub
ErrorHandle:
    ReportErr "optNueva_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtPuerto_GotFocus()
    On Error GoTo ErrorHandle
    
    'Pinta el textbox
    txtPuerto.SelStart = 0
    txtPuerto.SelLength = Len(txtPuerto.Text)
    
    Exit Sub
ErrorHandle:
    ReportErr "txtPuerto_GotFocus", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtPuerto_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    
    If Not ValidaEntero(Chr$(KeyAscii)) Then
      KeyAscii = 0   'Cancela el caracter.
      Beep
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "txtPuerto_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtServidor_GotFocus()
    On Error GoTo ErrorHandle
    
    'Pinta el textbox
    txtServidor.SelStart = 0
    txtServidor.SelLength = Len(txtServidor.Text)
    
    Exit Sub
ErrorHandle:
    ReportErr "txtServidor_GotFocus", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtServidorRemoto_Change()
    On Error GoTo ErrorHandle
    
    If Len(txtServidorRemoto.Text) > 0 Then
        cmdFinalizar(2).Enabled = True
    Else
        cmdFinalizar(2).Enabled = False
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "txtServidorRemoto_Change", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtServidorRemoto_GotFocus()
    On Error GoTo ErrorHandle
    
    'Pinta el textbox
    txtServidorRemoto.SelStart = 0
    txtServidorRemoto.SelLength = Len(txtServidorRemoto.Text)
    
    Exit Sub
ErrorHandle:
    ReportErr "txtServidorRemoto_GotFocus", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtPuertoRemoto_GotFocus()
    On Error GoTo ErrorHandle
    
    'Pinta el textbox
    txtPuertoRemoto.SelStart = 0
    txtPuertoRemoto.SelLength = Len(txtPuertoRemoto.Text)
    
    Exit Sub
ErrorHandle:
    ReportErr "txtPuertoRemoto_GotFocus", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtPuertoRemoto_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    
    If Not ValidaEntero(Chr$(KeyAscii)) Then
      KeyAscii = 0   'Cancela el caracter.
      Beep
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "txtPuertoRemoto_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub lstPartidasGuardadas_Click()
    On Error GoTo ErrorHandle
    
    If lstPartidasGuardadas.SelCount = 0 Then
        cmdFinalizar(enuEtapas.etaGuardada).Enabled = False
    Else
        cmdFinalizar(enuEtapas.etaGuardada).Enabled = True
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "lstPartidasGuardadas_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub optVieja_Click()
    On Error GoTo ErrorHandle
    
    lstPartidasGuardadas.Enabled = True
    If lstPartidasGuardadas.SelCount = 0 Then
        cmdFinalizar(enuEtapas.etaGuardada).Enabled = False
    Else
        cmdFinalizar(enuEtapas.etaGuardada).Enabled = True
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "optVieja_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub optUltima_Click()
    On Error GoTo ErrorHandle
    
    lstPartidasGuardadas.Enabled = False
    cmdFinalizar(enuEtapas.etaGuardada).Enabled = True
    
    Exit Sub
ErrorHandle:
    ReportErr "optUltima_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub MostrarEtapa(Indice As enuEtapas)
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    'Oculta todas las etapas
    For i = 0 To fraEtapas.Count - 1
        fraEtapas(i).Visible = False
    Next
    
    fraEtapas(Indice).Visible = True
    cmdCancelar(Indice).Cancel = True
    
    'Setea botón por default según la etapa
    Select Case Indice
        Case enuEtapas.etaInicio
            Me.Caption = strCaption1
            cmdSiguiente(Indice).Default = True
        Case enuEtapas.etaOrganizar
            Me.Caption = strCaption3
            cmdFinalizar(Indice).Default = True
        Case enuEtapas.etaGuardada
            Me.Caption = strCaption4
            cmdFinalizar(Indice).Default = True
        Case enuEtapas.etaUnirse
            Me.Caption = strCaption2
            cmdFinalizar(Indice).Default = True
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "MostrarEtapa", Me.Name, Err.Description, Err.Number, Err.Source
End Sub
