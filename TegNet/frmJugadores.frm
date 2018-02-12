VERSION 5.00
Begin VB.Form frmJugadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jugadores"
   ClientHeight    =   5550
   ClientLeft      =   9195
   ClientTop       =   330
   ClientWidth     =   2445
   Icon            =   "frmJugadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   2445
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle"
      Height          =   2070
      Left            =   60
      TabIndex        =   13
      Top             =   3435
      Width           =   2340
      Begin VB.Label lblCanje 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1785
         TabIndex        =   23
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label lblCantTarjetas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1785
         TabIndex        =   22
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label lblTropasDisponibles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1785
         TabIndex        =   21
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lblCantTropas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1785
         TabIndex        =   20
         Top             =   600
         Width           =   450
      End
      Begin VB.Label lblCantPaises 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1780
         TabIndex        =   19
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblTitCanje 
         Caption         =   "Canjes Realizados:"
         Height          =   270
         Left            =   75
         TabIndex        =   18
         Top             =   1687
         Width           =   1660
      End
      Begin VB.Label lblTitCantTarjetas 
         Caption         =   "Tarjetas:"
         Height          =   270
         Left            =   75
         TabIndex        =   17
         Top             =   1327
         Width           =   1660
      End
      Begin VB.Label lblTitTropasDisponibles 
         Caption         =   "Tropas para agregar:"
         Height          =   270
         Left            =   75
         TabIndex        =   16
         Top             =   967
         Width           =   1660
      End
      Begin VB.Label lblTitCantTropas 
         Caption         =   "Tropas:"
         Height          =   270
         Left            =   75
         TabIndex        =   15
         Top             =   607
         Width           =   1660
      End
      Begin VB.Label lblTitCantPaises 
         Caption         =   "Paises:"
         Height          =   270
         Left            =   75
         TabIndex        =   14
         Top             =   247
         Width           =   1660
      End
   End
   Begin VB.Frame fraRonda 
      Caption         =   "Ronda"
      Height          =   3060
      Left            =   45
      TabIndex        =   0
      Top             =   330
      Width           =   2355
      Begin VB.OptionButton optDetalle 
         Caption         =   "Option1"
         ForeColor       =   &H00000040&
         Height          =   240
         Index           =   5
         Left            =   2085
         TabIndex        =   29
         Top             =   2670
         Width           =   195
      End
      Begin VB.OptionButton optDetalle 
         Caption         =   "Option1"
         ForeColor       =   &H00000040&
         Height          =   240
         Index           =   4
         Left            =   2085
         TabIndex        =   28
         Top             =   2190
         Width           =   195
      End
      Begin VB.OptionButton optDetalle 
         Caption         =   "Option1"
         ForeColor       =   &H00000040&
         Height          =   240
         Index           =   3
         Left            =   2085
         TabIndex        =   27
         Top             =   1710
         Width           =   195
      End
      Begin VB.OptionButton optDetalle 
         Caption         =   "Option1"
         ForeColor       =   &H00000040&
         Height          =   240
         Index           =   2
         Left            =   2085
         TabIndex        =   26
         Top             =   1230
         Width           =   195
      End
      Begin VB.OptionButton optDetalle 
         Caption         =   "Option1"
         ForeColor       =   &H00000040&
         Height          =   240
         Index           =   1
         Left            =   2085
         TabIndex        =   25
         Top             =   750
         Width           =   195
      End
      Begin VB.OptionButton optDetalle 
         Caption         =   "Option1"
         ForeColor       =   &H00000040&
         Height          =   240
         Index           =   0
         Left            =   2085
         TabIndex        =   24
         Top             =   270
         Width           =   195
      End
      Begin VB.Label LblNom 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   225
         Width           =   1575
      End
      Begin VB.Label LblNom 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   705
         Width           =   1575
      End
      Begin VB.Label LblNom 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   10
         Top             =   1185
         Width           =   1575
      End
      Begin VB.Label LblNom 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   9
         Top             =   1665
         Width           =   1575
      End
      Begin VB.Label LblNom 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   480
         TabIndex        =   8
         Top             =   2145
         Width           =   1575
      End
      Begin VB.Label LblNom 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   480
         TabIndex        =   7
         Top             =   2625
         Width           =   1575
      End
      Begin VB.Label LblCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   225
         Width           =   375
      End
      Begin VB.Label LblCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   705
         Width           =   375
      End
      Begin VB.Label LblCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1185
         Width           =   375
      End
      Begin VB.Label LblCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   1665
         Width           =   375
      End
      Begin VB.Label LblCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   2
         Top             =   2145
         Width           =   375
      End
      Begin VB.Label LblCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   1
         Top             =   2625
         Width           =   375
      End
   End
   Begin VB.Label lblLogoR 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "®"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000084&
      Height          =   225
      Left            =   1905
      TabIndex        =   32
      Top             =   45
      Width           =   270
   End
   Begin VB.Label lblLogoNET 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Net"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000084&
      Height          =   405
      Left            =   1305
      TabIndex        =   31
      Top             =   -30
      Width           =   600
   End
   Begin VB.Label lblLogoTEG 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TEG"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Left            =   525
      TabIndex        =   30
      Top             =   -30
      Width           =   765
   End
End
Attribute VB_Name = "frmJugadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error GoTo ErrorHandle
    
    ActualizarTurno (Int(Rnd() * 6) + 1)
    
    Exit Sub
ErrorHandle:
    ReportErr "Command1_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    'Regional
    Me.Caption = ObtenerTextoRecurso(CintJugadoresCaption)
    Me.fraRonda.Caption = ObtenerTextoRecurso(CintJugadoresRonda)
    Me.fraDetalle.Caption = ObtenerTextoRecurso(CintJugadoresDetalle)
    Me.lblTitCantPaises.Caption = ObtenerTextoRecurso(CintJugadoresPaises)
    Me.lblTitCantTropas.Caption = ObtenerTextoRecurso(CintJugadoresTropas)
    Me.lblTitTropasDisponibles.Caption = ObtenerTextoRecurso(CintJugadoresTropasDisponibles)
    Me.lblTitCantTarjetas.Caption = ObtenerTextoRecurso(CintJugadoresTarjetas)
    Me.lblTitCanje.Caption = ObtenerTextoRecurso(CintJugadoresCanjes)
    
    ActualizarRonda
    'Posiciona inicialmente el formulario
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Load", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorHandle
    
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
    Me.Visible = False
    mdifrmPrincipal.mnuVerJugadores.Checked = False

    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub ActualizarRonda()
    On Error GoTo ErrorHandle
    
    'Cambia el orden de la ronda
    Dim Color As Integer
    
    'Oculta todos las posiciones
    For Color = 0 To LblCol.Count - 1
        LblCol(Color).Visible = False
        LblNom(Color).Visible = False
        optDetalle(Color).Visible = False
    Next Color
    
    
    For Color = 1 To UBound(GvecJugadores)
        'Si el jugador soy Yo muestro mis detalles
        If Color = GintMiColor Then
            optDetalle(GvecJugadores(Color).intOrdenRonda - 1).Value = True
        End If
        'Si el jugador juega
        If GvecJugadores(Color).intEstado <> conNoJuega Then
            LblCol(GvecJugadores(Color).intOrdenRonda - 1).Visible = True
            LblNom(GvecJugadores(Color).intOrdenRonda - 1).Visible = True
            optDetalle(GvecJugadores(Color).intOrdenRonda - 1).Visible = True
            LblCol(GvecJugadores(Color).intOrdenRonda - 1).BackColor = GvecColores(Color)
            LblNom(GvecJugadores(Color).intOrdenRonda - 1).Caption = GvecJugadores(Color).strNombre
        End If
    Next Color
    
    'Actualiza el formulario de jugadores
    Actualizar
    
    Exit Sub
ErrorHandle:
    ReportErr "ActualizarRonda", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub ActualizarTurno(ColorActual As Integer)
    On Error GoTo ErrorHandle
    
    'Cambia el turno actual
    Dim Color As Integer
    
    For Color = 1 To UBound(GvecJugadores)
        'Si el jugador no juega
        If GvecJugadores(Color).intOrdenRonda <> 0 Then
            If Color = ColorActual Then
                LblNom(GvecJugadores(Color).intOrdenRonda - 1).BackColor = &H80C0FF   'vbGrayText
                'LblNom(GvecJugadores(color).intOrdenRonda - 1).ForeColor = vbWhite
                LblNom(GvecJugadores(Color).intOrdenRonda - 1).FontBold = True
            Else
                LblNom(GvecJugadores(Color).intOrdenRonda - 1).BackColor = &HC0FFFF   'vbWhite
                'LblNom(GvecJugadores(color).intOrdenRonda - 1).ForeColor = vbBlack
                LblNom(GvecJugadores(Color).intOrdenRonda - 1).FontBold = False
            End If
        End If
    Next

    Exit Sub
ErrorHandle:
    ReportErr "ActualizarTurno", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub ActualizarDetalleJugadores(intColor As Integer)
    On Error GoTo ErrorHandle
    
    'Solo refresca los valores si son del jugador seleccionado
    If optDetalle(MayorACero(GvecJugadores(intColor).intOrdenRonda - 1)).Value = True Then
        With GvecJugadores(intColor)
            fraDetalle.Caption = CompilarMensaje(ObtenerTextoRecurso(CintJugadoresDetalleDe), Array(.strNombre)) '"Detalle de " & .strNombre
            lblCanje = .intCanje
            lblCantPaises = CantidadPaises(intColor)
            lblCantTarjetas = .intCantidadTarjetas
            lblCantTropas = CantidadTropas(intColor)
            lblTropasDisponibles = .intTropasDisponibles
        End With
    End If

    Exit Sub
ErrorHandle:
    ReportErr "ActualizarDetalleJugadores", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub LblCol_Click(Index As Integer)
    On Error GoTo ErrorHandle
    
    optDetalle(Index).Value = True

    Exit Sub
ErrorHandle:
    ReportErr "LblCol_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub LblNom_Click(Index As Integer)
    On Error GoTo ErrorHandle
    
    optDetalle(Index).Value = True

    Exit Sub
ErrorHandle:
    ReportErr "LblNom_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub optDetalle_Click(Index As Integer)
    On Error GoTo ErrorHandle
    
    Dim Color As Integer
    
    For Color = 1 To UBound(GvecJugadores)
        If GvecJugadores(Color).intOrdenRonda = Index + 1 Then
            ActualizarDetalleJugadores Color
            cMostrarTropasDisponibles Color
        End If
    Next

    Exit Sub
ErrorHandle:
    ReportErr "optDetalle_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub ActualizarDetalleJugadorSeleccionado()
    On Error GoTo ErrorHandle
    Dim intOrdenJugadorSeleccionado As Integer
    Dim i As Integer
    
    'Busca el jugador seleccionado
    For i = 0 To optDetalle.Count - 1
        If optDetalle(i).Value = True Then
            intOrdenJugadorSeleccionado = i
        End If
    Next i
    
    'Dispara el evento click del option seleccionado
    optDetalle_Click intOrdenJugadorSeleccionado
    
    Exit Sub
ErrorHandle:
    ReportErr "ActualizarDetalleJugadorSeleccionado", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub Actualizar()
    On Error GoTo ErrorHandle
    Dim i As Integer
    'Si está jugando actualiza el estado de los clientes
    If GEstadoCliente >= estEsperandoTurno Then
        'Recorre el vector de jugadores
        For i = 1 To UBound(GvecJugadores)
            If GvecJugadores(i).intOrdenRonda > 0 Then
                LblNom(GvecJugadores(i).intOrdenRonda - 1).Caption = GvecJugadores(i).strNombre
                LblNom(GvecJugadores(i).intOrdenRonda - 1).ToolTipText = ObtenerTextoRecurso(enuIndiceArchivoRecurso.pmsInteligenciaJugador + GvecJugadores(i).byTipoJugador) & " - " & GvecJugadores(i).strDirIP & " - " & GvecJugadores(i).strVersion
                LblCol(GvecJugadores(i).intOrdenRonda - 1).ToolTipText = LblNom(GvecJugadores(i).intOrdenRonda - 1).ToolTipText
                optDetalle(GvecJugadores(i).intOrdenRonda - 1).ToolTipText = LblNom(GvecJugadores(i).intOrdenRonda - 1).ToolTipText
                If GvecJugadores(i).intEstado = conConectado Then
                    LblNom(GvecJugadores(i).intOrdenRonda - 1).Enabled = True
                ElseIf GvecJugadores(i).intEstado = conDesconectado Then
                    LblNom(GvecJugadores(i).intOrdenRonda - 1).Enabled = False
                End If
            End If
        Next i
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "Actualizar", Me.Name, Err.Description, Err.Number, Err.Source
End Sub
