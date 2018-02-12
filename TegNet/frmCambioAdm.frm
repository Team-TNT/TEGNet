VERSION 5.00
Begin VB.Form frmCambioAdm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Administrador"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmCambioAdm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSeleccionColor 
      Caption         =   "Seleccione un nuevo Administrador"
      Height          =   1830
      Left            =   45
      TabIndex        =   2
      Top             =   30
      Width           =   4605
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   6
         Left            =   2925
         TabIndex        =   8
         Top             =   1320
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   660
         TabIndex        =   7
         Top             =   420
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   660
         TabIndex        =   6
         Top             =   870
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   660
         TabIndex        =   5
         Top             =   1320
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   4
         Left            =   2925
         TabIndex        =   4
         Top             =   420
         Width           =   1635
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   270
         Index           =   5
         Left            =   2925
         TabIndex        =   3
         Top             =   870
         Width           =   1635
      End
      Begin VB.Image imgFicha 
         Height          =   435
         Index           =   2
         Left            =   105
         Top             =   765
         Width           =   495
      End
      Begin VB.Image imgFicha 
         Height          =   435
         Index           =   3
         Left            =   105
         Top             =   1215
         Width           =   495
      End
      Begin VB.Image imgFicha 
         Height          =   435
         Index           =   5
         Left            =   2400
         Top             =   765
         Width           =   495
      End
      Begin VB.Image imgFicha 
         Height          =   435
         Index           =   6
         Left            =   2400
         Top             =   1215
         Width           =   495
      End
      Begin VB.Image imgFicha 
         Height          =   435
         Index           =   4
         Left            =   2400
         Top             =   315
         Width           =   495
      End
      Begin VB.Image imgFicha 
         Height          =   435
         Index           =   1
         Left            =   105
         Top             =   315
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   2070
      TabIndex        =   1
      Top             =   1995
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3420
      TabIndex        =   0
      Top             =   1995
      Width           =   1230
   End
End
Attribute VB_Name = "frmCambioAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    On Error GoTo ErrorHandle
    
    Dim intNuevoAdm As Integer
    Dim color As Integer
    
    intNuevoAdm = 0
    
    'Busca la opción seleccionada
    For color = 1 To optColor.Count
        If optColor(color).Value = True Then
            intNuevoAdm = color
        End If
    Next color
    
    If intNuevoAdm = 0 Then
        'Debe seleccionar un jugador.
        MsgBox ObtenerTextoRecurso(CintCambioAdmErrSeleccionJugador), vbInformation, ObtenerTextoRecurso(CintCambioAdmErrSeleccionJugadorCaption)
    Else
        'Envía el mensaje al servidor
        cCambiarAdm intNuevoAdm
    End If
    
    Unload Me
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdAceptar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    Dim color As Integer
    
    'Regional
    Me.Caption = ObtenerTextoRecurso(CintCambioAdmCaption)
    Me.fraSeleccionColor.Caption = ObtenerTextoRecurso(CintCambioAdmFrame)
    Me.cmdAceptar.Caption = ObtenerTextoRecurso(CintCambioAdmBtnAceptar)
    Me.cmdCancelar.Caption = ObtenerTextoRecurso(CintCambioAdmBtnCancelar)
    
    cmdAceptar.Enabled = False
    
    For color = 1 To optColor.Count
        imgFicha(color).Picture = mdifrmPrincipal.imgLstFichas.ListImages(color).Picture
    Next color
    
    For color = 1 To UBound(GvecJugadores)
        If GvecJugadores(color).strNombre = "" Then
            optColor(color).Enabled = False
            optColor(color).Caption = ObtenerTextoRecurso(CintCambioAdmNoDisponible) '"No Disponible"
        Else
            'Deshabilita al administrador actual
            If color = GintMiColor Then
                optColor(color).Enabled = False
            Else
                optColor(color).Enabled = True
            End If
            optColor(color).Caption = GvecJugadores(color).strNombre
        End If
    Next color
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Load", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub imgFicha_Click(Index As Integer)
    On Error GoTo ErrorHandle
    
    If optColor(Index).Enabled Then
        optColor(Index).Value = True
    End If

    Exit Sub
ErrorHandle:
    ReportErr "imgFicha_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub optColor_Click(Index As Integer)
    
    cmdAceptar.Enabled = optColor(Index).Value
    
End Sub
