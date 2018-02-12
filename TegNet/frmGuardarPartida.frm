VERSION 5.00
Begin VB.Form frmGuardarPartida 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guardar Partida"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmGuardarPartida.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   2745
      TabIndex        =   5
      Top             =   2895
      Width           =   1230
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Guardar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   1335
      TabIndex        =   4
      Top             =   2895
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Height          =   2745
      Left            =   210
      TabIndex        =   0
      Top             =   15
      Width           =   4845
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   3345
         TabIndex        =   2
         Top             =   1800
         Width           =   1185
      End
      Begin VB.TextBox txtNombrePartida 
         Height          =   300
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   3
         Top             =   2325
         Width           =   2700
      End
      Begin VB.ListBox lstPartidasGuardadas 
         Height          =   1230
         Left            =   330
         TabIndex        =   1
         Top             =   510
         Width           =   4185
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Partidas Guardadas:"
         Height          =   225
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   3900
      End
      Begin VB.Label lblNombrePartida 
         Caption         =   "Nombre de Partida:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2370
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmGuardarPartida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    On Error GoTo ErrorHandle
    
    'Valida el texto ingresado
    If Not ValidaTexto(txtNombrePartida.Text, 20) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgCaracterInvalido), vbInformation, ObtenerTextoRecurso(CintGralMsgCaracterInvalidoCaption)
        Exit Sub
    End If
    
    If Len(txtNombrePartida.Text) > 0 Then
        cGuardarPartida Trim(txtNombrePartida.Text)
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdAceptar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo ErrorHandle
    
    If lstPartidasGuardadas.ListIndex >= 0 Then
        cEliminarPartida GvecNombresPartidas(lstPartidasGuardadas.ListIndex)
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdEliminar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    Me.Caption = ObtenerTextoRecurso(CintGuardarCaption)
    Me.lblTitulo = ObtenerTextoRecurso(CintGuardarTitulo)
    Me.cmdEliminar.Caption = ObtenerTextoRecurso(CintGuardarEliminar)
    Me.lblNombrePartida.Caption = ObtenerTextoRecurso(CintGuardarNombrePartida)
    Me.cmdAceptar.Caption = ObtenerTextoRecurso(CintGuardarGuardar)
    Me.cmdCancelar.Caption = ObtenerTextoRecurso(CintGuardarCancelar)

    Exit Sub
ErrorHandle:
    ReportErr "Form_Load", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub lstPartidasGuardadas_Click()
    On Error GoTo ErrorHandle
    
    cmdEliminar.Enabled = True
    
    txtNombrePartida.Text = GvecNombresPartidas(lstPartidasGuardadas.ListIndex)
    
    Exit Sub
ErrorHandle:
    ReportErr "lstPartidasGuardadas_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtNombrePartida_Change()
    On Error GoTo ErrorHandle
    
    If Len(txtNombrePartida.Text) > 0 Then
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "txtNombrePartida_Change", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtNombrePartida_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    
    If Not ValidaTexto(Chr$(KeyAscii), 0) Then
        
        KeyAscii = 0
        Beep
    
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "txtNombrePartida_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub
