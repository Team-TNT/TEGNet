VERSION 5.00
Begin VB.Form frmIdioma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Idioma / Language"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4650
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1020
      TabIndex        =   2
      Top             =   1995
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2430
      TabIndex        =   3
      Top             =   1995
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione su Idioma / Select your Language"
      Height          =   1800
      Left            =   75
      TabIndex        =   4
      Top             =   75
      Width           =   4500
      Begin VB.OptionButton optEnglish 
         Caption         =   "&English"
         Height          =   315
         Left            =   1710
         TabIndex        =   1
         Top             =   1140
         Width           =   1860
      End
      Begin VB.OptionButton optSpanish 
         Caption         =   "E&spañol"
         Height          =   315
         Left            =   1695
         TabIndex        =   0
         Top             =   465
         Value           =   -1  'True
         Width           =   1860
      End
      Begin VB.Image imgEnglish 
         Height          =   480
         Left            =   900
         Picture         =   "frmIdioma.frx":0000
         Stretch         =   -1  'True
         Top             =   1065
         Width           =   600
      End
      Begin VB.Image imgSpanish 
         Height          =   480
         Left            =   885
         Picture         =   "frmIdioma.frx":043F
         Stretch         =   -1  'True
         Top             =   375
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmIdioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnIdiomaSeteado As Boolean

Public Property Let IdiomaSeteado(blnValor As Boolean)
    'Propiedad que indica si el idioma ya fué seteado alguna vez
    'o si es la primera vez que se elige
    On Error GoTo ErrorHandle
    
    blnIdiomaSeteado = blnValor
    
    cmdCancelar.Enabled = blnValor
    
    Exit Property
ErrorHandle:
    ReportErr "IdiomaSeteado.Let", Me.Name, Err.Description, Err.Number, Err.Source
End Property

Private Sub cmdAceptar_Click()
    On Error GoTo ErrorHandle
    
    Dim intBaseIdioma As Integer
    
    'Guarda en la registry el idioma seleccionado
    If optSpanish.Value = True Then
        intBaseIdioma = CintBaseSpanish
    Else
        intBaseIdioma = CintBaseEnglish
    End If
    
    GrabarSeteo "BaseIdioma", CStr(intBaseIdioma)
    
    'Si no está seteado el idioma
    If Not blnIdiomaSeteado Then
        'aplica el idioma seleccionado
        GintBaseIdioma = intBaseIdioma
    Else
        'solo si se modificó el idioma
        If GintBaseIdioma <> intBaseIdioma Then
            'avisa que es necesario reiniciar.
            MsgBox ObtenerTextoRecurso(CintIdiomaMsgCambio), vbInformation, ObtenerTextoRecurso(CintIdiomaMsgCambioCaption)
        End If
    End If
    
    Unload Me
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdCancelar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ErrorHandle
    
    Unload Me

    Exit Sub
ErrorHandle:
    ReportErr "cmdCancelar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    'Selecciona el idioma actual
    Select Case GintBaseIdioma
        Case CintBaseSpanish
            optSpanish.Value = True
        Case CintBaseEnglish
            optEnglish.Value = True
    End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdCancelar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub imgSpanish_Click()
    On Error GoTo ErrorHandle
    
    optSpanish.Value = True
    
    Exit Sub
ErrorHandle:
    ReportErr "imgSpanish_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub imgEnglish_Click()
    On Error GoTo ErrorHandle
    
    optEnglish.Value = True
    
    Exit Sub
ErrorHandle:
    ReportErr "imgEnglish_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

