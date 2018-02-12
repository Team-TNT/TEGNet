VERSION 5.00
Begin VB.Form frmPropiedades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   Icon            =   "frmPropiedades.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3045
   Visible         =   0   'False
   Begin VB.Label lblTitTropasFijas 
      BackStyle       =   0  'Transparent
      Caption         =   "Tropas Fijas:"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1170
      TabIndex        =   6
      Top             =   960
      Width           =   960
   End
   Begin VB.Label lblTropasFijas 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2205
      TabIndex        =   5
      Top             =   975
      Width           =   300
   End
   Begin VB.Label lblNombrePais 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   270
      Left            =   1215
      TabIndex        =   4
      Top             =   45
      Width           =   2475
   End
   Begin VB.Label lblEjercitos 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1890
      TabIndex        =   3
      Top             =   675
      Width           =   300
   End
   Begin VB.Label lblTitEjercitos 
      BackStyle       =   0  'Transparent
      Caption         =   "Tropas:"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1170
      TabIndex        =   2
      Top             =   667
      Width           =   675
   End
   Begin VB.Label lblOwner 
      BackStyle       =   0  'Transparent
      Caption         =   "dueño"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1890
      TabIndex        =   1
      Top             =   375
      Width           =   1035
   End
   Begin VB.Label lblTitOwner 
      BackStyle       =   0  'Transparent
      Caption         =   "Dueño:"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1170
      TabIndex        =   0
      Top             =   375
      Width           =   675
   End
   Begin VB.Image imgPais 
      Height          =   1020
      Left            =   120
      Stretch         =   -1  'True
      Top             =   135
      Width           =   960
   End
End
Attribute VB_Name = "frmPropiedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    'Regional
    Me.Caption = ObtenerTextoRecurso(CintInformacionCaption)
    Me.lblTitOwner.Caption = ObtenerTextoRecurso(CintInformacionDuenio)
    Me.lblTitEjercitos.Caption = ObtenerTextoRecurso(CintInformacionTropas)
    Me.lblTitTropasFijas.Caption = ObtenerTextoRecurso(CintInformacionFijas)
    
    'Limpia los labels
    lblOwner.Caption = ""
    lblEjercitos.Caption = ""
    lblTropasFijas.Caption = ""

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
    mdifrmPrincipal.mnuVerInfo.Checked = False

    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub lblNombrePais_Change()
    On Error GoTo ErrorHandle
    
    'Cambia el caption de la ventana
    Me.Caption = CompilarMensaje(ObtenerTextoRecurso(CintInformacionCaptionDinamica), Array(lblNombrePais.Caption)) '"Información - " & lblNombrePais.Caption

    Exit Sub
ErrorHandle:
    ReportErr "lblNombrePais_Change", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

