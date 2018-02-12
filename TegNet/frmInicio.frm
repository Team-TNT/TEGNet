VERSION 5.00
Begin VB.Form frmInicio 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4530
   ClientLeft      =   1335
   ClientTop       =   375
   ClientWidth     =   5685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmInicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   135
      Top             =   75
   End
   Begin VB.Label lblCopyRight 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2001-2004 Argentina"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   1425
      TabIndex        =   3
      Top             =   4155
      Width           =   2745
   End
   Begin VB.Label lblSlogan 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESTRATEGIA SIN LÍMITES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1245
      TabIndex        =   2
      Top             =   3840
      Width           =   3195
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión 1.3.4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   300
      TabIndex        =   1
      Top             =   3525
      Width           =   1215
   End
   Begin VB.Label lblTimer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   5265
      TabIndex        =   0
      Top             =   3210
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   4575
      Left            =   0
      Picture         =   "frmInicio.frx":030B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5700
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Regional
    Me.lblSlogan.Caption = UCase(ObtenerTextoRecurso(CintAppSlogan))
    lblVersion.Caption = ObtenerTextoRecurso(CintInicioVersion) & ": " & App.Major & "." & App.Minor & "." & App.Revision
    lblCopyRight.Caption = ObtenerTextoRecurso(CintInicioCopyright)
    'lblArgentina.Caption = ObtenerTextoRecurso(CintInicioArgentina)
End Sub

Private Sub Timer1_Timer()
    On Error GoTo ErrorHandle
    
    lblTimer.Caption = lblTimer.Caption - 1
    
    If lblTimer.Caption <= 0 Then
        Unload Me
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "Timer1_Timer", Me.Name, Err.Description, Err.Number, Err.Source
End Sub
