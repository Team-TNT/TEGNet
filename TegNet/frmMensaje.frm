VERSION 5.00
Begin VB.Form frmMensaje 
   BackColor       =   &H00C5FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   2100
   ClientLeft      =   3810
   ClientTop       =   3225
   ClientWidth     =   3780
   Icon            =   "frmMensaje.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3780
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5FFFF&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1485
      Width           =   930
   End
   Begin VB.Label lblMensaje 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1230
      Left            =   90
      TabIndex        =   1
      Top             =   105
      Width           =   3585
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   2115
      Left            =   0
      Picture         =   "frmMensaje.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3750
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle

    cmdAceptar.Caption = ObtenerTextoRecurso(CintMensajeAceptar)

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
    
    '###
    frmMapa.Enabled = True
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub


