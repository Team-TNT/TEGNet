VERSION 5.00
Begin VB.Form frmMision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Misión"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Garamond"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMision.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3750
   Begin VB.Label lblMision 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00004080&
      Height          =   1290
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   3300
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   2115
      Left            =   0
      Picture         =   "frmMision.frx":014A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3750
   End
End
Attribute VB_Name = "frmMision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    'Para que funcione bien en XP y en Clasico...
    Me.Height = 2085 + (Me.Height - Me.ScaleHeight)
    Me.Width = 3750 + (Me.Width - Me.ScaleWidth)

    'Regional
    Me.Caption = ObtenerTextoRecurso(CintMisionCaption)
    
    Actualizar
End Sub

Public Function Actualizar()
    lblMision.Caption = UCase(GstrMision)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
    Me.Visible = False
    mdifrmPrincipal.mnuVerMision.Checked = False
End Sub


