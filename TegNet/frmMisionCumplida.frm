VERSION 5.00
Begin VB.Form frmMisionCumplida 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Misión Cumplida"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMisionCumplida.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1620
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2580
      Width           =   1395
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   195
      Picture         =   "frmMisionCumplida.frx":014A
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1020
   End
   Begin VB.Label lblMision 
      BackStyle       =   0  'Transparent
      Caption         =   "Conquistar el mundo"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1290
      Left            =   330
      TabIndex        =   1
      Top             =   1320
      Width           =   3930
   End
   Begin VB.Label lblGanador 
      BackStyle       =   0  'Transparent
      Caption         =   "JUANCITO HA LOGRADO CUMPLIR SU MISIÓN:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   990
      Left            =   1410
      TabIndex        =   0
      Top             =   195
      Width           =   3150
   End
End
Attribute VB_Name = "frmMisionCumplida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Regional
    Me.Caption = ObtenerTextoRecurso(CintMisionCumplidaCaption)
    Me.cmdAceptar.Caption = ObtenerTextoRecurso(CintMisionCumplidaAceptar)
End Sub

