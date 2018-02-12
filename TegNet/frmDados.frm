VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1620
   Icon            =   "frmDados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   1620
   Begin MSComctlLib.ImageList iLstDados 
      Left            =   1485
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDados.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDados.frx":0812
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDados.frx":0ED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDados.frx":15AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDados.frx":1C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDados.frx":239E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgJugador2 
      Height          =   540
      Index           =   2
      Left            =   915
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   540
   End
   Begin VB.Image imgJugador2 
      Height          =   540
      Index           =   1
      Left            =   915
      Stretch         =   -1  'True
      Top             =   780
      Width           =   540
   End
   Begin VB.Image imgJugador2 
      Height          =   540
      Index           =   0
      Left            =   915
      Stretch         =   -1  'True
      Top             =   210
      Width           =   540
   End
   Begin VB.Image imgJugador1 
      Height          =   540
      Index           =   2
      Left            =   135
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   540
   End
   Begin VB.Image imgJugador1 
      Height          =   540
      Index           =   1
      Left            =   135
      Stretch         =   -1  'True
      Top             =   780
      Width           =   540
   End
   Begin VB.Image imgJugador1 
      Height          =   540
      Index           =   0
      Left            =   135
      Stretch         =   -1  'True
      Top             =   210
      Width           =   540
   End
   Begin VB.Label lblAtacante 
      Alignment       =   2  'Center
      Caption         =   "Ataque"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   765
   End
   Begin VB.Label lblDefensor 
      Caption         =   "Defensa"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   180
      Left            =   870
      TabIndex        =   1
      Top             =   0
      Width           =   795
   End
End
Attribute VB_Name = "frmDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    'Regional
    Me.Caption = ObtenerTextoRecurso(CintDadosCaption)
    Me.lblAtacante.Caption = ObtenerTextoRecurso(CintDadosAtaque)
    Me.lblDefensor.Caption = ObtenerTextoRecurso(CintDadosDefensa)

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
    mdifrmPrincipal.mnuVerDados.Checked = False

    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

