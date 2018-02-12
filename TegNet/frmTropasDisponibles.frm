VERSION 5.00
Begin VB.Form frmTropasDisponibles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   1665
   ClipControls    =   0   'False
   Icon            =   "frmTropasDisponibles.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   1665
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle"
      Height          =   2410
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   1650
      Begin VB.Label lblTDI 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   6
         Left            =   1185
         TabIndex        =   7
         Top             =   2025
         Width           =   375
      End
      Begin VB.Label lblTDI 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   1185
         TabIndex        =   6
         Top             =   1725
         Width           =   375
      End
      Begin VB.Label lblTDI 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   1185
         TabIndex        =   5
         Top             =   1425
         Width           =   375
      End
      Begin VB.Label lblTDI 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   1185
         TabIndex        =   4
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label lblTDI 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   1185
         TabIndex        =   3
         Top             =   825
         Width           =   375
      End
      Begin VB.Label lblTDI 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   1185
         TabIndex        =   2
         Top             =   525
         Width           =   375
      End
      Begin VB.Label lblTDI 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   1185
         TabIndex        =   1
         Top             =   225
         Width           =   375
      End
      Begin VB.Label lblTitTDI 
         Caption         =   "Oceanía:"
         Height          =   255
         Index           =   6
         Left            =   135
         TabIndex        =   14
         Top             =   2055
         Width           =   1050
      End
      Begin VB.Label lblTitTDI 
         Caption         =   "Europa:"
         Height          =   255
         Index           =   5
         Left            =   135
         TabIndex        =   13
         Top             =   1755
         Width           =   1050
      End
      Begin VB.Label lblTitTDI 
         Caption         =   "Asia:"
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   12
         Top             =   1455
         Width           =   1050
      End
      Begin VB.Label lblTitTDI 
         Caption         =   "A. del Sur:"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   11
         Top             =   1155
         Width           =   1050
      End
      Begin VB.Label lblTitTDI 
         Caption         =   "A. del Norte:"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   10
         Top             =   855
         Width           =   1050
      End
      Begin VB.Label lblTitTDI 
         Caption         =   "África:"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   555
         Width           =   1050
      End
      Begin VB.Label lblTitTDI 
         Caption         =   "Libres:"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   255
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmTropasDisponibles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    'Para que funcione bien en XP y en Clasico...
    Me.Height = 2415 + (Me.Height - Me.ScaleHeight)
    Me.Width = 1660 + (Me.Width - Me.ScaleWidth)
    
    'Regional
    Me.Caption = ObtenerTextoRecurso(CintTropasCaption)
    Me.fraDetalle.Caption = ObtenerTextoRecurso(CintTropasDetalle)
    Me.lblTitTDI(0).Caption = ObtenerTextoRecurso(CintTropasLibres)
    Me.lblTitTDI(1).Caption = ObtenerTextoRecurso(CintTropasAfrica)
    Me.lblTitTDI(2).Caption = ObtenerTextoRecurso(CintTropasANorte)
    Me.lblTitTDI(3).Caption = ObtenerTextoRecurso(CintTropasASur)
    Me.lblTitTDI(4).Caption = ObtenerTextoRecurso(CintTropasAsia)
    Me.lblTitTDI(5).Caption = ObtenerTextoRecurso(CintTropasEuropa)
    Me.lblTitTDI(6).Caption = ObtenerTextoRecurso(CintTropasOceania)
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorHandle
    
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
    Me.Visible = False
    mdifrmPrincipal.mnuVerTropasDisponibles.Checked = False

    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

