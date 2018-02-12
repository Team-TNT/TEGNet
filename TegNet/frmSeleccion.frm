VERSION 5.00
Begin VB.Form frmSeleccion 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Selección"
   ClientHeight    =   495
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2490
   Icon            =   "frmSeleccion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblHasta 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   765
      TabIndex        =   3
      Top             =   270
      Width           =   1620
   End
   Begin VB.Label lblDesde 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   765
      TabIndex        =   2
      Top             =   15
      Width           =   1620
   End
   Begin VB.Label lblTitHasta 
      Caption         =   "Hasta:"
      Height          =   270
      Left            =   90
      TabIndex        =   1
      Top             =   270
      Width           =   585
   End
   Begin VB.Label lblTitDesde 
      Caption         =   "Desde:"
      Height          =   270
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   585
   End
End
Attribute VB_Name = "frmSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    'Regional
    Me.Caption = " " & ObtenerTextoRecurso(CintSeleccionCaption)
    Me.lblTitDesde.Caption = ObtenerTextoRecurso(CintSeleccionDesde)
    Me.lblTitHasta.Caption = ObtenerTextoRecurso(CintSeleccionHasta)
    
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
    mdifrmPrincipal.mnuVerSeleccion.Checked = False

    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Resize()
    With Me
        'Centrado del contenido
        'Reparte el espacio libre (2/5, 1/5 y 2/5)
        .lblDesde.Top = (.ScaleHeight - (.lblDesde.Height + .lblHasta.Height)) * (2 / 5)
        .lblHasta.Top = .lblDesde.Top + .lblDesde.Height + ((.ScaleHeight - (.lblDesde.Height + .lblHasta.Height)) * (1 / 5))
        .lblTitDesde.Top = .lblDesde.Top
        .lblTitHasta.Top = .lblHasta.Top
    End With
End Sub
