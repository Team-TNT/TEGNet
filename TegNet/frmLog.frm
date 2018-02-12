VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Bitácora"
   ClientHeight    =   915
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   8955
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstLog 
      BackColor       =   &H00C0E0FF&
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8760
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    'Regional
    Me.Caption = " " & ObtenerTextoRecurso(CintLogCaption)
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Load", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lstLog.Width = Me.ScaleWidth 'ScaleWidth es el ancho sin los bordes de la ventana
    lstLog.Height = Me.ScaleHeight 'ScaleHeight es el alto sin los bordes de la ventana
    Me.Height = (Me.Height - Me.ScaleHeight) + lstLog.Height
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorHandle
    
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
    Me.Visible = False
    mdifrmPrincipal.mnuVerLog.Checked = False

    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub


