VERSION 5.00
Begin VB.Form frmMisiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Misiones"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "frmMisiones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   5940
   Begin VB.Frame Frame1 
      Height          =   6630
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   5775
      Begin VB.Label lblMisiones 
         BackStyle       =   0  'Transparent
         Height          =   6255
         Left            =   180
         TabIndex        =   1
         Top             =   255
         Width           =   5400
      End
   End
End
Attribute VB_Name = "frmMisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    Dim strAuxiliar As String
    
    '###Regional
    Me.Caption = ObtenerTextoRecurso(CintMisionesCaption)
    
    strAuxiliar = ""
    For i = 2 To 16
        strAuxiliar = strAuxiliar & (i - 1) & ". " & ObtenerTextoRecurso(enuIndiceArchivoRecurso.pmsMisiones + i) & IIf(i = 16, "", vbNewLine & vbNewLine)
    Next
    
    'txtMisiones.Text = strAuxiliar
    lblMisiones.Caption = strAuxiliar
    
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
    mdifrmPrincipal.mnuVerListaMisiones.Checked = False

    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub


