VERSION 5.00
Begin VB.Form frmConquista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conquista"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5025
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   330
      Left            =   1897
      TabIndex        =   4
      Top             =   1785
      Width           =   1230
   End
   Begin VB.Frame fraConquista 
      Caption         =   "Cantidad de tropas a mover al País Conquistado:"
      Height          =   1500
      Left            =   105
      TabIndex        =   3
      Top             =   90
      Width           =   4815
      Begin VB.OptionButton optConquista 
         Caption         =   "Tres Tropas"
         Height          =   240
         Index           =   2
         Left            =   480
         TabIndex        =   2
         Top             =   1035
         Width           =   2445
      End
      Begin VB.OptionButton optConquista 
         Caption         =   "Dos Tropas"
         Height          =   240
         Index           =   1
         Left            =   480
         TabIndex        =   1
         Top             =   675
         Width           =   2445
      End
      Begin VB.OptionButton optConquista 
         Caption         =   "Una Tropa"
         Height          =   240
         Index           =   0
         Left            =   480
         TabIndex        =   0
         Top             =   315
         Value           =   -1  'True
         Width           =   2445
      End
      Begin VB.Image Image1 
         Height          =   1155
         Left            =   3075
         Picture         =   "frmConquista.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmConquista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function TropasAMover(intCantMax As Integer, strPaisConquistado As String)
    Dim i As Integer
    
    'Deshabilita las opciones que no se pueden seleccionar
    For i = intCantMax To optConquista.Count - 1
        optConquista(i).Enabled = False
    Next
    
    '###Regional
    Me.Caption = CompilarMensaje(ObtenerTextoRecurso(CintGralMsgTropasMoverCaption), Array(strPaisConquistado))
    optConquista(0).Caption = ObtenerTextoRecurso(CintConquistaUnaTropa)
    optConquista(1).Caption = ObtenerTextoRecurso(CintConquistaDosTropas)
    optConquista(2).Caption = ObtenerTextoRecurso(CintConquistaTresTropas)
    
    fraConquista.Caption = ObtenerTextoRecurso(CintGralMsgTropasMover)
    MostrarFormulario Me, vbModal
    
    'Busca el option seleccionado
    For i = 0 To optConquista.Count - 1
        If optConquista(i).Value Then
            Exit For
        End If
    Next
    
    TropasAMover = i + 1
    
    Unload Me

End Function

Private Sub cmdAceptar_Click()
    Me.Hide
End Sub

