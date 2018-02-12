VERSION 5.00
Begin VB.UserControl ctlPais 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaskColor       =   &H00000000&
   Picture         =   "ctlPais.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Windowless      =   -1  'True
   Begin VB.Label lblPais 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1890
      TabIndex        =   0
      Top             =   630
      UseMnemonic     =   0   'False
      Width           =   300
   End
   Begin VB.Shape shpPais 
      BorderColor     =   &H00000080&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1650
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape shpSombra 
      BorderColor     =   &H00404040&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1650
      Shape           =   3  'Circle
      Top             =   1005
      Width           =   255
   End
End
Attribute VB_Name = "ctlPais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MpctNormal As Picture
'Private MpctMouseOver As Picture
Private MpctSeleccionado As Picture

Private MblnEstaSeleccionado As Boolean
Private MblnPuedeSerSeleccionado As Boolean

Private MstrNombre As String
Private MintIndiceColor As Integer
Private MintCantTropas As Integer
Private MintTropasFijas As Integer      'Almacena la cantidad de tropas que ya fueron movidas en el turno

Private MvecColores(5) As Long 'Almacena los colores
Private MvecColoresInv(5) As Long 'Almacena los colores inversos
Private MvecColoresBorde(5) As Long 'Almacena los colores de los bordes

Private MlngShapeTop As Long
Private MlngShapeLeft As Long

Private McolLimitrofes As New Collection

Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DblClick()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Rutinas Privadas
'----------------
Private Sub lblPais_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub lblPais_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblPais_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    'Carga la matriz de colores
    MvecColores(0) = vbBlack
    MvecColores(1) = vbMagenta  '&HC000C0
    MvecColores(2) = &HC0&      'vbRed
    MvecColores(3) = &HC00000   'vbBlue
    MvecColores(4) = &HC0C0&    'vbYellow
    MvecColores(5) = &HC000&    'vbGreen

    MvecColoresInv(0) = vbWhite
    MvecColoresInv(1) = vbWhite
    MvecColoresInv(2) = vbWhite
    MvecColoresInv(3) = vbWhite
    MvecColoresInv(4) = vbBlack
    MvecColoresInv(5) = vbBlack
    
    MvecColoresBorde(0) = vbBlack
    MvecColoresBorde(1) = &H400040
    MvecColoresBorde(2) = &H80&
    MvecColoresBorde(3) = &H400000
    MvecColoresBorde(4) = &H4040&
    MvecColoresBorde(5) = &H4000&
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Set MpctNormal = PropBag.ReadProperty("ImagenNormal")
'    Set MpctMouseOver = PropBag.ReadProperty("ImagenMouseOver")
    Set MpctSeleccionado = PropBag.ReadProperty("ImagenSeleccionado")
    MintCantTropas = PropBag.ReadProperty("CantTropas", 0)
    MintIndiceColor = 2 'PropBag.ReadProperty("Color", 0)
    MintTropasFijas = PropBag.ReadProperty("TropasFijas", 0)
    MstrNombre = PropBag.ReadProperty("Nombre", "")
    MlngShapeTop = PropBag.ReadProperty("ShapeTop")
    MlngShapeLeft = PropBag.ReadProperty("ShapeLeft")
        
    Inicializar
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ImagenNormal", MpctNormal
'    PropBag.WriteProperty "ImagenMouseOver", MpctMouseOver
    PropBag.WriteProperty "ImagenSeleccionado", MpctSeleccionado
    PropBag.WriteProperty "CantTropas", MintCantTropas
    PropBag.WriteProperty "Color", MintIndiceColor
    PropBag.WriteProperty "TropasFijas", MintTropasFijas
    PropBag.WriteProperty "Nombre", MstrNombre
    PropBag.WriteProperty "ShapeTop", MlngShapeTop
    PropBag.WriteProperty "ShapeLeft", MlngShapeLeft
End Sub

Private Sub Inicializar()
    UserControl.Picture = MpctNormal
    UserControl.MaskPicture = MpctNormal
    
    lblPais.Width = shpPais.Width
    lblPais.Height = shpPais.Height
    shpPais.Move MlngShapeLeft, MlngShapeTop
    
    lblPais.Move MlngShapeLeft + 20, MlngShapeTop + 30
    shpSombra.Move MlngShapeLeft + 20, MlngShapeTop + 20
    
    lblPais.Caption = CStr(MintCantTropas)
    
    lblPais.ForeColor = MvecColoresInv(MintIndiceColor)
    shpPais.FillColor = MvecColores(MintIndiceColor)
    shpPais.BorderColor = MvecColoresBorde(MintIndiceColor)
End Sub

Private Sub UserControl_Terminate()
    Set MpctNormal = Nothing
'    Set MpctMouseOver = Nothing
    Set MpctSeleccionado = Nothing
End Sub


'Propiedades
'--------------
Public Property Get ImagenNormal() As StdPicture
    Set ImagenNormal = MpctNormal
End Property

Public Property Set ImagenNormal(ByVal PpctNueva As StdPicture)
    Set MpctNormal = PpctNueva
    
    UserControl.Picture = MpctNormal
    UserControl.MaskPicture = MpctNormal
    
    UserControl.Width = ScaleX(MpctNormal.Width)
    UserControl.Height = ScaleY(MpctNormal.Height)
    
    shpPais.Left = ScaleX(MpctNormal.Width) / 2 - shpPais.Width / 2
    shpPais.Top = ScaleY(MpctNormal.Height) / 2 - shpPais.Height / 2
    lblPais.Left = shpPais.Left
    lblPais.Top = shpPais.Top
    
    MlngShapeTop = shpPais.Top
    MlngShapeLeft = shpPais.Left

End Property

'Public Property Get ImagenMouseOver() As StdPicture
'    Set ImagenMouseOver = MpctMouseOver
'End Property

'Public Property Set ImagenMouseOver(ByVal PpctNueva As StdPicture)
'    Set MpctMouseOver = PpctNueva
'End Property

Public Property Get ImagenSeleccionado() As StdPicture
    Set ImagenSeleccionado = MpctSeleccionado
End Property

Public Property Set ImagenSeleccionado(ByVal PpctNueva As StdPicture)
    Set MpctSeleccionado = PpctNueva
End Property

Public Property Get Ancho() As Single
    Ancho = UserControl.Picture.Width
End Property

Public Property Get Alto() As Single
    Alto = UserControl.Picture.Height
End Property

Public Property Let Nombre(ByVal PstrNombre As String)
    MstrNombre = PstrNombre
    lblPais.ToolTipText = PstrNombre
End Property

Public Property Get Nombre() As String
    Nombre = MstrNombre
End Property

Public Property Let Color(ByVal PintColor As Integer)
    MintIndiceColor = PintColor - 1
    lblPais.ForeColor = MvecColoresInv(MintIndiceColor)
    shpPais.FillColor = MvecColores(MintIndiceColor)
    shpPais.BorderColor = MvecColoresBorde(MintIndiceColor)
End Property

Public Property Get Color() As Integer
    Color = MintIndiceColor + 1
End Property

Public Property Let CantTropas(ByVal PintCantTropas As Integer)
    MintCantTropas = PintCantTropas
    lblPais.Caption = CStr(MintCantTropas)
End Property

Public Property Get CantTropas() As Integer
    CantTropas = MintCantTropas
End Property

Public Property Let TropasFijas(ByVal PintTropasFijas As Integer)
    MintTropasFijas = PintTropasFijas
End Property

Public Property Get TropasFijas() As Integer
    TropasFijas = MintTropasFijas
End Property

Public Property Get PosicionFichaX() As Long
    PosicionFichaX = MlngShapeLeft
End Property

Public Property Let PosicionFichaX(ByVal PlngValor As Long)
    MlngShapeLeft = PlngValor
    shpPais.Left = MlngShapeLeft
    lblPais.Left = MlngShapeLeft + 20
    shpSombra.Left = MlngShapeLeft + 20
End Property

Public Property Get PosicionFichaY() As Long
    PosicionFichaY = MlngShapeTop
End Property

Public Property Let PosicionFichaY(ByVal PlngValor As Long)
    MlngShapeTop = PlngValor
    shpPais.Top = MlngShapeTop
    lblPais.Top = MlngShapeTop + 30
    shpSombra.Top = MlngShapeTop + 20

End Property

Public Property Let MostrarFicha(PblnMostrar As Boolean)
    shpPais.Visible = PblnMostrar
    lblPais.Visible = PblnMostrar
    shpSombra.Visible = PblnMostrar
End Property

Public Property Get MostrarFicha() As Boolean
    MostrarFicha = shpPais.Visible
End Property

Public Property Get EstaSeleccionado() As Boolean
    EstaSeleccionado = MblnEstaSeleccionado
End Property

Public Property Let EstaSeleccionado(PblnValor As Boolean)
    MblnEstaSeleccionado = PblnValor
    If MblnEstaSeleccionado Then
        UserControl.Picture = MpctSeleccionado
    Else
        UserControl.Picture = MpctNormal
    End If
End Property

Public Property Get PuedeSerSeleccionado() As Boolean
    PuedeSerSeleccionado = MblnPuedeSerSeleccionado
End Property

Public Property Let PuedeSerSeleccionado(PblnValor As Boolean)
    MblnPuedeSerSeleccionado = PblnValor
End Property

Public Property Let MousePointer(PintValor As MousePointerConstants)
    UserControl.MousePointer = PintValor
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Set MouseIcon(PpctIcono As Picture)
    UserControl.MouseIcon = PpctIcono
    MousePointer = vbCustom
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

'Métodos
'-------
Public Sub MouseOver()
    If Not MblnEstaSeleccionado Then
        UserControl.Picture = MpctSeleccionado 'MpctMouseOver
    End If
End Sub

Public Sub MouseOut()
    If Not MblnEstaSeleccionado Then
        UserControl.Picture = MpctNormal
    End If
End Sub

Public Sub IniciarDestello()
    UserControl.Picture = MpctSeleccionado 'MpctMouseOver
    'UserControl.Refresh
End Sub

Public Sub FinalizarDestello()
    UserControl.Picture = MpctNormal
    'UserControl.Refresh
End Sub

Public Sub Restaurar()
    If MblnEstaSeleccionado Then
        UserControl.Picture = MpctSeleccionado
    Else
        UserControl.Picture = MpctNormal
    End If
End Sub

Public Sub LimpiarLimitrofes()
    'Elimina todos los elementos de la coleccion
    While McolLimitrofes.Count > 0
        McolLimitrofes.Remove McolLimitrofes.Count
    Wend
End Sub

Public Sub AgregarLimitrofe(PbyPaisLimitrofe As Byte)
    'Agrega un pais a la coleccion de limitrofes
    McolLimitrofes.Add PbyPaisLimitrofe
End Sub

Public Function EsLimitrofe(PbyPais As Byte) As Boolean
    Dim LbyPaisActual As Variant
    Dim LblnReturn As Boolean
    
    LblnReturn = False
    For Each LbyPaisActual In McolLimitrofes
        If CByte(LbyPaisActual) = PbyPais Then
            LblnReturn = True
        End If
    Next
    
    EsLimitrofe = LblnReturn
    
End Function

Public Sub Refresh()
    UserControl.Refresh
End Sub
