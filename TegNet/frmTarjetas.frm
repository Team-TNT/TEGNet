VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTarjetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mis Tarjetas"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7890
   Icon            =   "frmTarjetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   7890
   Begin VB.CommandButton cmdCobrar 
      Caption         =   "Cobrar"
      Height          =   315
      Left            =   2910
      TabIndex        =   12
      Top             =   2520
      Width           =   990
   End
   Begin VB.CommandButton cmdCanjear 
      Caption         =   "Canjear"
      Height          =   315
      Left            =   3990
      TabIndex        =   1
      Top             =   2520
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Height          =   2490
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   7830
      Begin VB.Shape shpTarjeta 
         BorderWidth     =   2
         Height          =   2160
         Index           =   0
         Left            =   300
         Top             =   585
         Width           =   1455
      End
      Begin VB.Shape shpTarjeta 
         BorderWidth     =   2
         Height          =   2160
         Index           =   1
         Left            =   1890
         Top             =   600
         Width           =   1410
      End
      Begin VB.Shape shpTarjeta 
         BorderWidth     =   2
         Height          =   2160
         Index           =   2
         Left            =   3465
         Top             =   570
         Width           =   1455
      End
      Begin VB.Shape shpTarjeta 
         BorderWidth     =   2
         Height          =   2160
         Index           =   3
         Left            =   5115
         Top             =   555
         Width           =   1410
      End
      Begin VB.Shape shpTarjeta 
         BorderWidth     =   2
         Height          =   2160
         Index           =   4
         Left            =   6585
         Top             =   510
         Width           =   1455
      End
      Begin VB.Shape shpTarjetaSel 
         BorderColor     =   &H000040C0&
         BorderWidth     =   6
         Height          =   2160
         Index           =   0
         Left            =   525
         Top             =   765
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Shape shpTarjetaSel 
         BorderColor     =   &H000040C0&
         BorderWidth     =   6
         Height          =   2160
         Index           =   1
         Left            =   2085
         Top             =   765
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Shape shpTarjetaSel 
         BorderColor     =   &H000040C0&
         BorderWidth     =   6
         Height          =   2160
         Index           =   2
         Left            =   3645
         Top             =   765
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Shape shpTarjetaSel 
         BorderColor     =   &H000040C0&
         BorderWidth     =   6
         Height          =   2160
         Index           =   3
         Left            =   5190
         Top             =   750
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Shape shpTarjetaSel 
         BorderColor     =   &H000040C0&
         BorderWidth     =   6
         Height          =   2160
         Index           =   4
         Left            =   6690
         Top             =   795
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2130
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Index           =   1
         Left            =   1665
         TabIndex        =   10
         Top             =   2130
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Index           =   2
         Left            =   3210
         TabIndex        =   9
         Top             =   2130
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Index           =   3
         Left            =   4755
         TabIndex        =   8
         Top             =   2130
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Index           =   4
         Left            =   6300
         TabIndex        =   7
         Top             =   2130
         Width           =   1410
      End
      Begin VB.Label lblCobrada 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "COBRADA"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label lblCobrada 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "COBRADA"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   1
         Left            =   1695
         TabIndex        =   5
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label lblCobrada 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "COBRADA"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   4
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label lblCobrada 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "COBRADA"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   3
         Left            =   4785
         TabIndex        =   3
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label lblCobrada 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "COBRADA"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   4
         Left            =   6330
         TabIndex        =   2
         Top             =   195
         Width           =   1335
      End
      Begin VB.Image imgTarjeta 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2205
         Index           =   0
         Left            =   90
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1455
      End
      Begin VB.Image imgTarjeta 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2205
         Index           =   1
         Left            =   1635
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1455
      End
      Begin VB.Image imgTarjeta 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2205
         Index           =   2
         Left            =   3180
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1455
      End
      Begin VB.Image imgTarjeta 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2205
         Index           =   3
         Left            =   4725
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1455
      End
      Begin VB.Image imgTarjeta 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2205
         Index           =   4
         Left            =   6270
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1455
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   95
      ImageHeight     =   145
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTarjetas.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTarjetas.frx":0F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTarjetas.frx":1DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTarjetas.frx":343A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTarjetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intTarjetaActiva As Integer

Public Sub Actualizar()
    On Error GoTo ErrorHandle
    Dim i As Integer
    For i = 1 To 5
        'Si el pais de la tarjeta es 0 es porque no hay tarjeta
        If GvecTarjetas(i).byPais <> 0 Then
            imgTarjeta(i - 1).Picture = ImageList1.ListImages(GvecTarjetas(i).byFigura).Picture
            lblTarjeta(i - 1).Caption = UCase(frmMapa.objPais(GvecTarjetas(i).byPais).Nombre)
            If GvecTarjetas(i).blCobrada Then
                lblCobrada(i - 1).Visible = True
            Else
                lblCobrada(i - 1).Visible = False
            End If
        Else
            imgTarjeta(i - 1).Picture = Nothing
            lblTarjeta(i - 1).Caption = ""
            lblCobrada(i - 1).Visible = False
        End If
    Next

    Exit Sub
ErrorHandle:
    ReportErr "Actualizar", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdCanjear_Click()
    On Error GoTo ErrorHandle
    
    cCanjearTarjeta
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdCanjear_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdCobrar_Click()
    On Error GoTo ErrorHandle
    
    cCobrarTarjeta
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdCobrar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    Dim intLeft As Integer
    Dim intTarjeta As Integer
    
    'Regional
    Me.Caption = ObtenerTextoRecurso(CintTarjetasCaption)
    Me.cmdCobrar.Caption = ObtenerTextoRecurso(CintTarjetasCobrar)
    Me.cmdCanjear.Caption = ObtenerTextoRecurso(CintTarjetasCanjear)
    For intTarjeta = Me.lblCobrada.LBound To Me.lblCobrada.UBound
        Me.lblCobrada.Item(intTarjeta) = ObtenerTextoRecurso(CintTarjetasCobrada)
    Next
    
    intTarjetaActiva = -1
    
    intLeft = 90
    
    'Inicializa valores
    For i = 0 To 4
        shpTarjeta(i).Visible = False
        shpTarjetaSel(i).Visible = False
        lblCobrada(i).Visible = False
        
        shpTarjetaSel(i).BorderWidth = 6
        shpTarjetaSel(i).BorderColor = &H40C0&
        
        imgTarjeta(i).Top = 180
        shpTarjeta(i).Top = 180
        shpTarjetaSel(i).Top = 180
        lblCobrada(i).Top = 195
        
        shpTarjeta(i).Height = imgTarjeta(i).Height
        shpTarjetaSel(i).Height = imgTarjeta(i).Height
        
        shpTarjeta(i).Width = imgTarjeta(i).Width
        shpTarjetaSel(i).Width = imgTarjeta(i).Width
        lblTarjeta(i).Width = imgTarjeta(i).Width
        lblCobrada(i).Width = imgTarjeta(i).Width - 60
        
        shpTarjeta(i).Left = intLeft
        shpTarjetaSel(i).Left = intLeft
        imgTarjeta(i).Left = intLeft
        lblTarjeta(i).Left = intLeft
        lblCobrada(i).Left = intLeft + 30
        
        intLeft = intLeft + imgTarjeta(i).Width + 90
    Next
    
    Actualizar
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Load", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub imgTarjeta_Click(Index As Integer)
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim intCantidadSel As Integer
    Dim intCantMaxSel As Integer
    
    intCantidadSel = 0
    
    'De acuerdo al estado...
    Select Case GEstadoCliente
        Case enuEstadoCli.estTarjetaTomada, enuEstadoCli.estTarjetaCobradaTomada, _
             enuEstadoCli.estAtacando, enuEstadoCli.estMoviendo, _
             enuEstadoCli.estTarjetaCobrada
            
            'Puede seleccionar solo una (para cobrarla)
            intCantMaxSel = 1
        Case enuEstadoCli.estAgregando
            'Puede seleccionar 3 tarjetas (para canjearlas)
            intCantMaxSel = 3
        Case Else
            Exit Sub
    End Select
    
    'Cuenta la cantidad de tarjetas seleccionadas
    For i = 0 To imgTarjeta.Count - 1
        If i <> Index Then
            If shpTarjetaSel(i).Visible Then
                intCantidadSel = intCantidadSel + 1
            End If
        End If
    Next
    
    'Verifica que no haya mas de la cantidad permitida
    If intCantidadSel < intCantMaxSel Then
        If GvecTarjetas(Index + 1).byPais <> 0 Then
            shpTarjetaSel(Index).Visible = Not shpTarjetaSel(Index).Visible
        End If
    Else
        If GvecTarjetas(Index + 1).byPais <> 0 Then
            If shpTarjetaSel(Index).Visible = True Then
                'Si ya está seleccionada solo la desmarca
                shpTarjetaSel(Index).Visible = False
            Else
                'Deselecciona todas las seleccionadas
                For i = 0 To imgTarjeta.Count - 1
                    shpTarjetaSel(i).Visible = False
                Next i
                'Selecciona la nueva
                shpTarjetaSel(Index).Visible = True
            End If
        End If
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "imgTarjeta_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub imgTarjeta_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    If frmMapa.Visible Then
        frmMapa.PaisActivo = GvecTarjetas(CByte(Index + 1)).byPais
        
        If Index <> intTarjetaActiva Then
            intTarjetaActiva = Index
            For i = 0 To 4
                shpTarjeta(i).Visible = False
            Next
            If GvecTarjetas(Index + 1).byPais <> 0 Then
                shpTarjeta(Index).Visible = True
            End If
        End If
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "imgTarjeta_MouseMove", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorHandle
    
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
    Me.Visible = False
    mdifrmPrincipal.mnuVerTarjetas.Checked = False
    
    frmMapa.LimpiarMapa

    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub lblCobrada_Click(Index As Integer)
    On Error GoTo ErrorHandle
    
    imgTarjeta_Click Index

    Exit Sub
ErrorHandle:
    ReportErr "lblCobrada_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub lblCobrada_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandle
    
    imgTarjeta_MouseMove Index, Button, Shift, X, Y

    Exit Sub
ErrorHandle:
    ReportErr "lblCobrada_MouseMove", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub lblTarjeta_Click(Index As Integer)
    On Error GoTo ErrorHandle
    
    imgTarjeta_Click Index

    Exit Sub
ErrorHandle:
    ReportErr "lblTarjeta_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub lblTarjeta_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandle
    
    imgTarjeta_MouseMove Index, Button, Shift, X, Y

    Exit Sub
ErrorHandle:
    ReportErr "lblTarjeta_MouseMove", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub DeseleccionarTarjetas()
    'Deselecciona las tarjetas seleccionadas
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    For i = 0 To shpTarjetaSel.Count - 1
        shpTarjetaSel(i).Visible = False
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "DeseleccionarTarjetas", Me.Name, Err.Description, Err.Number, Err.Source
End Sub
