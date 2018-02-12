VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCreditos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de TegNet"
   ClientHeight    =   11145
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3960
   Icon            =   "Creditos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11145
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   30
      Top             =   2445
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Creditos.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Creditos.frx":129A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3480
      Top             =   3360
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   105
      TabIndex        =   10
      Top             =   10470
      Width           =   3735
      Begin VB.CommandButton cmdStopPlay 
         Height          =   255
         Left            =   1920
         Picture         =   "Creditos.frx":16F2
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   315
         Left            =   2640
         TabIndex        =   12
         Top             =   210
         Width           =   975
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   360
         Max             =   50
         TabIndex        =   11
         Top             =   240
         Value           =   20
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   10335
      Left            =   135
      ScaleHeight     =   10275
      ScaleWidth      =   3645
      TabIndex        =   0
      Top             =   120
      Width           =   3705
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   18915
         Left            =   210
         ScaleHeight     =   18915
         ScaleWidth      =   3255
         TabIndex        =   1
         Top             =   -11100
         Width           =   3255
         Begin VB.Frame Frame4 
            Height          =   30
            Left            =   120
            TabIndex        =   18
            Top             =   6240
            Width           =   3000
         End
         Begin VB.Frame Frame2 
            Height          =   30
            Left            =   120
            TabIndex        =   7
            Top             =   16305
            Width           =   3000
         End
         Begin VB.Frame Frame1 
            Height          =   30
            Left            =   120
            TabIndex        =   3
            Top             =   4320
            Width           =   3000
         End
         Begin VB.Label lblArgentina 
            BackStyle       =   0  'Transparent
            Caption         =   "Hecho en Argentina"
            Height          =   225
            Left            =   1005
            TabIndex        =   39
            Top             =   18540
            Width           =   1560
         End
         Begin VB.Image imgArgentina 
            Height          =   180
            Left            =   645
            Picture         =   "Creditos.frx":1B34
            Top             =   18555
            Width           =   270
         End
         Begin VB.Label lblAgradecimiento12 
            BackStyle       =   0  'Transparent
            Caption         =   $"Creditos.frx":1BA8
            Height          =   855
            Left            =   165
            TabIndex        =   38
            Top             =   14610
            Width           =   3000
         End
         Begin VB.Label lblAgradecimientoFinal 
            BackStyle       =   0  'Transparent
            Caption         =   "A la Universidad Nacional de La Matanza por formarnos como profesionales."
            Height          =   705
            Left            =   150
            TabIndex        =   37
            Top             =   15600
            Width           =   3000
         End
         Begin VB.Label lblAgradecimiento9 
            BackStyle       =   0  'Transparent
            Caption         =   $"Creditos.frx":1C2F
            Height          =   1005
            Left            =   150
            TabIndex        =   36
            Top             =   12675
            Width           =   3000
         End
         Begin VB.Label lblAgradecimiento11 
            BackStyle       =   0  'Transparent
            Caption         =   "A Andrea Doorman por su colaboración en la difusión del juego."
            Height          =   570
            Left            =   150
            TabIndex        =   35
            Top             =   14055
            Width           =   3000
         End
         Begin VB.Label lblAgradecimiento10 
            BackStyle       =   0  'Transparent
            Caption         =   "A Sabri (novia de Guille) por su paciencia."
            Height          =   525
            Left            =   150
            TabIndex        =   34
            Top             =   13680
            Width           =   3000
         End
         Begin VB.Label lblAgradecimiento7 
            BackStyle       =   0  'Transparent
            Caption         =   "A Gabriel Mansilla (nuestro contacto en Uruguay), por su ayuda en la difusión del juego a nivel internacional."
            Height          =   795
            Left            =   150
            TabIndex        =   32
            Top             =   11265
            Width           =   3000
         End
         Begin VB.Label lblAgradecimiento8 
            BackStyle       =   0  'Transparent
            Caption         =   "A Andrea (novia de Cráneo), por averiguarnos cómo y dónde registrar nuestro producto."
            Height          =   720
            Left            =   150
            TabIndex        =   31
            Top             =   11985
            Width           =   3000
         End
         Begin VB.Label lblAgradecimiento6 
            BackStyle       =   0  'Transparent
            Caption         =   "A Rafael Matesanz por la contribución a nuestro laboratorio (y por los asados)."
            Height          =   615
            Left            =   150
            TabIndex        =   30
            Top             =   10650
            Width           =   3000
         End
         Begin VB.Label lblAgradecimiento5 
            BackStyle       =   0  'Transparent
            Caption         =   "A nuestros Beta-Testers (Gladys, Jose Maria, Marcos y Pablo) por sus sacrificadas horas de pruebas."
            Height          =   825
            Left            =   150
            TabIndex        =   29
            Top             =   9825
            Width           =   3000
         End
         Begin VB.Label lblAgradecimiento4 
            BackStyle       =   0  'Transparent
            Caption         =   "A Fernando Losinno por ayudarnos a configurar la red de nuestro laboratorio."
            Height          =   735
            Left            =   150
            TabIndex        =   28
            Top             =   9090
            Width           =   3000
         End
         Begin VB.Label lblAgradecimiento3 
            BackStyle       =   0  'Transparent
            Caption         =   "A Héctor Rebagliatti (papá de Javi), por ayudarnos en la elaboración de los algoritmos de Inteligencia Artificial."
            Height          =   795
            Left            =   150
            TabIndex        =   27
            Top             =   8295
            Width           =   3000
         End
         Begin VB.Label lblAgradecimiento2 
            BackStyle       =   0  'Transparent
            Caption         =   "A Betty (mamá de Javi) por su aguante."
            Height          =   450
            Left            =   150
            TabIndex        =   26
            Top             =   7845
            Width           =   3000
         End
         Begin VB.Label lblAgradecimiento1 
            BackStyle       =   0  'Transparent
            Caption         =   "A nuestros profesores Carlos Tomassino y Roberto Eribe, por confiar en este proyecto."
            Height          =   780
            Left            =   150
            TabIndex        =   25
            Top             =   7065
            Width           =   3000
         End
         Begin VB.Label lblTEG 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEG"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0039C6F7&
            Height          =   495
            Left            =   600
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblNET 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Net"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000084&
            Height          =   495
            Left            =   1635
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblDesarrolladores 
            BackStyle       =   0  'Transparent
            Caption         =   "Emiliano Cavia Guillermo Giannotti Guido Pons     Javier Rebagliatti Ariel Clocchiatti"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            Left            =   480
            TabIndex        =   20
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label lblAgradecimientos 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Agradecimientos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   6600
            Width           =   2775
         End
         Begin VB.Label lblEMail 
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   5400
            Width           =   495
         End
         Begin VB.Label lblVisitenos 
            BackStyle       =   0  'Transparent
            Caption         =   "Visítenos en:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   4680
            Width           =   2895
         End
         Begin VB.Label lblDisclaimer 
            BackStyle       =   0  'Transparent
            Caption         =   "TegNet es una aplicación Freeware y puede ser distribuido libremente."
            Height          =   495
            Left            =   120
            TabIndex        =   15
            Top             =   17535
            Width           =   2895
         End
         Begin VB.Label LblVersion 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Versión 1.0.0 - 18/09/2001"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label lblSlogan 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Estrategias sin límites."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label lblDescCopyright 
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright © 2001 TNT Software"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   17085
            Width           =   2895
         End
         Begin VB.Label lblCopyright 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   16545
            Width           =   2775
         End
         Begin VB.Label lblEmailAddress 
            BackStyle       =   0  'Transparent
            Caption         =   "info@tegnet.com.ar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   6
            Top             =   5400
            Width           =   2295
         End
         Begin VB.Label lblWebAddress 
            BackStyle       =   0  'Transparent
            Caption         =   "http://www.tegnet.com.ar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   4920
            Width           =   2895
         End
         Begin VB.Label lblDesarrolladoPor 
            BackStyle       =   0  'Transparent
            Caption         =   "Desarrolado por:"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   2640
            Width           =   2895
         End
         Begin VB.Label lblCreditos 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Créditos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label lblTEG2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEG"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   630
            TabIndex        =   23
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label lblNet2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Net"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   735
            Left            =   1665
            TabIndex        =   24
            Top             =   270
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    On Error GoTo ErrorHandle
    
    Unload Me
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdAceptar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub



Private Sub cmdStopPlay_Click()
    On Error GoTo ErrorHandle
    
    If Timer1.Interval = 0 Then
        Timer1.Interval = 10
        cmdStopPlay.Picture = ImageList1.ListImages(1).Picture
    Else
        Timer1.Interval = 0
        cmdStopPlay.Picture = ImageList1.ListImages(2).Picture
    End If
    
    'Le pasa el foco a la barra de desplazamiento, para que no quede con la linea punteada del foco que no permite ver la imagen del boton.
    HScroll1.SetFocus

    Exit Sub
ErrorHandle:
    ReportErr "cmdStopPlay_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    'Regional
    Me.Caption = ObtenerTextoRecurso(CintCreditosCaption)
    Me.cmdAceptar.Caption = ObtenerTextoRecurso(CintCreditosAceptar)
    Me.lblSlogan.Caption = ObtenerTextoRecurso(CintCreditosSlogan)
    Me.lblCreditos.Caption = ObtenerTextoRecurso(CintCreditosTitCreditos)
    Me.lblDesarrolladoPor.Caption = ObtenerTextoRecurso(CintCreditosDesarrolladoPor)
    Me.lblVisitenos.Caption = ObtenerTextoRecurso(CintCreditosVisitenos)
    Me.lblEMail.Caption = ObtenerTextoRecurso(CintCreditosEMail)
    Me.lblAgradecimientos.Caption = ObtenerTextoRecurso(CintCreditosTitAgradecimientos)
    Me.lblAgradecimiento1.Caption = ObtenerTextoRecurso(CintCreditosAgradecimiento1)
    Me.lblAgradecimiento2.Caption = ObtenerTextoRecurso(CintCreditosAgradecimiento2)
    Me.lblAgradecimiento3.Caption = ObtenerTextoRecurso(CintCreditosAgradecimiento3)
    Me.lblAgradecimiento4.Caption = ObtenerTextoRecurso(CintCreditosAgradecimiento4)
    Me.lblAgradecimiento5.Caption = ObtenerTextoRecurso(CintCreditosAgradecimiento5)
    Me.lblAgradecimiento6.Caption = ObtenerTextoRecurso(CintCreditosAgradecimiento6)
    Me.lblAgradecimiento7.Caption = ObtenerTextoRecurso(CintCreditosAgradecimiento7)
    Me.lblAgradecimiento8.Caption = ObtenerTextoRecurso(CintCreditosAgradecimiento8)
    Me.lblAgradecimiento9.Caption = ObtenerTextoRecurso(CintCreditosAgradecimiento9)
    Me.lblAgradecimiento10.Caption = ObtenerTextoRecurso(CintCreditosAgradecimiento10)
    Me.lblAgradecimiento11.Caption = ObtenerTextoRecurso(CintCreditosAgradecimiento11)
    Me.lblAgradecimiento12.Caption = ObtenerTextoRecurso(CintCreditosAgradecimiento12)
    Me.lblAgradecimientoFinal.Caption = ObtenerTextoRecurso(CintCreditosAgradecimientoFinal)
    Me.lblCopyright.Caption = ObtenerTextoRecurso(CintCreditosTitCopyright)
    Me.lblDescCopyright.Caption = ObtenerTextoRecurso(CintCreditosCopyright)
    Me.lblDisclaimer.Caption = ObtenerTextoRecurso(CintCreditosDisclaimer)
 
    LblVersion = ObtenerTextoRecurso(CintInicioVersion) & " " & App.Major & "." & App.Minor & "." & App.Revision & " - " & Format(FileDateTime(App.Path & "\" & App.EXEName & ".exe"), "dd/mm/yyyy")
    
    lblArgentina.Caption = ObtenerTextoRecurso(CintInicioArgentina)
    
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Picture1.Height = 3600
    Picture1.Width = 3700
    Picture2.Width = Picture1.Width - 400
    Picture2.Left = 200
    Picture2.Top = 3300
    Me.Height = Picture1.Height + Frame3.Height + 700
    Frame3.Left = Picture1.Left
    Frame3.Top = Me.Height - Frame3.Height - 500
    
    cmdStopPlay.Picture = ImageList1.ListImages(1).Picture
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Load", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Timer1_Timer()
    On Error GoTo ErrorHandle
    
    If Picture2.Top < Picture2.Height * -1 Then
        Picture2.Top = 3300
    End If
    
    Picture2.Top = Picture2.Top - (0.5 * HScroll1.Value)
    
    Exit Sub
ErrorHandle:
    ReportErr "Timer1_Timer", Me.Name, Err.Description, Err.Number, Err.Source
End Sub
