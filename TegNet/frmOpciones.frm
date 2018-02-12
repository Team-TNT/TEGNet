VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOpciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRecuperarDefault 
      Height          =   255
      Left            =   195
      TabIndex        =   70
      Top             =   3255
      Width           =   285
   End
   Begin VB.CommandButton cmdGuardarDefault 
      Height          =   255
      Left            =   195
      TabIndex        =   23
      Top             =   3585
      Width           =   285
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   330
      Left            =   2580
      TabIndex        =   24
      Top             =   3945
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   3990
      TabIndex        =   25
      Top             =   3945
      Width           =   1230
   End
   Begin VB.Frame fraSolapas 
      Height          =   2500
      Index           =   0
      Left            =   300
      TabIndex        =   26
      Top             =   510
      Visible         =   0   'False
      Width           =   4800
      Begin VB.TextBox txtTolerancia 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2775
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1395
         Width           =   375
      End
      Begin MSComCtl2.UpDown updTolerancia 
         Height          =   315
         Left            =   3150
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1395
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         BuddyControl    =   "txtTolerancia"
         BuddyDispid     =   196614
         OrigLeft        =   2670
         OrigTop         =   585
         OrigRight       =   2865
         OrigBottom      =   855
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updDuracion 
         Height          =   315
         Left            =   3150
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   690
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtDuracion"
         BuddyDispid     =   196615
         OrigLeft        =   2685
         OrigTop         =   210
         OrigRight       =   2880
         OrigBottom      =   495
         Max             =   500
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtDuracion 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2775
         MaxLength       =   3
         TabIndex        =   1
         Top             =   690
         Width           =   375
      End
      Begin VB.Label lblTurnoTolerancia 
         Alignment       =   1  'Right Justify
         Caption         =   "Tolerancia (seg):"
         Height          =   255
         Left            =   135
         TabIndex        =   28
         Top             =   1425
         Width           =   2550
      End
      Begin VB.Label lblTurnoDuracion 
         Alignment       =   1  'Right Justify
         Caption         =   "Duración del Turno (min):"
         Height          =   240
         Left            =   135
         TabIndex        =   27
         Top             =   750
         Width           =   2550
      End
   End
   Begin VB.Frame fraSolapas 
      Height          =   2500
      Index           =   3
      Left            =   300
      TabIndex        =   61
      Top             =   510
      Visible         =   0   'False
      Width           =   4800
      Begin VB.TextBox txtBonusCanjeIncremento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4095
         MaxLength       =   2
         TabIndex        =   17
         Top             =   1770
         Width           =   390
      End
      Begin VB.TextBox txtBonusCanje3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4095
         MaxLength       =   2
         TabIndex        =   16
         Top             =   1340
         Width           =   390
      End
      Begin VB.TextBox txtBonusCanje2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4095
         MaxLength       =   2
         TabIndex        =   15
         Top             =   910
         Width           =   390
      End
      Begin VB.TextBox txtBonusCanje1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4095
         MaxLength       =   2
         TabIndex        =   14
         Top             =   480
         Width           =   390
      End
      Begin MSComCtl2.UpDown UpDBonusCanje1 
         Height          =   285
         Left            =   4485
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   480
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtBonusCanje1"
         BuddyDispid     =   196621
         OrigLeft        =   2535
         OrigTop         =   480
         OrigRight       =   2730
         OrigBottom      =   765
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDBonusCanje2 
         Height          =   285
         Left            =   4485
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   915
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtBonusCanje2"
         BuddyDispid     =   196620
         OrigLeft        =   2535
         OrigTop         =   900
         OrigRight       =   2730
         OrigBottom      =   1185
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDBonusCanje3 
         Height          =   285
         Left            =   4485
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1335
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtBonusCanje3"
         BuddyDispid     =   196619
         OrigLeft        =   2535
         OrigTop         =   1320
         OrigRight       =   2730
         OrigBottom      =   1605
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDBonusCanjeIncremento 
         Height          =   285
         Left            =   4485
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1770
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtBonusCanjeIncremento"
         BuddyDispid     =   196618
         OrigLeft        =   2535
         OrigTop         =   1770
         OrigRight       =   2730
         OrigBottom      =   2055
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCanjeIncremento 
         Caption         =   "Incremento del Bonus de los Canjes Siguientes:"
         Height          =   240
         Left            =   150
         TabIndex        =   69
         Top             =   1770
         Width           =   3915
      End
      Begin VB.Label lblCanjeSegundo 
         Caption         =   "Bonus del Segundo Canje:"
         Height          =   240
         Left            =   150
         TabIndex        =   65
         Top             =   915
         Width           =   3915
      End
      Begin VB.Label lblCanjePrimero 
         Caption         =   "Bonus del Primer Canje:"
         Height          =   240
         Left            =   150
         TabIndex        =   63
         Top             =   480
         Width           =   3915
      End
      Begin VB.Label lblCanjeTercero 
         Caption         =   "Bonus del Tercer Canje:"
         Height          =   240
         Left            =   150
         TabIndex        =   67
         Top             =   1335
         Width           =   3915
      End
   End
   Begin VB.Frame fraSolapas 
      Height          =   2500
      Index           =   5
      Left            =   300
      TabIndex        =   58
      Top             =   510
      Visible         =   0   'False
      Width           =   4800
      Begin VB.TextBox txtTropasInicio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   22
         Top             =   540
         Width           =   390
      End
      Begin MSComCtl2.UpDown UpDTropasInicio 
         Height          =   315
         Left            =   4290
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   540
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtTropasInicio"
         BuddyDispid     =   196626
         OrigLeft        =   2685
         OrigTop         =   210
         OrigRight       =   2880
         OrigBottom      =   495
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTropasInicio 
         Caption         =   "Cantidad de Tropas por País en el reparto inicial:"
         Height          =   420
         Left            =   210
         TabIndex        =   60
         Top             =   585
         Width           =   3615
      End
   End
   Begin VB.Frame fraSolapas 
      Height          =   2500
      Index           =   4
      Left            =   300
      TabIndex        =   54
      Top             =   510
      Visible         =   0   'False
      Width           =   4800
      Begin VB.TextBox txtObjetivoComun 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2145
         MaxLength       =   2
         TabIndex        =   21
         Top             =   1635
         Width           =   390
      End
      Begin VB.OptionButton optMisiones 
         Caption         =   "Jugar con Misiones"
         Height          =   285
         Left            =   435
         TabIndex        =   19
         Top             =   855
         Width           =   3060
      End
      Begin VB.OptionButton optConquistarMundo 
         Caption         =   "Jugar a conquistar el Mundo"
         Height          =   285
         Left            =   435
         TabIndex        =   18
         Top             =   330
         Width           =   3060
      End
      Begin VB.CheckBox chkDestruir 
         Caption         =   "Incluir misión de Destruir"
         Height          =   375
         Left            =   825
         TabIndex        =   20
         Top             =   1245
         Width           =   2910
      End
      Begin MSComCtl2.UpDown UpDObjetivoComun 
         Height          =   315
         Left            =   2535
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1635
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtObjetivoComun"
         BuddyDispid     =   196628
         OrigLeft        =   2685
         OrigTop         =   210
         OrigRight       =   2880
         OrigBottom      =   495
         Max             =   50
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblObjetivoComun2 
         Caption         =   "países"
         Height          =   240
         Left            =   2835
         TabIndex        =   57
         Top             =   1695
         Width           =   975
      End
      Begin VB.Label lblObjetivoComun 
         Alignment       =   1  'Right Justify
         Caption         =   "Objetivo Común:"
         Height          =   285
         Left            =   420
         TabIndex        =   55
         Top             =   1680
         Width           =   1620
      End
   End
   Begin VB.Frame fraSolapas 
      Height          =   2500
      Index           =   2
      Left            =   300
      TabIndex        =   38
      Top             =   510
      Visible         =   0   'False
      Width           =   4800
      Begin VB.Frame fraBonusContinente 
         Caption         =   "Bonus por Continente"
         Height          =   1560
         Left            =   180
         TabIndex        =   41
         Top             =   690
         Width           =   4440
         Begin VB.TextBox txtANorte 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1575
            MaxLength       =   2
            TabIndex        =   9
            Top             =   713
            Width           =   390
         End
         Begin VB.TextBox txtASur 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1575
            MaxLength       =   2
            TabIndex        =   10
            Top             =   1088
            Width           =   390
         End
         Begin VB.TextBox txtAfrica 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1575
            MaxLength       =   2
            TabIndex        =   8
            Top             =   360
            Width           =   390
         End
         Begin VB.TextBox txtOceania 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3660
            MaxLength       =   2
            TabIndex        =   13
            Top             =   1088
            Width           =   390
         End
         Begin VB.TextBox txtEuropa 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3660
            MaxLength       =   2
            TabIndex        =   12
            Top             =   713
            Width           =   390
         End
         Begin VB.TextBox txtAsia 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3660
            MaxLength       =   2
            TabIndex        =   11
            Top             =   338
            Width           =   390
         End
         Begin MSComCtl2.UpDown UpDAsia 
            Height          =   315
            Left            =   4050
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   345
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtAsia"
            BuddyDispid     =   196640
            OrigLeft        =   3780
            OrigTop         =   360
            OrigRight       =   3975
            OrigBottom      =   645
            Max             =   99
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDEuropa 
            Height          =   315
            Left            =   4050
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   720
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtEuropa"
            BuddyDispid     =   196639
            OrigLeft        =   2685
            OrigTop         =   210
            OrigRight       =   2880
            OrigBottom      =   495
            Max             =   99
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDOceania 
            Height          =   315
            Left            =   4050
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1095
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtOceania"
            BuddyDispid     =   196638
            OrigLeft        =   3780
            OrigTop         =   1110
            OrigRight       =   3975
            OrigBottom      =   1395
            Max             =   99
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDAfrica 
            Height          =   315
            Left            =   1965
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   345
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtAfrica"
            BuddyDispid     =   196637
            OrigLeft        =   1860
            OrigTop         =   353
            OrigRight       =   2055
            OrigBottom      =   638
            Max             =   99
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDASur 
            Height          =   315
            Left            =   1965
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1095
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtASur"
            BuddyDispid     =   196636
            OrigLeft        =   1860
            OrigTop         =   1103
            OrigRight       =   2055
            OrigBottom      =   1388
            Max             =   99
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDANorte 
            Height          =   315
            Left            =   1965
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   720
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtANorte"
            BuddyDispid     =   196635
            OrigLeft        =   1860
            OrigTop         =   728
            OrigRight       =   2055
            OrigBottom      =   1013
            Max             =   99
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lblOceania 
            Caption         =   "Oceania:"
            Height          =   270
            Left            =   2670
            TabIndex        =   53
            Top             =   1110
            Width           =   990
         End
         Begin VB.Label lblEuropa 
            Caption         =   "Europa:"
            Height          =   270
            Left            =   2670
            TabIndex        =   52
            Top             =   735
            Width           =   990
         End
         Begin VB.Label lblAsia 
            Caption         =   "Asia:"
            Height          =   270
            Left            =   2670
            TabIndex        =   51
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lblASur 
            Caption         =   "América del Sur:"
            Height          =   270
            Left            =   120
            TabIndex        =   50
            Top             =   1110
            Width           =   1428
         End
         Begin VB.Label lblANorte 
            Caption         =   "América del Norte:"
            Height          =   270
            Left            =   120
            TabIndex        =   49
            Top             =   735
            Width           =   1428
         End
         Begin VB.Label lblAfrica 
            Caption         =   "África:"
            Height          =   270
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   1428
         End
      End
      Begin VB.TextBox txtTarjetaPropio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4020
         MaxLength       =   2
         TabIndex        =   7
         Top             =   195
         Width           =   390
      End
      Begin MSComCtl2.UpDown UpDTarjetaPropio 
         Height          =   315
         Left            =   4410
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   195
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtTarjetaPropio"
         BuddyDispid     =   196647
         OrigLeft        =   2685
         OrigTop         =   210
         OrigRight       =   2880
         OrigBottom      =   495
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTarjetasPropio 
         Alignment       =   1  'Right Justify
         Caption         =   "Tropas por Tarjeta de País propio:"
         Height          =   300
         Left            =   150
         TabIndex        =   39
         Top             =   255
         Width           =   3765
      End
   End
   Begin VB.Frame fraSolapas 
      Height          =   2500
      Index           =   1
      Left            =   300
      TabIndex        =   32
      Top             =   510
      Visible         =   0   'False
      Width           =   4800
      Begin VB.Frame fraTipoRonda 
         Caption         =   "Tipo de Ronda"
         Height          =   975
         Left            =   330
         TabIndex        =   37
         Top             =   1215
         Width           =   2685
         Begin VB.OptionButton optRotativa 
            Caption         =   "Primero Rotativo"
            Height          =   315
            Left            =   180
            TabIndex        =   6
            Top             =   525
            Value           =   -1  'True
            Width           =   2265
         End
         Begin VB.OptionButton optFija 
            Caption         =   "Ronda Fija"
            Height          =   270
            Left            =   180
            TabIndex        =   5
            Top             =   255
            Width           =   2265
         End
      End
      Begin VB.TextBox txtSegundaRonda 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2445
         MaxLength       =   2
         TabIndex        =   4
         Top             =   735
         Width           =   390
      End
      Begin VB.TextBox txtPrimeraRonda 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2445
         MaxLength       =   2
         TabIndex        =   3
         Top             =   255
         Width           =   390
      End
      Begin MSComCtl2.UpDown UpDPrimeraRonda 
         Height          =   315
         Left            =   2820
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPrimeraRonda"
         BuddyDispid     =   196653
         OrigLeft        =   2685
         OrigTop         =   210
         OrigRight       =   2880
         OrigBottom      =   495
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDSegundaRonda 
         Height          =   315
         Left            =   2820
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   735
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSegundaRonda"
         BuddyDispid     =   196652
         OrigLeft        =   2685
         OrigTop         =   210
         OrigRight       =   2880
         OrigBottom      =   495
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblSegundaRonda 
         Alignment       =   1  'Right Justify
         Caption         =   "Cantidad Tropas 2da ronda:"
         Height          =   300
         Left            =   60
         TabIndex        =   34
         Top             =   765
         Width           =   2355
      End
      Begin VB.Label lblPrimerRonda 
         Alignment       =   1  'Right Justify
         Caption         =   "Cantidad Tropas 1er ronda:"
         Height          =   300
         Left            =   60
         TabIndex        =   33
         Top             =   300
         Width           =   2355
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3015
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   5318
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Turno"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ronda"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bonus"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Canje"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Misión"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Otras"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRecuperarDefecto 
      Caption         =   "Recuperar valores por Defecto"
      Height          =   240
      Left            =   600
      TabIndex        =   71
      Top             =   3270
      Width           =   3330
   End
   Begin VB.Label lblGuardarDefecto 
      Caption         =   "Guardar como Valores por Defecto"
      Height          =   210
      Left            =   600
      TabIndex        =   31
      Top             =   3630
      Width           =   3390
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim salioPorAceptar As Boolean

'Regional - Mensajes
Dim strMsgDesconectar As String
Dim strMsgDesconectarCaption As String
Dim strMsgErrorCaption As String

Private Sub cmdAceptar_Click()
    On Error GoTo ErrorHandle
    
    'Valida las opciones ingresadas
    If Not ValidaOpciones Then
        Exit Sub
    End If
    
    salioPorAceptar = True
    
    CapturarOpciones GvecOpciones
    
    cEnviarOpciones
    
    If GEstadoCliente = estConectado Then
        cPedirConexionesActuales
    End If
    
    Unload Me
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdAceptar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ErrorHandle
    
    salioPorAceptar = False
    
    Unload Me
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdGuardarDefault_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdGuardarDefault_Click()
    On Error GoTo ErrorHandle
    
    'Valida las opciones ingresadas
    If Not ValidaOpciones Then
        Exit Sub
    End If
    
    'Guarda las opciones actuales como default
    CapturarOpciones GvecOpcionesDefault
    cEnviarOpcionesDefault
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdGuardarDefault_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub cmdRecuperarDefault_Click()
    On Error GoTo ErrorHandle
    
    'Recupera valores por Defecto
    MostrarOpciones GvecOpcionesDefault
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdRecuperarDefault_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrorHandle
    
    salioPorAceptar = False
    
    MostrarOpciones GvecOpciones

    fraSolapas(0).Visible = True
    TabStrip1.Tabs.Item(1).Selected = True
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_Activate", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    'Regional
    Me.Caption = ObtenerTextoRecurso(CintOpcionesCaption)
    Me.lblRecuperarDefecto.Caption = ObtenerTextoRecurso(CintOpcionesRecuperarDefecto)
    Me.lblGuardarDefecto.Caption = ObtenerTextoRecurso(CintOpcionesGuardarDefecto)
    Me.cmdAceptar.Caption = ObtenerTextoRecurso(CintOpcionesAceptar)
    Me.cmdCancelar.Caption = ObtenerTextoRecurso(CintOpcionesCancelar)
    Me.TabStrip1.Tabs(1).Caption = ObtenerTextoRecurso(CintOpcionesTurno)
    Me.lblTurnoDuracion.Caption = ObtenerTextoRecurso(CintOpcionesTurnoDuracion)
    Me.lblTurnoTolerancia.Caption = ObtenerTextoRecurso(CintOpcionesTurnoTolerancia)
    Me.TabStrip1.Tabs(2).Caption = ObtenerTextoRecurso(CintOpcionesRonda)
    Me.lblPrimerRonda.Caption = ObtenerTextoRecurso(CintOpcionesRondaTropasPrimera)
    Me.lblSegundaRonda.Caption = ObtenerTextoRecurso(CintOpcionesRondaTropasSegunda)
    Me.fraTipoRonda.Caption = ObtenerTextoRecurso(CintOpcionesRondaTipo)
    Me.optFija.Caption = ObtenerTextoRecurso(CintOpcionesRondaFija)
    Me.optRotativa.Caption = ObtenerTextoRecurso(CintOpcionesRondaRotativa)
    Me.TabStrip1.Tabs(3).Caption = ObtenerTextoRecurso(CintOpcionesBonus)
    Me.lblTarjetasPropio.Caption = ObtenerTextoRecurso(CintOpcionesBonusPaisPropio)
    Me.fraBonusContinente.Caption = ObtenerTextoRecurso(CintOpcionesBonusContinente)
    Me.lblAfrica.Caption = ObtenerTextoRecurso(CintOpcionesBonusAfrica)
    Me.lblANorte.Caption = ObtenerTextoRecurso(CintOpcionesBonusANorte)
    Me.lblASur.Caption = ObtenerTextoRecurso(CintOpcionesBonusASur)
    Me.lblAsia.Caption = ObtenerTextoRecurso(CintOpcionesBonusAsia)
    Me.lblEuropa.Caption = ObtenerTextoRecurso(CintOpcionesBonusEuropa)
    Me.lblOceania.Caption = ObtenerTextoRecurso(CintOpcionesBonusOceania)
    Me.TabStrip1.Tabs(4).Caption = ObtenerTextoRecurso(CintOpcionesCanje)
    Me.lblCanjePrimero.Caption = ObtenerTextoRecurso(CintOpcionesCanjePrimero)
    Me.lblCanjeSegundo.Caption = ObtenerTextoRecurso(CintOpcionesCanjeSegundo)
    Me.lblCanjeTercero.Caption = ObtenerTextoRecurso(CintOpcionesCanjeTercero)
    Me.lblCanjeIncremento.Caption = ObtenerTextoRecurso(CintOpcionesCanjeIncremento)
    Me.TabStrip1.Tabs(5).Caption = ObtenerTextoRecurso(CintOpcionesMision)
    Me.optConquistarMundo.Caption = ObtenerTextoRecurso(CintOpcionesMisionConquistarMundo)
    Me.optMisiones.Caption = ObtenerTextoRecurso(CintOpcionesMisionMisiones)
    Me.chkDestruir.Caption = ObtenerTextoRecurso(CintOpcionesMisionDestruir)
    Me.lblObjetivoComun.Caption = ObtenerTextoRecurso(CintOpcionesMisionObjetivoComun)
    Me.lblObjetivoComun2.Caption = ObtenerTextoRecurso(CintOpcionesMisionPaises)
    Me.TabStrip1.Tabs(6).Caption = ObtenerTextoRecurso(CintOpcionesOtras)
    Me.lblTropasInicio.Caption = ObtenerTextoRecurso(CintOpcionesOtrasRepartoInicial)
    strMsgDesconectar = ObtenerTextoRecurso(CintOpcionesMsgDesconectar)
    strMsgDesconectarCaption = ObtenerTextoRecurso(CintOpcionesMsgDesconectarCaption)
    strMsgErrorCaption = ObtenerTextoRecurso(CintOpcionesMsgErrorCaption)

    Exit Sub
ErrorHandle:
    ReportErr "Form_Load", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorHandle
    
    'Si se forzó el unload por código no pregunta si desea finalizar
    'la partida
    If GblnSeCierra Then
        Exit Sub
    End If
    
    If Not salioPorAceptar Then
        'Si todavía no seleccionó el color (estado conectado) se desconecta
        If GEstadoCliente <= estConectado Then
            If MsgBox(strMsgDesconectar, vbQuestion + vbYesNo + vbDefaultButton2, strMsgDesconectarCaption) = vbNo Then
                Cancel = 1
                Exit Sub
            Else
                If GsoyAdministrador Then
                    'Baja el servidor
                    cBajarServidor
                Else
                    'Por si las moscas...
                    cDesconectar
                End If
            End If
        End If
    End If
    
    Me.Visible = False
    Cancel = 1
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub optConquistarMundo_Click()
    On Error GoTo ErrorHandle
    
    Me.txtObjetivoComun.Enabled = Not optConquistarMundo.Value
    Me.chkDestruir.Enabled = Not optConquistarMundo.Value
    Me.lblObjetivoComun.Enabled = Not optConquistarMundo.Value
    Me.UpDObjetivoComun.Enabled = Not optConquistarMundo.Value
    Me.lblObjetivoComun2.Enabled = Not optConquistarMundo.Value
    
    Exit Sub
ErrorHandle:
    ReportErr "optConquistarMundo_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub optMisiones_Click()
    On Error GoTo ErrorHandle
    
    Me.txtObjetivoComun.Enabled = optMisiones.Value
    Me.chkDestruir.Enabled = optMisiones.Value
    Me.lblObjetivoComun.Enabled = optMisiones.Value
    Me.UpDObjetivoComun.Enabled = optMisiones.Value
    Me.lblObjetivoComun2.Enabled = optMisiones.Value
    
    Exit Sub
ErrorHandle:
    ReportErr "optMisiones_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub TabStrip1_Click()
    On Error GoTo ErrorHandle
    
    SeleccionarSolapa TabStrip1.SelectedItem.Index
    
    Exit Sub
ErrorHandle:
    ReportErr "TabStrip1_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub SeleccionarSolapa(intSolapa As Integer)
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    For i = 1 To TabStrip1.Tabs.Count
        If i = intSolapa Then
            fraSolapas(i - 1).Visible = True
        Else
            fraSolapas(i - 1).Visible = False
        End If
    Next i
    
    Exit Sub
ErrorHandle:
    ReportErr "SeleccionarSolapa", Me.Name, Err.Description, Err.Number, Err.Source
End Sub


Private Sub txtAfrica_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtAfrica_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtANorte_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtANorte_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtAsia_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtAsia_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtASur_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtASur_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtBonusCanje1_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtBonusCanje1_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub
Private Sub txtBonusCanjeIncremento_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtBonusCanjeIncremento_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtBonusCanje3_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtBonusCanje3_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtBonusCanje2_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtBonusCanje2_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtDuracion_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtDuracion_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtEuropa_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtEuropa_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtObjetivoComun_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtObjetivoComun_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtOceania_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtOceania_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtPrimeraRonda_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtPrimeraRonda_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtSegundaRonda_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtSegundaRonda_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtTarjetaPropio_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtTarjetaPropio_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtTolerancia_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtTolerancia_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub txtTropasInicio_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
        
    If Not ValidaEntero(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
        
    Exit Sub
ErrorHandle:
    ReportErr "txtTropasInicio_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Function ValidaOpciones() As Boolean
    On Error GoTo ErrorHandle
    Dim blnAux As Boolean
    
    blnAux = True
    
    If Not ValidaEntero(txtDuracion.Text, updDuracion.Min, updDuracion.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(1).Selected = True
        txtDuracion.SetFocus
        blnAux = False
    End If
        
    If Not ValidaEntero(txtTolerancia.Text, updTolerancia.Min, updTolerancia.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(1).Selected = True
        txtTolerancia.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtPrimeraRonda.Text, UpDPrimeraRonda.Min, UpDPrimeraRonda.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(2).Selected = True
        txtPrimeraRonda.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtSegundaRonda.Text, UpDSegundaRonda.Min, UpDSegundaRonda.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(2).Selected = True
        txtSegundaRonda.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtTarjetaPropio.Text, UpDTarjetaPropio.Min, UpDTarjetaPropio.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(3).Selected = True
        txtTarjetaPropio.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtAfrica.Text, UpDAfrica.Min, UpDAfrica.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(3).Selected = True
        txtAfrica.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtANorte.Text, UpDANorte.Min, UpDANorte.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(3).Selected = True
        txtANorte.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtASur.Text, UpDASur.Min, UpDASur.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(3).Selected = True
        txtASur.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtAsia.Text, UpDAsia.Min, UpDAsia.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(3).Selected = True
        txtAsia.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtEuropa.Text, UpDEuropa.Min, UpDEuropa.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(3).Selected = True
        txtEuropa.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtOceania.Text, UpDOceania.Min, UpDOceania.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(3).Selected = True
        txtOceania.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtObjetivoComun.Text, UpDObjetivoComun.Min, UpDObjetivoComun.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(5).Selected = True
        txtObjetivoComun.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtTropasInicio.Text, UpDTropasInicio.Min, UpDTropasInicio.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(6).Selected = True
        txtTropasInicio.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtBonusCanje1.Text, UpDBonusCanje1.Min, UpDBonusCanje1.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(4).Selected = True
        txtBonusCanje1.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtBonusCanje2.Text, UpDBonusCanje2.Min, UpDBonusCanje2.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(4).Selected = True
        txtBonusCanje2.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtBonusCanje3.Text, UpDBonusCanje3.Min, UpDBonusCanje3.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(4).Selected = True
        txtBonusCanje3.SetFocus
        blnAux = False
    End If
    
    If Not ValidaEntero(txtBonusCanjeIncremento.Text, UpDBonusCanjeIncremento.Min, UpDBonusCanjeIncremento.Max) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgNumeroInvalido), vbInformation, ObtenerTextoRecurso(CintGralMSgNumeroInvalidoCaption)
        TabStrip1.Tabs(4).Selected = True
        txtBonusCanjeIncremento.SetFocus
        blnAux = False
    End If
    
    ValidaOpciones = blnAux
    
    Exit Function
ErrorHandle:
    ReportErr "ValidaOpciones", Me.Name, Err.Description, Err.Number, Err.Source
    ValidaOpciones = False
End Function
