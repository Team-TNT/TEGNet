VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMapa 
   BackColor       =   &H00FFE0D0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mapa"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmMapa.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmMapa.frx":0442
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   593
   Begin MSComctlLib.ImageList imgLstIconos 
      Left            =   135
      Top             =   3480
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
            Picture         =   "frmMapa.frx":A7A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapa.frx":A90B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   22
      X1              =   0
      X2              =   28
      Y1              =   62
      Y2              =   74
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   21
      X1              =   546
      X2              =   593
      Y1              =   9
      Y2              =   26
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   20
      X1              =   547
      X2              =   564
      Y1              =   75
      Y2              =   62
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   19
      X1              =   548
      X2              =   561
      Y1              =   17
      Y2              =   27
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   18
      X1              =   194
      X2              =   245
      Y1              =   161
      Y2              =   88
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   17
      X1              =   200
      X2              =   230
      Y1              =   78
      Y2              =   61
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   16
      X1              =   320
      X2              =   359
      Y1              =   126
      Y2              =   128
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   15
      X1              =   303
      X2              =   298
      Y1              =   145
      Y2              =   166
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   14
      X1              =   270
      X2              =   297
      Y1              =   45
      Y2              =   60
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   13
      X1              =   306
      X2              =   310
      Y1              =   81
      Y2              =   103
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   12
      X1              =   321
      X2              =   349
      Y1              =   66
      Y2              =   57
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   11
      X1              =   560
      X2              =   539
      Y1              =   178
      Y2              =   222
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   10
      X1              =   487
      X2              =   499
      Y1              =   244
      Y2              =   204
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   9
      X1              =   0
      X2              =   172
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   8
      X1              =   567
      X2              =   593
      Y1              =   307
      Y2              =   307
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   7
      X1              =   561
      X2              =   574
      Y1              =   280
      Y2              =   259
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   6
      X1              =   532
      X2              =   537
      Y1              =   254
      Y2              =   282
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   5
      X1              =   501
      X2              =   510
      Y1              =   271
      Y2              =   290
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   4
      X1              =   428
      X2              =   440
      Y1              =   248
      Y2              =   185
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   3
      X1              =   416
      X2              =   423
      Y1              =   240
      Y2              =   177
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   2
      X1              =   412
      X2              =   414
      Y1              =   240
      Y2              =   167
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   327
      X2              =   311
      Y1              =   248
      Y2              =   201
   End
   Begin VB.Line lnConector 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   0
      X1              =   263
      X2              =   303
      Y1              =   254
      Y2              =   262
   End
   Begin TegNet.ctlPais objPais 
      Height          =   780
      Index           =   50
      Left            =   7590
      Top             =   4185
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   1376
      ImagenNormal    =   "frmMapa.frx":AA6F
      ImagenSeleccionado=   "frmMapa.frx":ABE0
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   263
      ShapeLeft       =   375
   End
   Begin TegNet.ctlPais objPais 
      Height          =   585
      Index           =   49
      Left            =   8490
      Top             =   3360
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   1032
      ImagenNormal    =   "frmMapa.frx":AD7D
      ImagenSeleccionado=   "frmMapa.frx":AE3E
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   165
      ShapeLeft       =   37
   End
   Begin TegNet.ctlPais objPais 
      Height          =   555
      Index           =   48
      Left            =   7905
      Top             =   3300
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   979
      ImagenNormal    =   "frmMapa.frx":AF2A
      ImagenSeleccionado=   "frmMapa.frx":AFE1
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   100
      ShapeLeft       =   37
   End
   Begin TegNet.ctlPais objPais 
      Height          =   450
      Index           =   47
      Left            =   7170
      Top             =   3645
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   794
      ImagenNormal    =   "frmMapa.frx":B0C3
      ImagenSeleccionado=   "frmMapa.frx":B17B
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   98
      ShapeLeft       =   90
   End
   Begin TegNet.ctlPais objPais 
      Height          =   855
      Index           =   46
      Left            =   6615
      Top             =   4485
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1508
      ImagenNormal    =   "frmMapa.frx":B25E
      ImagenSeleccionado=   "frmMapa.frx":B35A
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   300
      ShapeLeft       =   83
   End
   Begin TegNet.ctlPais objPais 
      Height          =   750
      Index           =   45
      Left            =   5880
      Top             =   4515
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1323
      ImagenNormal    =   "frmMapa.frx":B481
      ImagenSeleccionado=   "frmMapa.frx":B591
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   205
      ShapeLeft       =   190
   End
   Begin TegNet.ctlPais objPais 
      Height          =   645
      Index           =   44
      Left            =   5100
      Top             =   4275
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1138
      ImagenNormal    =   "frmMapa.frx":B6CC
      ImagenSeleccionado=   "frmMapa.frx":B7E8
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   150
      ShapeLeft       =   430
   End
   Begin TegNet.ctlPais objPais 
      Height          =   615
      Index           =   43
      Left            =   5430
      Top             =   4020
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1085
      ImagenNormal    =   "frmMapa.frx":B92F
      ImagenSeleccionado=   "frmMapa.frx":BA61
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   180
      ShapeLeft       =   465
   End
   Begin TegNet.ctlPais objPais 
      Height          =   810
      Index           =   42
      Left            =   5370
      Top             =   3585
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1429
      ImagenNormal    =   "frmMapa.frx":BBAE
      ImagenSeleccionado=   "frmMapa.frx":BD67
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   205
      ShapeLeft       =   670
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1050
      Index           =   41
      Left            =   4485
      Top             =   3660
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   1852
      ImagenNormal    =   "frmMapa.frx":BF1F
      ImagenSeleccionado=   "frmMapa.frx":C0C1
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   355
      ShapeLeft       =   400
   End
   Begin TegNet.ctlPais objPais 
      Height          =   930
      Index           =   40
      Left            =   7365
      Top             =   2145
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1640
      ImagenNormal    =   "frmMapa.frx":C27D
      ImagenSeleccionado=   "frmMapa.frx":C3F9
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   330
      ShapeLeft       =   263
   End
   Begin TegNet.ctlPais objPais 
      Height          =   390
      Index           =   39
      Left            =   6795
      Top             =   2505
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   688
      ImagenNormal    =   "frmMapa.frx":C579
      ImagenSeleccionado=   "frmMapa.frx":C623
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   68
      ShapeLeft       =   75
   End
   Begin TegNet.ctlPais objPais 
      Height          =   375
      Index           =   38
      Left            =   6450
      Top             =   2460
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      ImagenNormal    =   "frmMapa.frx":C6F8
      ImagenSeleccionado=   "frmMapa.frx":C7A5
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   70
      ShapeLeft       =   75
   End
   Begin TegNet.ctlPais objPais 
      Height          =   690
      Index           =   37
      Left            =   8055
      Top             =   2040
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1217
      ImagenNormal    =   "frmMapa.frx":C87D
      ImagenSeleccionado=   "frmMapa.frx":C97A
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   220
      ShapeLeft       =   180
   End
   Begin TegNet.ctlPais objPais 
      Height          =   540
      Index           =   36
      Left            =   6240
      Top             =   2130
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   953
      ImagenNormal    =   "frmMapa.frx":CAA2
      ImagenSeleccionado=   "frmMapa.frx":CBD0
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   75
      ShapeLeft       =   300
   End
   Begin TegNet.ctlPais objPais 
      Height          =   705
      Index           =   35
      Left            =   8355
      Top             =   330
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   1244
      ImagenNormal    =   "frmMapa.frx":CD1B
      ImagenSeleccionado=   "frmMapa.frx":CDEB
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   225
      ShapeLeft       =   40
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1875
      Index           =   34
      Left            =   7350
      Top             =   465
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   3307
      ImagenNormal    =   "frmMapa.frx":CEE6
      ImagenSeleccionado=   "frmMapa.frx":D14D
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   865
      ShapeLeft       =   473
   End
   Begin TegNet.ctlPais objPais 
      Height          =   825
      Index           =   33
      Left            =   6720
      Top             =   1260
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1455
      ImagenNormal    =   "frmMapa.frx":D3E2
      ImagenSeleccionado=   "frmMapa.frx":D51F
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   250
      ShapeLeft       =   325
   End
   Begin TegNet.ctlPais objPais 
      Height          =   600
      Index           =   32
      Left            =   6585
      Top             =   930
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1058
      ImagenNormal    =   "frmMapa.frx":D688
      ImagenSeleccionado=   "frmMapa.frx":D7AB
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   100
      ShapeLeft       =   600
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1380
      Index           =   31
      Left            =   6360
      Top             =   1095
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2434
      ImagenNormal    =   "frmMapa.frx":D8F9
      ImagenSeleccionado=   "frmMapa.frx":DAE8
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   600
      ShapeLeft       =   310
   End
   Begin TegNet.ctlPais objPais 
      Height          =   870
      Index           =   30
      Left            =   6600
      Top             =   225
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1535
      ImagenNormal    =   "frmMapa.frx":DCD4
      ImagenSeleccionado=   "frmMapa.frx":DE4F
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   355
      ShapeLeft       =   750
   End
   Begin TegNet.ctlPais objPais 
      Height          =   600
      Index           =   29
      Left            =   7590
      Top             =   75
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1058
      ImagenNormal    =   "frmMapa.frx":DFF6
      ImagenSeleccionado=   "frmMapa.frx":E0E4
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   100
      ShapeLeft       =   180
   End
   Begin TegNet.ctlPais objPais 
      Height          =   510
      Index           =   28
      Left            =   6840
      Top             =   255
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   900
      ImagenNormal    =   "frmMapa.frx":E1FD
      ImagenSeleccionado=   "frmMapa.frx":E2D5
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   75
      ShapeLeft       =   180
   End
   Begin TegNet.ctlPais objPais 
      Height          =   690
      Index           =   27
      Left            =   6360
      Top             =   240
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1217
      ImagenNormal    =   "frmMapa.frx":E3D8
      ImagenSeleccionado=   "frmMapa.frx":E4FE
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   130
      ShapeLeft       =   250
   End
   Begin TegNet.ctlPais objPais 
      Height          =   675
      Index           =   26
      Left            =   6180
      Top             =   525
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   1191
      ImagenNormal    =   "frmMapa.frx":E63F
      ImagenSeleccionado=   "frmMapa.frx":E736
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   210
      ShapeLeft       =   150
   End
   Begin TegNet.ctlPais objPais 
      Height          =   735
      Index           =   25
      Left            =   4365
      Top             =   2370
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1296
      ImagenNormal    =   "frmMapa.frx":E858
      ImagenSeleccionado=   "frmMapa.frx":E971
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   225
      ShapeLeft       =   205
   End
   Begin TegNet.ctlPais objPais 
      Height          =   675
      Index           =   24
      Left            =   5295
      Top             =   2550
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1191
      ImagenNormal    =   "frmMapa.frx":EAB5
      ImagenSeleccionado=   "frmMapa.frx":EBB9
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   62
      ShapeLeft       =   195
   End
   Begin TegNet.ctlPais objPais 
      Height          =   825
      Index           =   23
      Left            =   4995
      Top             =   2040
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   1455
      ImagenNormal    =   "frmMapa.frx":ECE8
      ImagenSeleccionado=   "frmMapa.frx":EDF7
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   285
      ShapeLeft       =   160
   End
   Begin TegNet.ctlPais objPais 
      Height          =   780
      Index           =   22
      Left            =   5370
      Top             =   1830
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1376
      ImagenNormal    =   "frmMapa.frx":EF31
      ImagenSeleccionado=   "frmMapa.frx":F051
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   265
      ShapeLeft       =   167
   End
   Begin TegNet.ctlPais objPais 
      Height          =   825
      Index           =   21
      Left            =   5565
      Top             =   1740
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1455
      ImagenNormal    =   "frmMapa.frx":F19C
      ImagenSeleccionado=   "frmMapa.frx":F2D7
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   285
      ShapeLeft       =   323
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1485
      Index           =   20
      Left            =   5670
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2619
      ImagenNormal    =   "frmMapa.frx":F43E
      ImagenSeleccionado=   "frmMapa.frx":F646
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   655
      ShapeLeft       =   355
   End
   Begin TegNet.ctlPais objPais 
      Height          =   810
      Index           =   19
      Left            =   5055
      Top             =   450
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1429
      ImagenNormal    =   "frmMapa.frx":F86B
      ImagenSeleccionado=   "frmMapa.frx":F98E
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   278
      ShapeLeft       =   250
   End
   Begin TegNet.ctlPais objPais 
      Height          =   645
      Index           =   18
      Left            =   4350
      Top             =   1545
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   1138
      ImagenNormal    =   "frmMapa.frx":FADC
      ImagenSeleccionado=   "frmMapa.frx":FBB5
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   310
      ShapeLeft       =   130
   End
   Begin TegNet.ctlPais objPais 
      Height          =   390
      Index           =   17
      Left            =   4320
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   688
      ImagenNormal    =   "frmMapa.frx":FCB9
      ImagenSeleccionado=   "frmMapa.frx":FD64
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   45
      ShapeLeft       =   100
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1335
      Index           =   16
      Left            =   2895
      Top             =   75
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      ImagenNormal    =   "frmMapa.frx":FE3A
      ImagenSeleccionado=   "frmMapa.frx":100CB
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   460
      ShapeLeft       =   600
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1770
      Index           =   15
      Left            =   1020
      Top             =   15
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   3122
      ImagenNormal    =   "frmMapa.frx":102DD
      ImagenSeleccionado=   "frmMapa.frx":10637
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   655
      ShapeLeft       =   555
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1380
      Index           =   14
      Left            =   660
      Top             =   495
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   2434
      ImagenNormal    =   "frmMapa.frx":108DC
      ImagenSeleccionado=   "frmMapa.frx":10B37
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   655
      ShapeLeft       =   295
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1125
      Index           =   13
      Left            =   105
      Top             =   1005
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1984
      ImagenNormal    =   "frmMapa.frx":10D25
      ImagenSeleccionado=   "frmMapa.frx":10F40
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   435
      ShapeLeft       =   308
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1020
      Index           =   12
      Left            =   2385
      Top             =   840
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1799
      ImagenNormal    =   "frmMapa.frx":110F2
      ImagenSeleccionado=   "frmMapa.frx":112C6
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   385
      ShapeLeft       =   270
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1005
      Index           =   11
      Left            =   1905
      Top             =   1245
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   1773
      ImagenNormal    =   "frmMapa.frx":11459
      ImagenSeleccionado=   "frmMapa.frx":1165A
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   375
      ShapeLeft       =   443
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1215
      Index           =   10
      Left            =   150
      Top             =   1605
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   2143
      ImagenNormal    =   "frmMapa.frx":1180C
      ImagenSeleccionado=   "frmMapa.frx":11AFD
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   450
      ShapeLeft       =   818
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1200
      Index           =   9
      Left            =   1500
      Top             =   1380
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2117
      ImagenNormal    =   "frmMapa.frx":11D50
      ImagenSeleccionado=   "frmMapa.frx":11FD3
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   550
      ShapeLeft       =   550
   End
   Begin TegNet.ctlPais objPais 
      Height          =   960
      Index           =   8
      Left            =   810
      Top             =   2235
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1693
      ImagenNormal    =   "frmMapa.frx":121E4
      ImagenSeleccionado=   "frmMapa.frx":1247B
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   190
      ShapeLeft       =   610
   End
   Begin TegNet.ctlPais objPais 
      Height          =   855
      Index           =   7
      Left            =   1560
      Top             =   2460
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1508
      ImagenNormal    =   "frmMapa.frx":1267A
      ImagenSeleccionado=   "frmMapa.frx":1283F
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   300
      ShapeLeft       =   400
   End
   Begin TegNet.ctlPais objPais 
      Height          =   615
      Index           =   6
      Left            =   2325
      Top             =   3705
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1085
      ImagenNormal    =   "frmMapa.frx":129B7
      ImagenSeleccionado=   "frmMapa.frx":12C4A
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   180
      ShapeLeft       =   210
   End
   Begin TegNet.ctlPais objPais 
      Height          =   720
      Index           =   5
      Left            =   2310
      Top             =   3135
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1270
      ImagenNormal    =   "frmMapa.frx":1306C
      ImagenSeleccionado=   "frmMapa.frx":13369
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   235
      ShapeLeft       =   225
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1200
      Index           =   4
      Left            =   2535
      Top             =   4200
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   2117
      ImagenNormal    =   "frmMapa.frx":137CC
      ImagenSeleccionado=   "frmMapa.frx":13AD7
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   400
      ShapeLeft       =   25
   End
   Begin TegNet.ctlPais objPais 
      Height          =   720
      Index           =   3
      Left            =   3090
      Top             =   3960
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   1270
      ImagenNormal    =   "frmMapa.frx":13F25
      ImagenSeleccionado=   "frmMapa.frx":1419D
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   300
      ShapeLeft       =   150
   End
   Begin TegNet.ctlPais objPais 
      Height          =   975
      Index           =   2
      Left            =   2775
      Top             =   3270
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1720
      ImagenNormal    =   "frmMapa.frx":145BD
      ImagenSeleccionado=   "frmMapa.frx":149F8
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   360
      ShapeLeft       =   520
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1350
      Index           =   1
      Left            =   2775
      Top             =   4170
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   2381
      ImagenNormal    =   "frmMapa.frx":14F02
      ImagenSeleccionado=   "frmMapa.frx":152BF
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   450
      ShapeLeft       =   100
   End
   Begin TegNet.ctlPais objPais 
      Height          =   1350
      Index           =   0
      Left            =   705
      Top             =   3960
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   2381
      ImagenNormal    =   "frmMapa.frx":15786
      ImagenSeleccionado=   "frmMapa.frx":15B43
      CantTropas      =   0
      Color           =   2
      TropasFijas     =   0
      Nombre          =   ""
      ShapeTop        =   547
      ShapeLeft       =   143
   End
   Begin VB.Shape shpConector 
      BorderColor     =   &H00C00000&
      BorderStyle     =   3  'Dot
      Height          =   1320
      Left            =   5580
      Shape           =   3  'Circle
      Top             =   4020
      Width           =   1485
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dragNdrop As Integer
Dim byPaisOrigen As Byte
Dim dragNdropFicha As Integer

Dim MsngScaleX As Single '//Numero por el que hay que multiplicar las x para evitar el problema de las fuentes grandes
Dim MsngScaleY As Single '//Numero por el que hay que multiplicar las y para evitar el problema de las fuentes grandes

'Paises seleccionados
Private byPaisSeleccionadoOrigen As Byte
Private byPaisSeleccionadoDestino As Byte

'Pais activo (con el mouseover)
Private byPaisActivo As Byte

'Flag que indica si están cargados los paises limitrofes
Private blnLimitrofesCargados As Boolean

Public Enum enuTipoCursor
    tcDefault
    tcProhibido
End Enum

Public Property Get PaisActivo() As Byte
    PaisActivo = byPaisActivo
End Property

Public Property Let PaisActivo(byValor As Byte)
        
    If byValor <> byPaisActivo Then
    
        'Si es mi turno y no hay pausa, efectua el efecto sobre el pais en cuestion.
        If GintColorActual = GintMiColor And GEstadoCliente <> estPartidaPausada Then
            'Si hay otro pais seleccionado, lo oculta
            If byPaisActivo > 0 Then
                objPais(byPaisActivo).MouseOut
            End If
            objPais(byValor).MouseOver
            objPais(byValor).ZOrder 0
        End If
        
        byPaisActivo = byValor
        
        'Se actualiza el formulario de información
        If frmPropiedades.Visible Then
            MostrarPropiedades byPaisActivo
        End If
        
    End If
    
End Property

Public Property Get PaisSeleccionadoOrigen() As Byte
    PaisSeleccionadoOrigen = byPaisSeleccionadoOrigen
End Property

Public Property Let PaisSeleccionadoOrigen(byValor As Byte)
    'Deselecciona el pais destino
    PaisSeleccionadoDestino = 0
    'Deselecciona el anterior
    objPais(byPaisSeleccionadoOrigen).EstaSeleccionado = False
    If byValor = 0 Or byValor = byPaisSeleccionadoOrigen Then
        byPaisSeleccionadoOrigen = 0
        Me.ActualizarIconosMouse
        frmSeleccion.lblDesde.Caption = ""
    Else
        objPais(byValor).EstaSeleccionado = True
        frmSeleccion.lblDesde.Caption = objPais(byValor).Nombre
        byPaisSeleccionadoOrigen = byValor
    End If
    
End Property

Public Property Get PaisSeleccionadoDestino() As Byte
    PaisSeleccionadoDestino = byPaisSeleccionadoDestino
End Property

Public Property Let PaisSeleccionadoDestino(byValor As Byte)
    objPais(byPaisSeleccionadoDestino).EstaSeleccionado = False
    byPaisSeleccionadoDestino = byValor
    objPais(byPaisSeleccionadoDestino).EstaSeleccionado = True
    frmSeleccion.lblHasta.Caption = objPais(byValor).Nombre
End Property

Public Sub ActualizarIconosMouse()
    'Actualiza los iconos del mouse según el estado del cliente
    Dim byPais As Byte
    
    Select Case GEstadoCliente
        Case enuEstadoCli.estAgregando
            For byPais = 1 To objPais.Count - 1
                'Si el pais es mio
                If objPais(byPais).Color = GintMiColor Then
                    CambiarCursor tcDefault, byPais
                Else
                    CambiarCursor tcProhibido, byPais
                End If
            Next
        
        Case enuEstadoCli.estAtacando
            If byPaisSeleccionadoOrigen = 0 Then
                'Si todavía no se seleccionó el origen
                For byPais = 1 To objPais.Count - 1
                    'Si el pais es mio y tiene mas de una tropa
                    If objPais(byPais).Color = GintMiColor And objPais(byPais).CantTropas > 1 Then
                        CambiarCursor tcDefault, byPais
                    Else
                        CambiarCursor tcProhibido, byPais
                    End If
                Next
            Else
                If byPaisSeleccionadoDestino = 0 Then
                    'Si solo está seleccionado el origen
                    For byPais = 1 To objPais.Count - 1
                        'Si el pais es limitrofe
                        If objPais(byPaisSeleccionadoOrigen).EsLimitrofe(byPais) Then
                            CambiarCursor tcDefault, byPais
                        Else
                            CambiarCursor tcProhibido, byPais
                        End If
                    Next
                Else
                    'Están los dos seleccionados
                    For byPais = 1 To objPais.Count - 1
                        'Solo pueden seleccionarse el origen y el destino
                        'y aquellos que sean mios con mas de una tropa
                        If byPais = byPaisSeleccionadoDestino Or byPais = byPaisSeleccionadoOrigen _
                        Or (objPais(byPais).Color = GintMiColor And objPais(byPais).CantTropas > 1) Then
                            CambiarCursor tcDefault, byPais
                        Else
                            CambiarCursor tcProhibido, byPais
                        End If
                    Next
                End If
            End If
        
        Case enuEstadoCli.estMoviendo
            If byPaisSeleccionadoOrigen = 0 Then
                'Si todavía no se seleccionó el origen
                For byPais = 1 To objPais.Count - 1
                    'Si el pais es mio y tiene mas de una tropa
                    If objPais(byPais).Color = GintMiColor And objPais(byPais).CantTropas > 1 Then
                        CambiarCursor tcDefault, byPais
                    Else
                        CambiarCursor tcProhibido, byPais
                    End If
                Next
            Else
                If byPaisSeleccionadoDestino = 0 Then
                    'Si solo está seleccionado el origen
                    For byPais = 1 To objPais.Count - 1
                        'Si el pais es limitrofe y es mio
                        If objPais(byPaisSeleccionadoOrigen).EsLimitrofe(byPais) And objPais(byPais).Color = GintMiColor Then
                            CambiarCursor tcDefault, byPais
                        Else
                            CambiarCursor tcProhibido, byPais
                        End If
                    Next
                Else
                    'Están los dos seleccionados
                    For byPais = 1 To objPais.Count - 1
                        'Solo pueden seleccionarse el origen y el destino
                        'y aquellos que sean mios con mas de una tropa
                        If byPais = byPaisSeleccionadoDestino Or byPais = byPaisSeleccionadoOrigen _
                        Or (objPais(byPais).Color = GintMiColor And objPais(byPais).CantTropas > 1) Then
                            CambiarCursor tcDefault, byPais
                        Else
                            CambiarCursor tcProhibido, byPais
                        End If
                    Next
                End If
            End If
        
        Case Else
            For byPais = 1 To objPais.Count - 1
                CambiarCursor tcDefault, byPais
            Next
    End Select
End Sub

Private Sub ActualizarMenuAccion()
    'Actualiza las opciones del menú accion, el popup, y la barra de herramientas.
    Dim blnMover As Boolean
    Dim blnAtacar As Boolean
    
    If GEstadoCliente = estAtacando Then
        'Solo se actualiza en el estado atacando, de acuerdo al origen y destino seleccionados.
        If byPaisSeleccionadoDestino = 0 Then
            blnMover = False
            blnAtacar = False
        Else
            'de acuerdo al color del pais destino
            If objPais(byPaisSeleccionadoDestino).Color = GintMiColor Then
                blnMover = True
                blnAtacar = False
            Else
                blnAtacar = True
                blnMover = False
            End If
        End If
        
        mdifrmPrincipal.mnuAtacar.Enabled = blnAtacar
        mdifrmPrincipal.mnuMover.Enabled = blnMover
        mdifrmPrincipal.mnuMoverTodas.Enabled = blnMover
        mdifrmPrincipal.mnuJuegoAtacar.Enabled = blnAtacar
        mdifrmPrincipal.mnuJuegoMover.Enabled = blnMover
        mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbAtacar).Enabled = blnAtacar
        mdifrmPrincipal.Toolbar1.Buttons(enuToolBar.tbMover).Enabled = blnMover
    End If
End Sub

Private Sub Form_Click()
    'Deseleccionar los paises origen y destino
    Me.PaisSeleccionadoOrigen = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    'Regional
    Me.Caption = ObtenerTextoRecurso(CintMapaCaption)
    
    '//Obtiene scaleX y scaleY
    MsngScaleX = 1 - (objPais(0).Width - ScaleX(objPais(0).Ancho)) / objPais(0).Width
    MsngScaleY = 1 - (objPais(0).Height - ScaleX(objPais(0).Alto)) / objPais(0).Height
    
    '//Arregla los paises
    For i = 1 To objPais.Count - 1
        objPais(i).Left = objPais(i).Left * MsngScaleX
        objPais(i).Top = objPais(i).Top * MsngScaleY
        objPais(i).PosicionFichaX = objPais(i).PosicionFichaX * MsngScaleX
        objPais(i).PosicionFichaY = objPais(i).PosicionFichaY * MsngScaleY
    Next
    
    '//Arregla los conectores
    For i = 0 To lnConector.Count - 1
        lnConector(i).X1 = lnConector(i).X1 * MsngScaleX
        lnConector(i).X2 = lnConector(i).X2 * MsngScaleX
        lnConector(i).Y1 = lnConector(i).Y1 * MsngScaleY
        lnConector(i).Y2 = lnConector(i).Y2 * MsngScaleY
    Next
    shpConector.Left = shpConector.Left * MsngScaleX
    shpConector.Top = shpConector.Top * MsngScaleY
    shpConector.Width = shpConector.Width * MsngScaleX
    shpConector.Height = shpConector.Height * MsngScaleY
    
    Me.Width = Me.Width * MsngScaleX
    Me.Height = Me.Height * MsngScaleY
    
    'Inicializa variables
    blnLimitrofesCargados = False
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PaisActivo = 0
End Sub

Private Sub objPais_DblClick(Index As Integer)
    On Error GoTo ErrorHandle
    
    If GEstadoCliente = estAgregando Then
        cAgregarTropas CByte(Index), 1
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "objPais_DblClick", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub objPais_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case GEstadoCliente
    
        Case enuEstadoCli.estAgregando
            If objPais(Index).PuedeSerSeleccionado Then
                If Not (Button = vbRightButton And PaisSeleccionadoOrigen = Index) Then
                    'Solo se selecciona/deselecciona si se hizo click con el botón izquierdo
                    ' o si el pais seleccionado es distinto al pais clickeado.
                    PaisSeleccionadoOrigen = Index
                End If
            
                'De acuerdo al botón muestra o no el popup
                If Button = vbRightButton And byPaisActivo <> 0 Then
                    PopupMenu mdifrmPrincipal.mnuAgregar
                End If
            End If
        
        Case enuEstadoCli.estAtacando
            If objPais(Index).PuedeSerSeleccionado Then
                If byPaisSeleccionadoOrigen <> 0 Then
                    'Si está seleccionado el origen
                    If byPaisSeleccionadoDestino <> 0 Then
                        'Está seleccionado el origen y el destino,
                        If Button = vbLeftButton Then
                            If objPais(Index).Color = GintMiColor Then
                                'vuelve a seleccionar el origen
                                PaisSeleccionadoOrigen = Index
                                ActualizarMenuAccion
                                ActualizarIconosMouse
                            End If
                        Else
                            If Index = byPaisSeleccionadoOrigen Or _
                               Index = byPaisSeleccionadoDestino Then
                                'Si clickeo sobre algun seleccionado, muestra las
                                'opciones
                                PopupMenu mdifrmPrincipal.mnuAccion
                            End If
                        End If
                    Else
                        'Solo está seleccionado el origen,
                        'selecciona el destino
                        PaisSeleccionadoDestino = Index
                        ActualizarMenuAccion
                        ActualizarIconosMouse
                        If Button = vbRightButton Then
                            PopupMenu mdifrmPrincipal.mnuAccion
                        End If
                    End If
                Else
                    'No está seleccionado el origen
                    PaisSeleccionadoOrigen = Index
                    ActualizarMenuAccion
                    ActualizarIconosMouse
                End If
            End If
        
        Case enuEstadoCli.estMoviendo
            If objPais(Index).PuedeSerSeleccionado Then
                If byPaisSeleccionadoOrigen <> 0 Then
                    'Si está seleccionado el origen
                    If byPaisSeleccionadoDestino <> 0 Then
                        'Está seleccionado el origen y el destino,
                        If Button = vbLeftButton Then
                            'vuelve a seleccionar el origen
                            PaisSeleccionadoOrigen = Index
                            ActualizarIconosMouse
                        Else
                            If Index = byPaisSeleccionadoOrigen Or _
                               Index = byPaisSeleccionadoDestino Then
                                'Si clickeo sobre algun seleccionado, muestra las
                                'opciones
                                PopupMenu mdifrmPrincipal.mnuAccion
                            End If
                        End If
                    Else
                        'Solo está seleccionado el origen,
                        'selecciona el destino
                        PaisSeleccionadoDestino = Index
                        ActualizarIconosMouse
                        If Button = vbRightButton Then
                            PopupMenu mdifrmPrincipal.mnuAccion
                        End If
                    End If
                Else
                    'No está seleccionado el origen
                    PaisSeleccionadoOrigen = Index
                    ActualizarIconosMouse
                End If
            End If
        
    End Select

End Sub

Private Sub objPais_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandle
    
        PaisActivo = Index
    
    Exit Sub
ErrorHandle:
    ReportErr "objPais_MouseMove", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub CambiarCursor(intCursor As enuTipoCursor, byPais As Byte)
    On Error GoTo ErrorHandle
        Select Case intCursor
            Case enuTipoCursor.tcDefault
                objPais(byPais).MousePointer = vbDefault
                objPais(byPais).PuedeSerSeleccionado = True
            Case enuTipoCursor.tcProhibido
                objPais(byPais).MousePointer = vbCustom
                Set objPais(byPais).MouseIcon = imgLstIconos.ListImages(2).Picture
                objPais(byPais).PuedeSerSeleccionado = False
        End Select
    
    Exit Sub
ErrorHandle:
    ReportErr "CambiarCursor", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub OcultarImgPais(byPais)
    On Error GoTo ErrorHandle
    
    Dim i As Integer
    objPais(byPais).Visible = False

    Exit Sub
ErrorHandle:
    ReportErr "OcultarImgPais", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Function MostrarPropiedades(byPais As Byte)
    On Error GoTo ErrorHandle
    
    Dim dblZoomX As Double
    Dim dblZoomY As Double
    Dim dblZoom As Double
    
    If byPais = 0 Then
        frmPropiedades.lblNombrePais = ""
        frmPropiedades.lblEjercitos = ""
        frmPropiedades.lblTropasFijas = ""
        frmPropiedades.lblOwner = ""
        frmPropiedades.imgPais.Picture = Nothing
    Else
        'Asigna la imagen
        'Busca un multiplicador para que se mantenga el tamaño
        dblZoomY = 1300 / Me.objPais(byPais).Height
        dblZoomX = 1000 / Me.objPais(byPais).Width
        
        'Se toma el menor
        dblZoom = IIf(dblZoomY > dblZoomX, dblZoomX, dblZoomY)
        
        'Modifica los labels
        frmPropiedades.lblNombrePais = objPais(byPais).Nombre
        frmPropiedades.lblOwner.ForeColor = GvecColores(objPais(byPais).Color)
        frmPropiedades.lblOwner.Caption = GvecJugadores(objPais(byPais).Color).strNombre
        frmPropiedades.lblEjercitos.Caption = objPais(byPais).CantTropas
        frmPropiedades.lblTropasFijas.Caption = objPais(byPais).TropasFijas

        'Cambia el tamaño del picture box para mantener la relación de aspecto
        frmPropiedades.imgPais.Width = Me.objPais(byPais).Width * dblZoom
        frmPropiedades.imgPais.Height = Me.objPais(byPais).Height * dblZoom
        
        frmPropiedades.imgPais.Picture = Me.objPais(byPais).ImagenNormal
    End If

    Exit Function
ErrorHandle:
    ReportErr "MostrarPropiedades", Me.Name, Err.Description, Err.Number, Err.Source
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorHandle
    
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
    Me.Visible = False
    mdifrmPrincipal.mnuVerMapa.Checked = False

    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub EfectoDestelloPais(byPais As Byte)
    On Error GoTo ErrorHandle
    Dim sngInicio As Single
    
    objPais(byPais).IniciarDestello
    objPais(byPais).ZOrder 0
    
    Me.Refresh
    
    Pausa CintPausaMsAgregado, False
    
    objPais(byPais).FinalizarDestello
    
    objPais(byPais).Restaurar

    Exit Sub
ErrorHandle:
    ReportErr "EfectoDestelloPais", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub DeseleccionarOrigen()
    On Error GoTo ErrorHandle
    
    PaisSeleccionadoOrigen = 0
    
    Exit Sub
ErrorHandle:
    ReportErr "DeseleccionarOrigen", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub DeseleccionarDestino()
    On Error GoTo ErrorHandle
    
    PaisSeleccionadoDestino = 0
    
    Exit Sub
ErrorHandle:
    ReportErr "DeseleccionarDestino", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub LimpiarMapa()
    'Oculta todos los paises del mapa
    On Error GoTo ErrorHandle
    Dim byPais As Byte
    
    For byPais = 1 To frmMapa.objPais.Count - 1
        frmMapa.objPais(byPais).MouseOut
    Next byPais
    
    Exit Sub
ErrorHandle:
    ReportErr "LimpiarMapa", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub CargarLimitrofesArchivo()
    'Se cargan los paises limitrofes con la información obtenida del archivo de texto.
    'Solo se cargan si no han sido informados por el servidor (por compatibilidad)
    On Error GoTo ErrorHandle
    Dim byPais As Byte
    Dim byPaisDesde As Byte
    Dim byPaisHasta As Byte
    
    'Solo carga los limitrofes del archivo de texto si todavía no están cargados
    If Not blnLimitrofesCargados Then
    
        For byPais = 1 To objPais.Count - 1
            objPais(byPais).LimpiarLimitrofes
        Next
        
        Open App.Path & "\Limites.txt" For Input As #1
        
        While Not EOF(1)
            Input #1, byPaisDesde, byPaisHasta
            objPais(byPaisDesde).AgregarLimitrofe byPaisHasta
        Wend
        
        blnLimitrofesCargados = True
        
        Close #1
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarLimitrofesArchivo", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Public Sub CargarLimitrofesServidor(vecLimitrofes() As String)
    'Carga los paises limitrofes de acuerdo a lo que indica el servidor
    'Se cargan siempre
    On Error GoTo ErrorHandle
    Dim byPais As Byte
    Dim byPaisDesde As Byte
    Dim byPaisHasta As Byte
    Dim i As Integer
    Dim vecPais As Variant
    
    For byPais = 1 To frmMapa.objPais.Count - 1
        frmMapa.objPais(byPais).LimpiarLimitrofes
    Next
    
    For i = LBound(vecLimitrofes) To UBound(vecLimitrofes)
        vecPais = Split(vecLimitrofes(i), ",", 2)
        frmMapa.objPais(CByte(vecPais(0))).AgregarLimitrofe CByte(vecPais(1))
    Next
    
    'Hace una marca para que no levante los limítrofes desde el archivo de texto.
    blnLimitrofesCargados = True
    
    Exit Sub
ErrorHandle:
    ReportErr "CargarLimitrofesServidor", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

