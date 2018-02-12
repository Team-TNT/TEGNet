VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Chat"
   ClientHeight    =   795
   ClientLeft      =   60
   ClientTop       =   255
   ClientWidth     =   8895
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox txtRecibido 
      Height          =   690
      Left            =   3825
      TabIndex        =   2
      Top             =   0
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   1217
      _Version        =   393217
      BackColor       =   10867440
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":0442
   End
   Begin VB.TextBox txtEnviado 
      BackColor       =   &H00A5D2F0&
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3795
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Default         =   -1  'True
      Height          =   315
      Left            =   2715
      TabIndex        =   0
      Top             =   345
      Width           =   1095
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "Mostrar Bitácora"
      Height          =   240
      Left            =   30
      TabIndex        =   3
      Top             =   345
      Width           =   1800
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnviar_Click()
    On Error GoTo ErrorHandle
    
    'Valida el texto ingresado
    If Not ValidaTexto(txtEnviado.Text, 0) Then
        MsgBox ObtenerTextoRecurso(CintGralMsgCaracterInvalido), vbExclamation, ObtenerTextoRecurso(CintGralMsgCaracterInvalidoCaption)
        Exit Sub
    End If
    
    If Trim(txtEnviado.Text) <> "" Then
        cMensajeChatSaliente Me.txtEnviado.Text
        Me.txtEnviado.Text = ""
        txtEnviado.SetFocus
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "cmdEnviar_Click", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    Me.Caption = " " & ObtenerTextoRecurso(CintChatCaption)
    Me.cmdEnviar.Caption = ObtenerTextoRecurso(CintChatEnviar)
    Me.chkLog.Caption = ObtenerTextoRecurso(CintChatMostrarLog)

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
    mdifrmPrincipal.mnuVerChat.Checked = False
    
    Exit Sub
ErrorHandle:
    ReportErr "Form_QueryUnload", Me.Name, Err.Description, Err.Number, Err.Source
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrorHandle

    'PARTE COMUN
    txtRecibido.Top = 0
    txtEnviado.Left = 0
    cmdEnviar.Width = 1095
    chkLog.Width = 1545
    txtEnviado.Height = 285
    chkLog.Height = 285
    cmdEnviar.Height = 285

    'Segun el tamaño (alto) del Form, se acomoda de una u otra forma.
    'Hay 3 opciones:

    'OPCION 1 y 2
    If Me.ScaleHeight < 995 Then
        txtEnviado.Top = 0
        txtRecibido.Width = Me.ScaleWidth / 2
        txtRecibido.Left = Me.ScaleWidth / 2
        cmdEnviar.Left = MayorACero(txtRecibido.Left - cmdEnviar.Width - 30)
        txtRecibido.Height = Me.ScaleHeight
        chkLog.Left = 15
        chkLog.Top = txtEnviado.Height

        If Me.ScaleHeight < txtEnviado.Height + cmdEnviar.Height Then
            cmdEnviar.Top = 0
            txtEnviado.Width = cmdEnviar.Left - 15
        Else
            cmdEnviar.Top = txtEnviado.Height + 15
            txtEnviado.Width = MayorACero(txtRecibido.Left - 15)
        End If
    End If

    'OPCION 3
    If Me.ScaleHeight >= 995 Then
        txtRecibido.Left = 0
        txtEnviado.Left = 0
        txtRecibido.Width = Me.ScaleWidth
        chkLog.Left = MayorACero(Me.ScaleWidth - chkLog.Width - 30)

        'Si el form es mas Alto que Ancho,
        'el boton y el option van debajo de todo
        'sino, el boton y el option van junto al txt de envio
        If Me.ScaleHeight > Me.ScaleWidth Then
            txtRecibido.Height = MayorACero(Me.ScaleHeight - 15 - txtEnviado.Height - 15 - cmdEnviar.Height - 15)
            txtEnviado.Top = txtRecibido.Height + 15
            cmdEnviar.Left = 15
            txtEnviado.Width = Me.ScaleWidth
            cmdEnviar.Top = txtEnviado.Top + txtEnviado.Height + 15
            chkLog.Top = cmdEnviar.Top
        Else
            txtRecibido.Height = MayorACero(Me.ScaleHeight - 15 - txtEnviado.Height - 15)
            txtEnviado.Top = txtRecibido.Height + 15
            cmdEnviar.Left = MayorACero(chkLog.Left - cmdEnviar.Width - 30)
            txtEnviado.Width = MayorACero(cmdEnviar.Left - 15)
            cmdEnviar.Top = txtEnviado.Top
            chkLog.Top = txtEnviado.Top
        End If
    End If

    Exit Sub
ErrorHandle:
    ReportErr "Form_Resize", Me.Name, Err.Description, Err.Number, Err.Source
End Sub
Private Sub txtEnviado_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    
    If Not ValidaTexto(Chr$(KeyAscii), 0) Then
        KeyAscii = 0
        Beep
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "txtEnviado_KeyPress", Me.Name, Err.Description, Err.Number, Err.Source
End Sub
