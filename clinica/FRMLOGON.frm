VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRMLOGON 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SISTEMA CLINICO 1.0"
   ClientHeight    =   4230
   ClientLeft      =   1995
   ClientTop       =   3480
   ClientWidth     =   10710
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   10710
   Begin VB.TextBox CONTRASEÑA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00FF0000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   7680
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox USUARIO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7680
      TabIndex        =   0
      Top             =   2295
      Width           =   2895
   End
   Begin LVbuttons.LaVolpeButton INGRESAR 
      Height          =   465
      Left            =   6720
      TabIndex        =   1
      Top             =   3240
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "INGRESAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "FRMLOGON.frx":0000
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   465
      Left            =   8745
      TabIndex        =   2
      Top             =   3240
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "SALIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "FRMLOGON.frx":001C
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   3945
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9181
            Text            =   "SOFTWARE MEDICO"
            TextSave        =   "SOFTWARE MEDICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9181
            Text            =   "CONTACTO:JAMSTAR56@HOTMAIL.COM"
            TextSave        =   "CONTACTO:JAMSTAR56@HOTMAIL.COM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape2 
      Height          =   2775
      Left            =   150
      Top             =   525
      Width           =   3750
   End
   Begin VB.Image Image2 
      Height          =   2745
      Left            =   150
      Picture         =   "FRMLOGON.frx":0038
      Top             =   555
      Width           =   3750
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   5880
      X2              =   3930
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape Shape1 
      Height          =   1020
      Left            =   5655
      Top             =   2100
      Width           =   5070
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   6195
      TabIndex        =   5
      Top             =   2670
      Width           =   1440
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   6660
      TabIndex        =   3
      Top             =   2325
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   2340
      Left            =   5640
      Picture         =   "FRMLOGON.frx":7105
      Top             =   45
      Width           =   4260
   End
End
Attribute VB_Name = "FRMLOGON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CONTRASEÑA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then INGRESAR_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub INGRESAR_Click()
ssql = "SELECT * FROM USUARIO WHERE USUARIO='" & USUARIO.Text & "' AND CONTRASEÑA='" & CONTRASEÑA.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    Me.Hide
Else
    MsgBox "USUARIO O CONTRASEÑA NO EXISTE"
    USUARIO.Text = ""
    CONTRASEÑA.Text = ""
    USUARIO.SetFocus
End If
End Sub

Private Sub LaVolpeButton1_Click()
Unload Me
End Sub

Private Sub USUARIO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CONTRASEÑA.SetFocus
End Sub
