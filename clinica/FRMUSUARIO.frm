VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRMUSUARIO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REGISTRO DE USUARIO"
   ClientHeight    =   5940
   ClientLeft      =   2205
   ClientTop       =   2895
   ClientWidth     =   9660
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
   ScaleHeight     =   5940
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox USUARIO 
      Height          =   315
      Left            =   2565
      TabIndex        =   2
      Top             =   930
      Width           =   6450
   End
   Begin VB.TextBox CONTRASEÑA 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2565
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1425
      Width           =   3105
   End
   Begin LVbuttons.LaVolpeButton ELIMINAR 
      Height          =   465
      Left            =   7650
      TabIndex        =   0
      Top             =   3180
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "ELIMINAR"
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
      MICON           =   "FRMUSUARIO.frx":0000
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
   Begin LVbuttons.LaVolpeButton GUARDAR 
      Height          =   465
      Left            =   5730
      TabIndex        =   3
      Top             =   3180
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "GUARDAR"
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
      MICON           =   "FRMUSUARIO.frx":001C
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
   Begin LVbuttons.LaVolpeButton SALIR 
      Cancel          =   -1  'True
      Height          =   465
      Left            =   7635
      TabIndex        =   4
      Top             =   3690
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
      MICON           =   "FRMUSUARIO.frx":0038
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
      Height          =   270
      Left            =   0
      TabIndex        =   7
      Top             =   5670
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8255
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8255
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   9630
      X2              =   435
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      X1              =   420
      X2              =   9660
      Y1              =   615
      Y2              =   615
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
      Index           =   1
      Left            =   1545
      TabIndex        =   6
      Top             =   960
      Width           =   975
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
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   2325
      Left            =   300
      Picture         =   "FRMUSUARIO.frx":0054
      Top             =   -30
      Width           =   1965
   End
   Begin VB.Image Image2 
      Height          =   2325
      Left            =   180
      Picture         =   "FRMUSUARIO.frx":38C1
      Top             =   2325
      Width           =   1965
   End
End
Attribute VB_Name = "FRMUSUARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GUARDAR_Click()
'On Error Resume Next
ssql = "SELECT * FROM USUARIO WHERE USUARIO='" & USUARIO.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    ssql = "UPDATE USUARIO SET CONTRASEÑA='" & CONTRASEÑA.Text & "' WHERE USUARIO='" & USUARIO.Text & "';"
Else
    ssql = "INSERT INTO USUARIO VALUES('" & USUARIO.Text & "','" & CONTRASEÑA.Text & "');"
End If
Set tbl = conn.Execute(ssql)
If Err Then MsgBox Err.Description
Form_Load
End Sub

Private Sub CONTRASEÑA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then GUARDAR_Click
End Sub

Private Sub ELIMINAR_Click()
On Error Resume Next
ssql = "DELETE FROM USUARIO='" & USUARIO.Text & "';"
Set tbl = conn.Execute(ssql)
Form_Load
If Err Then MsgBox Err.Description

End Sub

Private Sub Form_Load()
ssql = "SELECT * FROM USUARIO ORDER BY USUARIO;"
Set tbl = conn.Execute(ssql)
USUARIO.Clear
Do Until tbl.EOF
    USUARIO.AddItem tbl!USUARIO
    tbl.MoveNext
Loop
End Sub

Private Sub SALIR_Click()
Unload Me
End Sub

Private Sub USUARIO_Click()
ssql = "SELECT * FROM USUARIO WHERE USUARIO='" & USUARIO.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    CONTRASEÑA.Text = tbl!CONTRASEÑA
Else
    CONTRASEÑA.Text = ""
End If
End Sub

Private Sub USUARIO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CONTRASEÑA.SetFocus
End If
End Sub
