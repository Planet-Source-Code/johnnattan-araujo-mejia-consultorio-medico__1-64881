VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FRMESPECIALIDAD 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "REGISTRO DE ESPECIALIDADES"
   ClientHeight    =   3360
   ClientLeft      =   2505
   ClientTop       =   3195
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   9180
   Begin MSMask.MaskEdBox PRECIO 
      Height          =   315
      Left            =   6330
      TabIndex        =   5
      Top             =   1455
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox ESPECIALIDAD 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4950
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   855
      Width           =   4170
   End
   Begin VB.TextBox CODESPECIALIDAD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6210
      TabIndex        =   0
      Top             =   135
      Width           =   2895
   End
   Begin LVbuttons.LaVolpeButton ELIMINAR 
      Height          =   465
      Left            =   7200
      TabIndex        =   6
      Top             =   1950
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
      MICON           =   "FRMESPECIALIDAD.frx":0000
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
   Begin LVbuttons.LaVolpeButton ACEPTAR 
      Height          =   465
      Left            =   5280
      TabIndex        =   7
      Top             =   1950
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "ACEPTAR"
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
      MICON           =   "FRMESPECIALIDAD.frx":001C
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
      Left            =   7200
      TabIndex        =   8
      Top             =   2460
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
      MICON           =   "FRMESPECIALIDAD.frx":0038
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
      TabIndex        =   9
      Top             =   3090
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO"
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
      Left            =   5505
      TabIndex        =   4
      Top             =   1500
      Width           =   810
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ESPECIALIDAD"
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
      Left            =   3285
      TabIndex        =   3
      Top             =   885
      Width           =   1590
   End
   Begin VB.Line Line1 
      X1              =   3660
      X2              =   9180
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO ESPECIALIDAD"
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
      Left            =   3630
      TabIndex        =   1
      Top             =   150
      Width           =   2520
   End
   Begin VB.Image Image1 
      Height          =   2910
      Left            =   -15
      Picture         =   "FRMESPECIALIDAD.frx":0054
      Stretch         =   -1  'True
      Top             =   15
      Width           =   2850
   End
End
Attribute VB_Name = "FRMESPECIALIDAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ACEPTAR_Click()
If ESPECIALIDAD.Text = "" Then MsgBox "INGRESE EL NOMBRE DE LA ESPECIALIDAD": Exit Sub
ssql = "SELECT * FROM ESPECIALIDAD WHERE CODESPECIALIDAD=" & Val(CODESPECIALIDAD.Text) & ";"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    ssql = "UPDATE ESPECIALIDAD SET ESPECIALIDAD='" & ESPECIALIDAD.Text & "',PRECIO=" & Val(PRECIO.Text) & " WHERE CODESPECIALIDAD=" & Val(CODESPECIALIDAD.Text) & ""
Else
    ssql = "INSERT INTO ESPECIALIDAD VALUES(" & Val(CODESPECIALIDAD.Text) & ",'" & ESPECIALIDAD.Text & "'," & Val(PRECIO.Text) & ");"
End If
Set tbl = conn.Execute(UCase(ssql))
Form_Load
CODESPECIALIDAD.SetFocus
End Sub

Private Sub CODESPECIALIDAD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ssql = "SELECT * FROM ESPECIALIDAD WHERE CODESPECIALIDAD=" & CODESPECIALIDAD.Text & ";"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        ESPECIALIDAD.Text = tbl!ESPECIALIDAD
        PRECIO.Text = tbl!PRECIO
    Else
        ESPECIALIDAD.Text = ""
        PRECIO.Text = ""
    End If
    ESPECIALIDAD.SetFocus
End If
End Sub

Private Sub ELIMINAR_Click()
ssql = "DELETE FROM ESPECIALIDAD WHERE CODESPECIALIDAD=" & Val(CODESPECIALIDAD.Text) & ";"
Set tbl = conn.Execute(ssql)
Form_Load
End Sub

Private Sub ESPECIALIDAD_Click()
ssql = "SELECT * FROM ESPECIALIDAD WHERE ESPECIALIDAD='" & ESPECIALIDAD.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    CODESPECIALIDAD.Text = tbl!CODESPECIALIDAD
End If
CODESPECIALIDAD_KeyPress 13
PRECIO.SetFocus
End Sub

Private Sub ESPECIALIDAD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then PRECIO.SetFocus
End Sub

Private Sub Form_Load()
CODESPECIALIDAD.Text = ""
ESPECIALIDAD.Text = ""
PRECIO.Text = ""
ssql = "sELECT * FROM ESPECIALIDAD ORDER BY ESPECIALIDAD;"
Set tbl = conn.Execute(ssql)
ESPECIALIDAD.Clear
Do Until tbl.EOF
    ESPECIALIDAD.AddItem tbl!ESPECIALIDAD
    tbl.MoveNext
Loop
End Sub

Private Sub PRECIO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ACEPTAR_Click
End Sub

Private Sub SALIR_Click()
Unload Me
End Sub
