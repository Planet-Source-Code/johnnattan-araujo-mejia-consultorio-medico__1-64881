VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FRMPRODUCTO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGISTRO DE PRODUCTOS"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   1425
   ClientWidth     =   15240
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
   ScaleHeight     =   8595
   ScaleWidth      =   15240
   Begin VB.TextBox CODPRODUCTO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6870
      TabIndex        =   0
      Top             =   60
      Width           =   2895
   End
   Begin VB.ComboBox PRODUCTO 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4470
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   765
      Width           =   10755
   End
   Begin LVbuttons.LaVolpeButton ELIMINAR 
      Height          =   465
      Left            =   13410
      TabIndex        =   1
      Top             =   2805
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
      MICON           =   "FRMPRODUCTO.frx":0000
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
      Left            =   11490
      TabIndex        =   2
      Top             =   2805
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
      MICON           =   "FRMPRODUCTO.frx":001C
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
      Left            =   13410
      TabIndex        =   3
      Top             =   3315
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
      MICON           =   "FRMPRODUCTO.frx":0038
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
   Begin MSMask.MaskEdBox PRECIO 
      Height          =   390
      Left            =   12135
      TabIndex        =   4
      Top             =   2235
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   9
      Top             =   8325
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13176
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13176
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4320
      X2              =   9840
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO PRODUCTO"
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
      Left            =   4725
      TabIndex        =   8
      Top             =   75
      Width           =   2100
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTO"
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
      Left            =   3255
      TabIndex        =   7
      Top             =   885
      Width           =   1170
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
      Index           =   6
      Left            =   11310
      TabIndex        =   6
      Top             =   2295
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   -1095
      Picture         =   "FRMPRODUCTO.frx":0054
      Top             =   -165
      Width           =   4500
   End
End
Attribute VB_Name = "FRMPRODUCTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ACEPTAR_Click()
If PRODUCTO.Text = "" Then MsgBox "DEBE INGRESAR EL NOMBRE DEL PRODUCTO": Exit Sub
ssql = "SELECT * FROM PRODUCTO WHERE CODPRODUCTO='" & CODPRODUCTO.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    ssql = "UPDATE PRODUCTO SET PRODUCTO='" & PRODUCTO.Text & "',PRECIO=" & Val(PRECIO.Text) & " WHERE CODPRODUCTO='" & Val(CODPRODUCTO.Text) & "';"
Else
    ssql = "INSERT INTO PRODUCTO VALUES('" & CODPRODUCTO.Text & "','" & PRODUCTO.Text & "'," & Val(PRECIO.Text) & ");"
End If
Set tbl = conn.Execute(ssql)
Form_Load
End Sub

Private Sub CODPRODUCTO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ssql = "SELECT * FROM PRODUCTO WHERE CODPRODUCTO='" & CODPRODUCTO.Text & "';"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        PRODUCTO.Text = tbl!PRODUCTO
        PRECIO.Text = tbl!PRECIO
        
    Else
        PRODUCTO.Text = ""
        PRECIO.Text = ""
    End If
    PRODUCTO.SetFocus
End If
End Sub

Private Sub ELIMINAR_Click()
ssql = "DELETE FROM PRODUCTO WHERE CODPRODUCTO='" & CODPRODUCTO.Text & "';"
Set tbl = conn.Execute(ssql)
Form_Load

End Sub

Private Sub Form_Load()
PRODUCTO.Text = ""
CODPRODUCTO.Text = ""
PRECIO.Text = ""
ssql = "SELECT * FROM PRODUCTO ORDER BY PRODUCTO;"
Set tbl = conn.Execute(ssql)
PRODUCTO.Clear
Do Until tbl.EOF
    PRODUCTO.AddItem tbl!PRODUCTO
    tbl.MoveNext
Loop

End Sub

Private Sub PRECIO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ACEPTAR_Click
End Sub

Private Sub PRODUCTO_Click()
ssql = "SELECT * FROM PRODUCTO WHERE PRODUCTO='" & PRODUCTO.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    CODPRODUCTO.Text = tbl!CODPRODUCTO
    PRECIO.Text = tbl!PRECIO
End If
    
End Sub

Private Sub PRODUCTO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then PRECIO.SetFocus
End Sub

Private Sub SALIR_Click()
Unload Me
End Sub
