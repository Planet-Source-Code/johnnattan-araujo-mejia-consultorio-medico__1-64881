VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRMDOCUMENTO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TIPOS DE DOCUMENTOS"
   ClientHeight    =   8865
   ClientLeft      =   1470
   ClientTop       =   1590
   ClientWidth     =   12720
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
   ScaleHeight     =   8865
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5085
      ScaleHeight     =   300
      ScaleWidth      =   5955
      TabIndex        =   12
      Top             =   1620
      Width           =   5955
      Begin VB.OptionButton T 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         Caption         =   "NINGUNO"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3975
         TabIndex        =   15
         Top             =   0
         Width           =   1965
      End
      Begin VB.OptionButton T 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "INGRESO"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1965
      End
      Begin VB.OptionButton T 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         Caption         =   "SALIDA"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   1995
         TabIndex        =   13
         Top             =   0
         Width           =   1965
      End
   End
   Begin VB.OptionButton D 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "NINGUNO"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   9045
      TabIndex        =   11
      Top             =   2220
      Width           =   1965
   End
   Begin VB.OptionButton D 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "SALIDA S/."
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   7050
      TabIndex        =   9
      Top             =   2220
      Width           =   1965
   End
   Begin VB.OptionButton D 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      Caption         =   "INGRESO S/."
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   5070
      TabIndex        =   8
      Top             =   2220
      Width           =   1965
   End
   Begin VB.TextBox CODTIPO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   7455
      TabIndex        =   0
      Top             =   180
      Width           =   2895
   End
   Begin VB.ComboBox TIPO 
      Height          =   315
      Left            =   5085
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   975
      Width           =   5970
   End
   Begin LVbuttons.LaVolpeButton ELIMINAR 
      Height          =   465
      Left            =   9150
      TabIndex        =   3
      Top             =   2625
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
      MICON           =   "FRMDOCUMENTO.frx":0000
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
      Left            =   7230
      TabIndex        =   4
      Top             =   2625
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
      MICON           =   "FRMDOCUMENTO.frx":001C
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
      Left            =   9150
      TabIndex        =   5
      Top             =   3120
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
      MICON           =   "FRMDOCUMENTO.frx":0038
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
      TabIndex        =   16
      Top             =   8595
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10954
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10954
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   12675
      X2              =   0
      Y1              =   3855
      Y2              =   3855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMIENTO DE DINERO"
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
      Index           =   3
      Left            =   5070
      TabIndex        =   10
      Top             =   1980
      Width           =   2655
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMENTO MERCADERIA"
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
      Left            =   5085
      TabIndex        =   7
      Top             =   1365
      Width           =   2805
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DOCUMENTO"
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
      Left            =   5055
      TabIndex        =   6
      Top             =   690
      Width           =   1365
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO DOCUMENTO"
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
      Left            =   5130
      TabIndex        =   1
      Top             =   195
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   3660
      Left            =   -60
      Picture         =   "FRMDOCUMENTO.frx":0054
      Top             =   -60
      Width           =   2835
   End
End
Attribute VB_Name = "frmdocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ACEPTAR_Click()
Select Case Val(CODTIPO.Text)
Case 1, 2, 3, 4, 5, 45, 56
MsgBox "ESTE DOCUMENTO NO SE PUEDE MODIFICAR"
Case Else

    ssql = "SELECT * FROM TIPO WHERE CODTIPO=" & Val(CODTIPO.Text) & ";"
    Set tbl = conn.Execute(ssql)
    Select Case True
    Case T(1).Value
    T1 = 1
    Case T(2).Value
    T1 = 2
    Case T(3).Value
    T1 = 3
    End Select
    Select Case True
    Case D(0).Value
    D1 = 0
    Case D(1).Value
    D1 = 1
    Case D(2).Value
    D1 = 2
    End Select
    If tbl.EOF = False Then
        ssql = "UPDATE TIPO SET TIPO='" & TIPO.Text & "',T=" & Val(T1) & ",D=" & Val(D1) & " WHERE CODTIPO=" & Val(tbl!CODTIPO) & ";"
    Else
        ssql = "INSERT INTO TIPO VALUES(" & Val(CODTIPO.Text) & ",'" & TIPO.Text & "'," & Val(T1) & "," & Val(D1) & ");"
    End If
    Set tbl = conn.Execute(ssql)
    Form_Load
End Select

End Sub

Private Sub CODTIPO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ssql = "SELECT * FROM TIPO WHERE CODTIPO=" & Val(CODTIPO.Text) & ";"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        TIPO.Text = tbl!TIPO
        T(tbl!T).Value = True
        D(tbl!D).Value = True
    Else
        TIPO.Text = ""
        T(1).Value = True
        D(1).Value = True
    End If
    TIPO.SetFocus
End If
End Sub

Private Sub ELIMINAR_Click()
Select Case Val(CODTIPO.Text)
Case 1, 2, 3, 4, 5, 45, 56
    MsgBox "NO SE PUEDEN ELMINAR ESTOS TIPOS DE DOCUMENTO"
Case Else
    ssql = "DELETE FROM TIPO WHERE CODTIPO=" & Val(CODTIPO.Text) & ";"
    Set tbl = conn.Execute(ssql)
    Form_Load
End Select
End Sub

Private Sub Form_Load()
CODTIPO.Text = ""
TIPO.Text = ""
T(1).Value = True
D(0).Value = True
ssql = "SELECT * FROM TIPO ORDER BY TIPO;"
Set tbl = conn.Execute(ssql)
TIPO.Clear
Do Until tbl.EOF
    TIPO.AddItem tbl!TIPO
    tbl.MoveNext
Loop
End Sub

Private Sub SALIR_Click()
Unload Me
End Sub

Private Sub T_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    D(0).SetFocus
End If
End Sub

Private Sub TIPO_Click()
ssql = "SELECT * FROM TIPO WHERE TIPO='" & TIPO.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    CODTIPO.Text = tbl!CODTIPO
End If
CODTIPO_KeyPress 13

End Sub

Private Sub TIPO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    T(1).SetFocus
End If
End Sub
