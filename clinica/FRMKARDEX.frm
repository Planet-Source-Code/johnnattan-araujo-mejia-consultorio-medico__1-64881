VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRMKARDEX 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELECCIONADOR DE DOCUMENTO"
   ClientHeight    =   8550
   ClientLeft      =   1980
   ClientTop       =   1755
   ClientWidth     =   12390
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
   ScaleHeight     =   8550
   ScaleWidth      =   12390
   Begin VB.ComboBox DOCUMENTO 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4890
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   870
      Width           =   5610
   End
   Begin LVbuttons.LaVolpeButton REALIZAR 
      Height          =   465
      Left            =   8580
      TabIndex        =   2
      Top             =   1200
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "NUEVO"
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
      MICON           =   "FRMKARDEX.frx":0000
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
   Begin LVbuttons.LaVolpeButton CB 
      Height          =   465
      Index           =   0
      Left            =   3915
      TabIndex        =   4
      Top             =   2700
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "CONSULTA"
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
      MICON           =   "FRMKARDEX.frx":001C
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
   Begin LVbuttons.LaVolpeButton CB 
      Cancel          =   -1  'True
      Height          =   465
      Index           =   2
      Left            =   7755
      TabIndex        =   5
      Top             =   2700
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
      MICON           =   "FRMKARDEX.frx":0038
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
   Begin LVbuttons.LaVolpeButton CB 
      Height          =   465
      Index           =   1
      Left            =   5835
      TabIndex        =   6
      Top             =   2700
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "STOCK"
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
      MICON           =   "FRMKARDEX.frx":0054
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
      TabIndex        =   8
      Top             =   8280
      Width           =   12390
      _ExtentX        =   21855
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10663
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10663
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   3030
      X2              =   12390
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ACCIONES"
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
      Left            =   6135
      TabIndex        =   7
      Top             =   2160
      Width           =   1125
   End
   Begin VB.Shape S 
      BackStyle       =   1  'Opaque
      Height          =   825
      Left            =   3780
      Top             =   2505
      Width           =   6000
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMIENTOS DE ALMACEN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   0
      Left            =   3075
      TabIndex        =   3
      Top             =   240
      Width           =   4575
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
      Left            =   3420
      TabIndex        =   1
      Top             =   900
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   3660
      Left            =   -30
      Picture         =   "FRMKARDEX.frx":0070
      Top             =   -90
      Width           =   2835
   End
End
Attribute VB_Name = "FRMKARDEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LaVolpeButton3_Click()

End Sub

Private Sub CB_Click(Index As Integer)
Select Case Index
Case 2
    Unload Me
Case 0
    FRMBUSCADOR.Show vbModal
Case 1
    FRMSTOCK.Show vbModal
End Select
End Sub

Private Sub CB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
S.BackColor = vbYellow
End Sub

Private Sub REALIZAR_Click()
ssql = "SELECT * FROM TIPO WHERE TIPO='" & DOCUMENTO.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    FRMMOVIMIENTO.TIPO.Tag = tbl!CODTIPO
    FRMMOVIMIENTO.TIPO.Text = DOCUMENTO.Text
    FRMMOVIMIENTO.Show vbModal
Else
    MsgBox "EL TIPO DE DOCUMENTO QUE USTED DESEA REALIZAR NO EXISTE": Exit Sub
End If
End Sub

Private Sub Form_Load()
ssql = "SELECT * FROM TIPO WHERE T<3 ORDER BY TIPO;"
Set tbl = conn.Execute(ssql)
DOCUMENTO.Clear
Do Until tbl.EOF
    DOCUMENTO.AddItem tbl!TIPO
    tbl.MoveNext
Loop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
S.BackColor = vbWhite
End Sub

