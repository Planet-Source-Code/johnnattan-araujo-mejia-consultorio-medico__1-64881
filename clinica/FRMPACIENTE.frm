VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRMPACIENTE 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGISTRO DE PACIENTES"
   ClientHeight    =   5310
   ClientLeft      =   1800
   ClientTop       =   2835
   ClientWidth     =   11010
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
   ScaleHeight     =   5310
   ScaleWidth      =   11010
   Begin MSComCtl2.DTPicker FECHA 
      Height          =   315
      Left            =   6270
      TabIndex        =   18
      Top             =   1530
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   556
      _Version        =   393216
      Format          =   23592961
      CurrentDate     =   38805
   End
   Begin LVbuttons.LaVolpeButton ELIMINAR 
      Height          =   465
      Left            =   6345
      TabIndex        =   14
      Top             =   3855
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
      MICON           =   "FRMPACIENTE.frx":0000
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
   Begin VB.TextBox REFERENCIA 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1740
      TabIndex        =   13
      Top             =   3465
      Width           =   6450
   End
   Begin VB.TextBox DIRECCION 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1740
      TabIndex        =   11
      Top             =   2985
      Width           =   6450
   End
   Begin VB.TextBox TELEFONO2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1740
      TabIndex        =   9
      Top             =   2505
      Width           =   3105
   End
   Begin VB.TextBox TELEFONO1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1740
      TabIndex        =   7
      Top             =   2025
      Width           =   3105
   End
   Begin VB.TextBox DNI 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1740
      TabIndex        =   5
      Top             =   1545
      Width           =   3105
   End
   Begin VB.ComboBox PACIENTE 
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1050
      Width           =   6450
   End
   Begin VB.TextBox CODPACIENTE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5265
      TabIndex        =   1
      Top             =   45
      Width           =   2895
   End
   Begin LVbuttons.LaVolpeButton ACEPTAR 
      Height          =   465
      Left            =   4425
      TabIndex        =   15
      Top             =   3855
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
      MICON           =   "FRMPACIENTE.frx":001C
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
      Left            =   6330
      TabIndex        =   16
      Top             =   4365
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
      MICON           =   "FRMPACIENTE.frx":0038
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
      TabIndex        =   19
      Top             =   5040
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9446
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9446
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA NAC."
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
      Index           =   7
      Left            =   4920
      TabIndex        =   17
      Top             =   1575
      Width           =   1320
   End
   Begin VB.Line Line2 
      X1              =   1560
      X2              =   1560
      Y1              =   1050
      Y2              =   3810
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "REFERENCIA"
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
      Left            =   75
      TabIndex        =   12
      Top             =   3495
      Width           =   1350
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION"
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
      Index           =   5
      Left            =   210
      TabIndex        =   10
      Top             =   3015
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TELEFONO 2"
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
      Index           =   4
      Left            =   105
      TabIndex        =   8
      Top             =   2535
      Width           =   1320
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TELEFONO 1"
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
      Left            =   105
      TabIndex        =   6
      Top             =   2070
      Width           =   1320
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DNI"
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
      Left            =   1035
      TabIndex        =   4
      Top             =   1605
      Width           =   390
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PACIENTE"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Line Line1 
      X1              =   2715
      X2              =   8235
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO HISTORIAL"
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
      Left            =   3090
      TabIndex        =   2
      Top             =   75
      Width           =   2100
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   7740
      Picture         =   "FRMPACIENTE.frx":0054
      Top             =   1950
      Width           =   3270
   End
End
Attribute VB_Name = "FRMPACIENTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ACEPTAR_Click()
If CODPACIENTE.Text = "" Then Exit Sub
If PACIENTE.Text = "" Then Exit Sub
ssql = "SELECT * FROM PACIENTE WHERE CODPACIENTE=" & CODPACIENTE.Text & ";"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    ssql = UCase("UPDATE PACIENTE SET FECHA='" & FECHA.Value & "', PACIENTE='" & PACIENTE.Text & "', DNI='" & DNI.Text & "',TELEFONO1='" & TELEFONO1.Text & "',TELEFONO2='" & TELEFONO2.Text & "',DIRECCION='" & DIRECCION.Text & "',REFERENCIA='" & REFERENCIA.Text & "' WHERE CODPACIENTE=" & CODPACIENTE.Text & ";")
Else
    ssql = UCase("INSERT INTO PACIENTE VALUES(" & Val(CODPACIENTE.Text) & ",'" & PACIENTE.Text & "','" & DNI.Text & "','" & TELEFONO1.Text & "','" & TELEFONO2.Text & "','" & DIRECCION.Text & "','" & REFERENCIA.Text & "','" & FECHA.Value & "');")
End If
'MsgBox ssql
Set tbl = conn.Execute(ssql)

FRMCITA.PACIENTE.Text = PACIENTE.Text
FRMCITA.PACIENTE.Tag = CODPACIENTE.Text
Unload Me
End Sub

Private Sub CODPACIENTE_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    ssql = "SELECT * FROM PACIENTE WHERE CODPACIENTE=" & Val(CODPACIENTE.Text) & ";"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        PACIENTE.Text = tbl!PACIENTE
        DNI.Text = tbl!DNI
        TELEFONO1.Text = tbl!TELEFONO1
        TELEFONO2.Text = tbl!TELEFONO2
        DIRECCION.Text = tbl!DIRECCION
        REFERENCIA.Text = tbl!REFERENCIA
        If IsNull(tbl!FECHA) = True Then
            FECHA.Value = Date
        Else
            FECHA.Value = Format(tbl!FECHA, "dd/MM/yyyy")
        End If
    Else
        FECHA.Value = Date
        PACIENTE.Text = ""
        DNI.Text = ""
        TELEFONO1.Text = ""
        TELEFONO2.Text = ""
        DIRECCION.Text = ""
        REFERENCIA.Text = ""
    End If
    PACIENTE.SetFocus
End If
End Sub

Private Sub DIRECCION_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then REFERENCIA.SetFocus
End Sub

Private Sub DNI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then FECHA.SetFocus
End Sub

Private Sub ELIMINAR_Click()
ssql = "DELETE FROM PACIENTE WHERE CODPACIENTE=" & Val(CODPACIENTE.Text) & ";"
Set tbl = conn.Execute(ssql)
Form_Load

End Sub

Private Sub FECHA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then TELEFONO1.SetFocus

End Sub

Private Sub PACIENTE_Click()
ssql = "SELECT * FROM PACIENTE WHERE PACIENTE='" & PACIENTE.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    CODPACIENTE.Text = tbl!CODPACIENTE
End If
CODPACIENTE_KeyPress 13

End Sub

Private Sub PACIENTE_GotFocus()
res = SendMessageLong(PACIENTE.hwnd, &H14F, True, 0)

End Sub

Private Sub Form_Load()
FECHA.Value = Date
CODPACIENTE.Text = ""
PACIENTE.Text = ""
DNI.Text = ""
TELEFONO1.Text = ""
TELEFONO2.Text = ""
DIRECCION.Text = ""
REFERENCIA.Text = ""
ssql = "SELECT * FROM PACIENTE ORDER BY PACIENTE;"
Set tbl = conn.Execute(ssql)
PACIENTE.Clear
Do Until tbl.EOF
    PACIENTE.AddItem tbl!PACIENTE
    tbl.MoveNext
Loop
ssql = "SELECT * FROM PACIENTE ORDER BY CODPACIENTE DESC;"
Set tbl = conn.Execute(ssql)
CODPACIENTE.Text = 0
If tbl.EOF = False Then CODPACIENTE.Text = Val(tbl!CODPACIENTE)
CODPACIENTE.Text = Val(CODPACIENTE.Text) + 1
End Sub

Private Sub paciente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DNI.SetFocus
End Sub

Private Sub REFERENCIA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ACEPTAR_Click
End Sub

Private Sub SALIR_Click()
Unload Me
End Sub

Private Sub TELEFONO1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TELEFONO2.SetFocus
End Sub

Private Sub TELEFONO2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DIRECCION.SetFocus
End Sub
