VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FRMCITA 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGISTRO DE CITAS"
   ClientHeight    =   10755
   ClientLeft      =   105
   ClientTop       =   360
   ClientWidth     =   15225
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
   ScaleHeight     =   10755
   ScaleWidth      =   15225
   Begin VB.CheckBox CANCELAR 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "CANCELAR AHORA"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1275
      TabIndex        =   18
      Top             =   5280
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox PACIENTE 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   6240
      TabIndex        =   17
      Text            =   "DOCTOR"
      Top             =   2970
      Width           =   7590
   End
   Begin VB.ComboBox DOCTOR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   5145
      TabIndex        =   13
      Text            =   "DOCTOR"
      Top             =   2295
      Width           =   9030
   End
   Begin MSFlexGridLib.MSFlexGrid F 
      Height          =   4710
      Left            =   4275
      TabIndex        =   10
      Top             =   5670
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8308
      _Version        =   393216
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin MSComCtl2.MonthView FECHA 
      Height          =   2310
      Left            =   750
      TabIndex        =   9
      Top             =   2370
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      StartOfWeek     =   23592961
      CurrentDate     =   38798
   End
   Begin VB.ComboBox ESPECIALIDAD 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1755
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1185
      Width           =   8790
   End
   Begin MSMask.MaskEdBox HORA 
      Height          =   300
      Left            =   6240
      TabIndex        =   15
      Top             =   3465
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox IMPORTE 
      Bindings        =   "FRMCITA.frx":0000
      DataMember      =   "IMPORTE"
      Height          =   405
      Left            =   8700
      TabIndex        =   20
      Top             =   4290
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin LVbuttons.LaVolpeButton ELIMINAR 
      Height          =   465
      Left            =   10980
      TabIndex        =   21
      Top             =   5175
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
      MICON           =   "FRMCITA.frx":0012
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
      Left            =   12900
      TabIndex        =   22
      Top             =   5175
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
      MICON           =   "FRMCITA.frx":002E
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
      Height          =   420
      Left            =   8700
      TabIndex        =   23
      Top             =   3765
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   741
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin LVbuttons.LaVolpeButton NUEVO 
      Height          =   465
      Left            =   13890
      TabIndex        =   25
      Top             =   2880
      Width           =   930
      _ExtentX        =   1640
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
      MICON           =   "FRMCITA.frx":004A
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
   Begin MSMask.MaskEdBox FECHA2 
      Height          =   300
      Left            =   12825
      TabIndex        =   27
      Top             =   1380
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   29
      Top             =   10485
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13150
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13150
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA DEL ACTUAL"
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
      Left            =   10680
      TabIndex        =   28
      Top             =   1410
      Width           =   2100
   End
   Begin VB.Label FE 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DOCTOR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   45
      TabIndex        =   26
      Top             =   105
      Width           =   1680
   End
   Begin VB.Shape B 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2400
      Left            =   705
      Top             =   2325
      Width           =   3180
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO DE CONSULTA"
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
      Left            =   6285
      TabIndex        =   24
      Top             =   3900
      Width           =   2355
   End
   Begin VB.Shape Shape1 
      Height          =   1905
      Left            =   4920
      Top             =   2910
      Width           =   8955
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTE A CANCELAR"
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
      Left            =   6270
      TabIndex        =   19
      Top             =   4365
      Width           =   2370
   End
   Begin VB.Line Line2 
      X1              =   5145
      X2              =   14190
      Y1              =   2730
      Y2              =   2730
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
      Index           =   4
      Left            =   5130
      TabIndex        =   16
      Top             =   3060
      Width           =   1065
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HORA CITA"
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
      Left            =   5040
      TabIndex        =   14
      Top             =   3480
      Width           =   1185
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DOCTOR"
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
      Left            =   4170
      TabIndex        =   12
      Top             =   2340
      Width           =   885
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RELACION DE CITAS"
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
      Left            =   4260
      TabIndex        =   11
      Top             =   5370
      Width           =   2145
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DOMINGO"
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   7
      Left            =   12330
      TabIndex        =   8
      Top             =   1890
      Width           =   1890
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SABADO"
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   6
      Left            =   10425
      TabIndex        =   7
      Top             =   1890
      Width           =   1890
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VIERNES"
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   5
      Left            =   8520
      TabIndex        =   6
      Top             =   1890
      Width           =   1890
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUEVES"
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   4
      Left            =   6615
      TabIndex        =   5
      Top             =   1890
      Width           =   1890
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MIERCOLES"
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   3
      Left            =   4710
      TabIndex        =   4
      Top             =   1890
      Width           =   1890
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MARTES"
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   2
      Left            =   2805
      TabIndex        =   3
      Top             =   1890
      Width           =   1890
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LUNES"
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   900
      TabIndex        =   2
      Top             =   1890
      Width           =   1890
   End
   Begin VB.Line Line1 
      X1              =   900
      X2              =   14865
      Y1              =   1770
      Y2              =   1770
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
      Left            =   120
      TabIndex        =   0
      Top             =   1305
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   3225
      Left            =   210
      Picture         =   "FRMCITA.frx":0066
      Top             =   7350
      Width           =   2700
   End
End
Attribute VB_Name = "FRMCITA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()

End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub CANCELAR_Click()
If CANCELAR.Value = 1 Then
    IMPORTE.Enabled = True
Else
    IMPORTE.Enabled = False
    IMPORTE.Text = 0
End If
End Sub

Private Sub CANCELAR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IMPORTE.Enabled = True Then
        IMPORTE.SetFocus
    Else
        IMPORTE_KeyPress 13
    End If
End If
End Sub

Private Sub DOCTOR_GotFocus()
On Error Resume Next
ssql = "SELECT ESPECIALIDAD.CODESPECIALIDAD, ESPECIALIDAD.ESPECIALIDAD, ESPECIALIDAD.PRECIO, DETDOCTOR.CODDOCTOR, DOCTOR.DOCTOR, DOCTOR.TELEFONO1, DOCTOR.TELEFONO2, DOCTOR.DIRECCION, DOCTOR.DIA, DOCTOR.HORAINI, DOCTOR.HORAFIN " & _
            "FROM ESPECIALIDAD INNER JOIN (DOCTOR INNER JOIN DETDOCTOR ON DOCTOR.CODDOCTOR = DETDOCTOR.CODDOCTOR) ON ESPECIALIDAD.CODESPECIALIDAD = DETDOCTOR.CODESPECIALIDAD " & _
            "WHERE ESPECIALIDAD.CODESPECIALIDAD=" & ESPECIALIDAD.Tag & ";"
            
Set tbl = conn.Execute(ssql)
DOCTOR.Clear
DIAD = UCase(Format(FECHA.Value, "dddd"))
Debug.Print DIAD
Select Case DIAD
Case "LUNES"
    D = 1
Case "MARTES"
    D = 2
Case "MIÉRCOLES"
    D = 3
Case "JUEVES"
    D = 4
Case "VIERNES"
    D = 5
Case "SABADO"
    D = 6
Case "DOMINGO"
    D = 7
End Select
DOCTOR.Clear
Do Until tbl.EOF
    If InStr(1, tbl!DIA, D) Then
        DOCTOR.AddItem UCase(tbl!DOCTOR) & "-----------" & "HOR: " & tbl!HORAINI & " - " & tbl!HORAFIN
        DOCTOR.ItemData(DOCTOR.NewIndex) = tbl!CODDOCTOR
    End If
    tbl.MoveNext
Loop
res = SendMessageLong(DOCTOR.hwnd, &H14F, True, 0)

End Sub

Private Sub DOCTOR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then PACIENTE.SetFocus
End Sub

Private Sub ELIMINAR_Click()
A = MsgBox("DESDE ELIMINAR", vbYesNo)
If A = vbNo Then Exit Sub
ssql = "DELETE FROM CITA WHERE CODCITA=" & F.TextMatrix(F.Row, 0) & ";"
Set tbl = conn.Execute(ssql)
'ssql = "SELECT * FROM CAJA WHERE FECHA=" & FECHAS(FECHA, True) & " AND  CONCEPTO LIKE '%CITA PACIENTE: " & tbl!PACIENTE & "%' ;"

ssql = "DELETE FROM CAJA WHERE NDOCUMENTO='C" & F.TextMatrix(F.Row, 0) & "';"
Set tbl = conn.Execute(ssql)
Form_Load
FECHA.SetFocus
F.Rows = 1
End Sub

Private Sub ESPECIALIDAD_Click()
ssql = "SELECT * FROM ESPECIALIDAD WHERE ESPECIALIDAD='" & ESPECIALIDAD.Text & "';"
Set tbl = conn.Execute(ssql)
PRECIO.Text = ""
If tbl.EOF = False Then
    ESPECIALIDAD.Tag = tbl!CODESPECIALIDAD
    PRECIO.Text = tbl!PRECIO
    ssql = "SELECT ESPECIALIDAD.CODESPECIALIDAD, ESPECIALIDAD.ESPECIALIDAD, ESPECIALIDAD.PRECIO, DETDOCTOR.CODDOCTOR, DOCTOR.DOCTOR, DOCTOR.TELEFONO1, DOCTOR.TELEFONO2, DOCTOR.DIRECCION, DOCTOR.DIA, DOCTOR.HORAINI, DOCTOR.HORAFIN " & _
            "FROM ESPECIALIDAD INNER JOIN (DOCTOR INNER JOIN DETDOCTOR ON DOCTOR.CODDOCTOR = DETDOCTOR.CODDOCTOR) ON ESPECIALIDAD.CODESPECIALIDAD = DETDOCTOR.CODESPECIALIDAD " & _
            "WHERE ESPECIALIDAD.CODESPECIALIDAD=" & ESPECIALIDAD.Tag & ";"
    Set tbl = conn.Execute(ssql)
    For I = L.LBound To L.UBound
        L(I).BackColor = vbWhite
    Next
    Do Until tbl.EOF
        For I = 1 To Len(tbl!DIA)
        NUM = Val(Mid(tbl!DIA, I, 1))
            L(NUM).BackColor = vbYellow
        Next
        tbl.MoveNext
    Loop

End If
End Sub

Private Sub ESPECIALIDAD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FECHA.SetFocus
End If
End Sub

Private Sub FECHA_Click()
ssql = "SELECT CITA.CODCITA, CITA.FECHA, CITA.HORA, CITA.CODPACIENTE, CITA.CODDOCTOR, CITA.CANCELADO, CITA.IMPORTE, DOCTOR.DOCTOR, PACIENTE.PACIENTE " & _
        "FROM PACIENTE INNER JOIN (DOCTOR INNER JOIN CITA ON DOCTOR.CODDOCTOR = CITA.CODDOCTOR) ON PACIENTE.CODPACIENTE = CITA.CODPACIENTE " & _
        "WHERE CITA.FECHA=" & FECHAS(FECHA, True) & " ORDER BY CITA.HORA;"
Set tbl = conn.Execute(ssql)
F.FormatString = "CODCITA|HORA|PACIENTE|DOCTOR"
F.ColWidth(0) = 1
F.ColWidth(1) = 1000
F.ColWidth(2) = 4500
F.ColWidth(3) = 4500
F.Rows = 1
Do Until tbl.EOF
    F.AddItem tbl!CODCITA & vbTab & tbl!HORA & vbTab & UCase(tbl!PACIENTE & vbTab & tbl!DOCTOR)
    tbl.MoveNext
Loop
FE.Caption = "FECHA DE CITA: " & Format(FECHA.Value, "dddd dd - mmmm - yyyy")
End Sub

Private Sub FECHA_GotFocus()
B.BackColor = vbYellow
End Sub

Private Sub FECHA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DOCTOR.SetFocus
End If

End Sub

Private Sub FECHA_LostFocus()
B.BackColor = vbWhite
End Sub

Private Sub FECHA_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
FECHA_Click
End Sub

Private Sub Form_Load()
DOCTOR.Text = ""
FECHA2.Text = Date
PACIENTE.Tag = ""
HORA.Text = ""
IMPORTE.Text = ""
CANCELAR.Value = 1
ssql = "SELECT * FROM PACIENTE ORDER BY PACIENTE;"
Set tbl = conn.Execute(ssql)
PACIENTE.Clear
Do Until tbl.EOF
    PACIENTE.AddItem tbl!PACIENTE
    tbl.MoveNext
Loop
ssql = "SELECT * FROM ESPECIALIDAD ORDER BY ESPECIALIDAD;"
Set tbl = conn.Execute(ssql)
ESPECIALIDAD.Clear
Do Until tbl.EOF
    ESPECIALIDAD.AddItem tbl!ESPECIALIDAD
    tbl.MoveNext
Loop
FECHA.Value = Date
FECHA_Click
End Sub

Private Sub HORA_GotFocus()
DOCTOR.BackColor = vbYellow

End Sub

Private Sub HORA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then PRECIO.SetFocus
End Sub

Private Sub HORA_LostFocus()
DOCTOR.BackColor = vbWhite
End Sub

Private Sub IMPORTE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ssql = "SELECT * FROM PACIENTE WHERE PACIENTE='" & PACIENTE.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then PACIENTE.Tag = tbl!CODPACIENTE
If ESPECIALIDAD.Tag = "" Then MsgBox "TIENE QUE INGRESAR LA ESPECIALIDAD": Exit Sub
If PRECIO.Text = "" Then MsgBox "INGRESE EL PRECIO": Exit Sub
If PACIENTE.Text = "" Then MsgBox "INGRESE EL PACIENTE": Exit Sub
If PACIENTE.Tag = "" Then MsgBox "INGRESE EL PACIENTE": Exit Sub
If HORA.Text = "" Then MsgBox "INGRESE LA HORA DE LA CITA": Exit Sub
ssql = "SELECT * FROM CITA ORDER BY CODCITA DESC;"
Set tbl = conn.Execute(ssql)
COD = 0
If tbl.EOF = False Then
   COD = Val(tbl!CODCITA) + 1
End If
F12 = FECHAS(FECHA, True)
'ssql = "SELECT * FROM DOCTOR WHERE DOCTOR='" & DOCTOR.Text & "';"
'Set tbl = conn.Execute(ssql)
'If tbl.EOF = True Then Exit Sub
'DOCTOR.Tag = tbl!CODDOCTOR
If DOCTOR.ListIndex < 0 Then MsgBox "INGRESE EL DOCTOR": Exit Sub
ssql = "INSERT INTO CITA VALUES(" & COD & "," & F12 & ",'" & HORA.Text & "'," & PACIENTE.Tag & "," & DOCTOR.ItemData(DOCTOR.ListIndex) & "," & Val(CANCELAR.Value) & "," & Val(IMPORTE.Text) & "," & Val(PRECIO.Text) & "," & ESPECIALIDAD.Tag & ");"
Set tbl = conn.Execute(ssql)
COD1 = COD
If Val(IMPORTE.Text) > 0 Then
    ssql = "SELECT * FROM CAJA ORDER BY CODCAJA DESC;"
    Set tbl = conn.Execute(ssql)
    COD = 0
    If tbl.EOF = False Then
        COD = Val(tbl!CODCAJA) + 1
    End If
    ssql = "INSERT INTO CAJA VALUES(" & COD & "," & FECHAS(FECHA2, True) & ",'C" & COD1 & "','INGRESO DE ADELANTO EN CITA PACIENTE: " & PACIENTE.Text & "',2," & Val(IMPORTE.Text) & ");"
    Set tbl = conn.Execute(ssql)
    
End If
MsgBox "PACIENTE REGISTRADO"
Unload Me
DOCTOR.Text = ""
FECHA_Click
'ESPECIALIDAD.SetFocus
End If
End Sub

Private Sub NUEVO_Click()
FRMPACIENTE.Show vbModal
A = PACIENTE.Text
B1 = PACIENTE.Tag
ssql = "SELECT * FROM PACIENTE ORDER BY PACIENTE;"
Set tbl = conn.Execute(ssql)
PACIENTE.Clear
Do Until tbl.EOF
    PACIENTE.AddItem tbl!PACIENTE
    tbl.MoveNext
Loop
PACIENTE.Text = A
PACIENTE.Tag = B1
PACIENTE.SetFocus


End Sub

Private Sub PACIENTE_GotFocus()
res = SendMessageLong(PACIENTE.hwnd, &H14F, True, 0)
End Sub

Private Sub paciente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ssql = "SELECT * FROM PACIENTE WHERE PACIENTE='" & PACIENTE.Text & "';"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = True Then
        NUEVO_Click
    Else
        PACIENTE.Tag = tbl!CODPACIENTE
        HORA.SetFocus
    End If
End If
    
End Sub

Private Sub PRECIO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    IMPORTE.SetFocus
End If
End Sub

Private Sub SALIR_Click()
Unload Me
End Sub
