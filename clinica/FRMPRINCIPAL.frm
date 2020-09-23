VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FRMPRINCIPAL 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "FRMPRINCIPAL"
   ClientHeight    =   8625
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   10035
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMPRINCIPAL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   10035
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox RT 
      Height          =   300
      Left            =   9585
      TabIndex        =   15
      Top             =   2790
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   529
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"FRMPRINCIPAL.frx":030A
   End
   Begin MSComCtl2.MonthView FECHA 
      Height          =   2310
      Left            =   12150
      TabIndex        =   9
      Top             =   3480
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      StartOfWeek     =   59179009
      CurrentDate     =   38800
   End
   Begin MSComctlLib.ImageList IM 
      Left            =   13320
      Top             =   1065
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8355
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8573
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8573
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton CB 
      Height          =   510
      Index           =   4
      Left            =   11595
      TabIndex        =   1
      Top             =   7455
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   900
      BTYPE           =   3
      TX              =   "CITA"
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
      COLTYPE         =   2
      BCOL            =   16761024
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   99
      MICON           =   "FRMPRINCIPAL.frx":0386
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
      Height          =   450
      Index           =   5
      Left            =   13440
      TabIndex        =   2
      Top             =   9750
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   794
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
      COLTYPE         =   2
      BCOL            =   16761024
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   99
      MICON           =   "FRMPRINCIPAL.frx":06A0
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
      Height          =   405
      Index           =   6
      Left            =   105
      TabIndex        =   3
      Top             =   4020
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "KARDEX"
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
      COLTYPE         =   2
      BCOL            =   16761024
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   99
      MICON           =   "FRMPRINCIPAL.frx":09BA
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
      Height          =   405
      Index           =   8
      Left            =   90
      TabIndex        =   4
      Top             =   5985
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "CREDITO"
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
      COLTYPE         =   2
      BCOL            =   16761024
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   99
      MICON           =   "FRMPRINCIPAL.frx":0CD4
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
      Height          =   405
      Index           =   9
      Left            =   105
      TabIndex        =   5
      Top             =   7965
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "CAJA"
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
      COLTYPE         =   2
      BCOL            =   16761024
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   99
      MICON           =   "FRMPRINCIPAL.frx":0FEE
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
      Height          =   510
      Index           =   12
      Left            =   13590
      TabIndex        =   6
      Top             =   7455
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   900
      BTYPE           =   3
      TX              =   "HIST. CLINICA"
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
      COLTYPE         =   2
      BCOL            =   16761024
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   99
      MICON           =   "FRMPRINCIPAL.frx":1308
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
      Index           =   16
      Left            =   13455
      TabIndex        =   7
      Top             =   1995
      Width           =   1845
      _ExtentX        =   3254
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
      MPTR            =   99
      MICON           =   "FRMPRINCIPAL.frx":1622
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
   Begin MSFlexGridLib.MSFlexGrid F 
      Height          =   7080
      Left            =   2370
      TabIndex        =   8
      Top             =   3390
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   12488
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   12648447
      BackColorFixed  =   16761024
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      MousePointer    =   99
      MouseIcon       =   "FRMPRINCIPAL.frx":193C
   End
   Begin LVbuttons.LaVolpeButton CB 
      Height          =   450
      Index           =   7
      Left            =   11595
      TabIndex        =   10
      Top             =   9750
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   794
      BTYPE           =   3
      TX              =   "FICHA DE CONSULTA"
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
      COLTYPE         =   2
      BCOL            =   16761024
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   99
      MICON           =   "FRMPRINCIPAL.frx":1C56
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
      Height          =   2190
      Index           =   11
      Left            =   11340
      TabIndex        =   11
      Top             =   3390
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   3863
      BTYPE           =   3
      TX              =   "PAGO DE DEUDA"
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
      COLTYPE         =   2
      BCOL            =   16761024
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   99
      MICON           =   "FRMPRINCIPAL.frx":1F70
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   1
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin MSMask.MaskEdBox FECHA2 
      Height          =   540
      Left            =   12150
      TabIndex        =   12
      Top             =   2925
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   953
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   16761024
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin LVbuttons.LaVolpeButton CB 
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   9870
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "CONFIGURACION"
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
      COLTYPE         =   2
      BCOL            =   16761024
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   99
      MICON           =   "FRMPRINCIPAL.frx":228A
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA ACTUAL"
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
      Height          =   375
      Left            =   12390
      TabIndex        =   16
      Top             =   2505
      Width           =   2625
   End
   Begin VB.Image Image6 
      Height          =   1815
      Index           =   3
      Left            =   105
      Picture         =   "FRMPRINCIPAL.frx":25A4
      Stretch         =   -1  'True
      Top             =   2565
      Width           =   2145
   End
   Begin VB.Image Image6 
      Height          =   1815
      Index           =   2
      Left            =   105
      Picture         =   "FRMPRINCIPAL.frx":5586
      Stretch         =   -1  'True
      Top             =   6555
      Width           =   2145
   End
   Begin VB.Image Image6 
      Height          =   1815
      Index           =   1
      Left            =   90
      Picture         =   "FRMPRINCIPAL.frx":D647
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   2145
   End
   Begin VB.Image Image6 
      Height          =   1815
      Index           =   0
      Left            =   105
      Picture         =   "FRMPRINCIPAL.frx":FB68
      Stretch         =   -1  'True
      Top             =   8460
      Width           =   2145
   End
   Begin VB.Image Image5 
      Height          =   1770
      Left            =   13485
      Picture         =   "FRMPRINCIPAL.frx":12873
      Stretch         =   -1  'True
      Top             =   8370
      Width           =   1515
   End
   Begin VB.Image Image3 
      Height          =   1770
      Left            =   11610
      Picture         =   "FRMPRINCIPAL.frx":19940
      Stretch         =   -1  'True
      Top             =   8355
      Width           =   1515
   End
   Begin VB.Image Image2 
      Height          =   1770
      Left            =   13605
      Picture         =   "FRMPRINCIPAL.frx":1DECA
      Stretch         =   -1  'True
      Top             =   6135
      Width           =   1515
   End
   Begin VB.Label EMPRESA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   435
      TabIndex        =   14
      Top             =   465
      Width           =   1770
   End
   Begin VB.Image Image1 
      Height          =   1770
      Left            =   11415
      Picture         =   "FRMPRINCIPAL.frx":2A3B5
      Stretch         =   -1  'True
      Top             =   6135
      Width           =   1755
   End
   Begin VB.Image Image4 
      Height          =   330
      Index           =   137
      Left            =   3750
      Picture         =   "FRMPRINCIPAL.frx":2D7C3
      Top             =   10965
      Width           =   750
   End
   Begin VB.Image Image4 
      Height          =   330
      Index           =   136
      Left            =   2970
      Picture         =   "FRMPRINCIPAL.frx":2DC8A
      Top             =   10965
      Width           =   750
   End
   Begin VB.Image Image4 
      Height          =   330
      Index           =   135
      Left            =   2190
      Picture         =   "FRMPRINCIPAL.frx":2E151
      Top             =   10965
      Width           =   750
   End
   Begin VB.Image Image4 
      Height          =   330
      Index           =   134
      Left            =   1410
      Picture         =   "FRMPRINCIPAL.frx":2E618
      Top             =   10965
      Width           =   750
   End
   Begin VB.Image Image4 
      Height          =   330
      Index           =   133
      Left            =   630
      Picture         =   "FRMPRINCIPAL.frx":2EADF
      Top             =   10965
      Width           =   750
   End
   Begin VB.Image Image4 
      Height          =   330
      Index           =   132
      Left            =   -150
      Picture         =   "FRMPRINCIPAL.frx":2EFA6
      Top             =   10965
      Width           =   750
   End
   Begin VB.Image IMA 
      Height          =   11520
      Left            =   -15
      Picture         =   "FRMPRINCIPAL.frx":2F46D
      Top             =   -105
      Width           =   15360
   End
End
Attribute VB_Name = "FRMPRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB_Click(Index As Integer)
Select Case Index
Case 0
    FRMCONFIGURACION.Show vbModal
Case 1
    FRMDOCTOR.Show vbModal
Case 16
    End
Case 10
    FRMESPECIALIDAD.Show vbModal
Case 2
    FRMPRODUCTO.Show vbModal
Case 3
    frmdocumento.Show vbModal
Case 4
    FRMCITA.Show vbModal
Case 5
    FRMCONSULTA.Show vbModal
Case 12
    FRMHISTORIAL.Show vbModal
Case 9
    FRMCAJA.Show vbModal
Case 6
    FRMKARDEX.Show vbModal
Case 8
    FRMCREDITO.Show vbModal
Case 7
    ssql = "SELECT CITA.CODCITA, CITA.FECHA, CITA.HORA, CITA.CODPACIENTE, CITA.CODDOCTOR, CITA.CANCELADO, CITA.IMPORTE, CITA.PRECIO, DOCTOR.DOCTOR, PACIENTE.PACIENTE, ESPECIALIDAD.ESPECIALIDAD, DETDOCTOR.CODESPECIALIDAD " & _
            "FROM ESPECIALIDAD INNER JOIN ((PACIENTE INNER JOIN (DOCTOR INNER JOIN CITA ON DOCTOR.CODDOCTOR = CITA.CODDOCTOR) ON PACIENTE.CODPACIENTE = CITA.CODPACIENTE) INNER JOIN DETDOCTOR ON DOCTOR.CODDOCTOR = DETDOCTOR.CODDOCTOR) ON ESPECIALIDAD.CODESPECIALIDAD = DETDOCTOR.CODESPECIALIDAD " & _
            "WHERE CODCITA=" & F.TextMatrix(F.Row, 0) & ";"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
    Dim DT As New ADODB.Recordset
    DT.Fields.Append "ESPECIALIDAD", adChar, 255
    DT.Open
    DT.AddNew
    DT!ESPECIALIDAD = "  "
    DT.Update
    Set DTFICHA.DataSource = DT
        DTFICHA.Sections("TITULO").Controls("FECHA").Caption = "FECHA DE CONSULTA " & FECHAS(tbl, False) & " " & tbl!HORA
        DTFICHA.Sections("TITULO").Controls("NUMERO").Caption = "HISTORIA CLINICA: " & tbl!CODPACIENTE
        DTFICHA.Sections("TITULO").Controls("PACIENTE").Caption = UCase("PACIENTE: " & tbl!PACIENTE)
        DTFICHA.Sections("TITULO").Controls("DOCTOR").Caption = UCase("DOCTOR: " & tbl!DOCTOR) & "                                        ESPECIALIDAD: " & F.TextMatrix(F.Row, 5)
        ssql = "SELECT * FROM PACIENTE WHERE PACIENTE='" & tbl!PACIENTE & "';"
        Set tbl = conn.Execute(ssql)
        If tbl.EOF = True Then Exit Sub
        DTFICHA.Sections("TITULO").Controls("DIRECCION").Caption = "DIRECCION: " & tbl!DIRECCION & "         TELEFONO: " & tbl!TELEFONO1
        DF = DateDiff("yyyy", tbl!FECHA, Date)
        DTFICHA.Sections("TITULO").Controls("EDAD").Caption = "EDAD DEL PACIENTE: " & DF & " AÑOS"
        DTFICHA.TopMargin = 200
        DTFICHA.BottomMargin = 200
        DTFICHA.Show vbModal
    End If
Case 11
If Val(F.TextMatrix(F.Row, 3)) > 0 Then
        
    A = InputBox("INGRESE EL MONTO QUE DESEA CANCELAR", "SISTEMA CLINICO", F.TextMatrix(F.Row, 3))
    If IsNumeric(A) = False Then Exit Sub
    ssql = "SELECT * FROM CAJA ORDER BY CODCAJA DESC;"
        Set tbl = conn.Execute(ssql)
        COD = 0
        If tbl.EOF = False Then
            COD = Val(tbl!CODCAJA) + 1
        End If
        ssql = "INSERT INTO CAJA VALUES(" & COD & "," & FECHAS(FECHA2, True) & ",'C" & F.TextMatrix(F.Row, 0) & "','INGRESO DE ADELANTO EN CITA PACIENTE: " & F.TextMatrix(F.Row, 2) & "',2," & A & ");"
        Set tbl = conn.Execute(ssql)
        FECHA_Click
End If
End Select
End Sub

Private Sub FECHA_Click()
 ssql = "SELECT CITA.CODESPECIALIDAD,CITA.PRECIO,CITA.CODCITA, CITA.FECHA, CITA.HORA, CITA.CODPACIENTE, CITA.CODDOCTOR, CITA.CANCELADO, CITA.IMPORTE, DOCTOR.DOCTOR, PACIENTE.PACIENTE " & _
        "FROM PACIENTE INNER JOIN (DOCTOR INNER JOIN CITA ON DOCTOR.CODDOCTOR = CITA.CODDOCTOR) ON PACIENTE.CODPACIENTE = CITA.CODPACIENTE " & _
        "WHERE CITA.FECHA=" & FECHAS(FECHA, True) & " ORDER BY CITA.HORA;"
Set tbl = conn.Execute(ssql)
F.FormatString = "CODCITA|HORA|PACIENTE|DEUDA|ATENDIDO|ESPECIALIDAD"
F.ColWidth(0) = 1
F.ColWidth(1) = 800
F.ColWidth(2) = 3500
F.ColWidth(3) = 1400
F.ColWidth(4) = 1000
F.ColWidth(5) = 2000
F.Rows = 1
Dim TB As New ADODB.Recordset
Do Until tbl.EOF
        ssql = "SELECT * FROM ESPECIALIDAD WHERE CODESPECIALIDAD=" & tbl!CODESPECIALIDAD & ";"
        Set TB = conn.Execute(ssql)
        If TB.EOF = False Then
            ESPE = TB!ESPECIALIDAD
        End If

        ssql = "SELECT * FROM CONSULTA WHERE CODCITA=" & tbl!CODCITA & ";"
        Set TB = conn.Execute(ssql)
        If TB.EOF = False Then
            ATE = "SI"
        Else
            ATE = "NO"
        End If
        ssql = "SELECT * FROM CAJA WHERE NDOCUMENTO='C" & tbl!CODCITA & "';"
        Set TB = conn.Execute(ssql)
        TOT = 0
        Do Until TB.EOF
            TOT = Val(TOT) + Val(TB!IMPORTE)
            TB.MoveNext
        Loop
        ssql = "SELECT * FROM CONSULTA WHERE CODCITA=" & tbl!CODCITA & ";"
        Set TB = conn.Execute(ssql)
        If TB.EOF = False Then
            ssql = "SELECT * FROM CAJA WHERE NDOCUMENTO='CS" & TB!DOCUMENTO & "' AND CONCEPTO LIKE '%CREDITO%';"
            Set TB = conn.Execute(ssql)
            If TB.EOF = False Then
                TOT = Val(TOT) + Val(TB!IMPORTE)
            End If
        End If
        F.AddItem tbl!CODCITA & vbTab & tbl!HORA & vbTab & UCase(tbl!PACIENTE) & vbTab & Val(Val(TOT) - Val(tbl!PRECIO)) * -1 & vbTab & ATE & vbTab & ESPE
    
        If Val(TOT) < Val(tbl!PRECIO) Then
            F.Row = Val(F.Rows) - 1
            For I = 0 To F.Cols - 1
                F.Col = I
                F.CellForeColor = vbRed
            Next
        Else
            F.Row = Val(F.Rows) - 1
            For I = 0 To F.Cols - 1
                F.Col = I
                F.CellForeColor = vbBlue
            Next
        End If
        
    tbl.MoveNext
Loop
End Sub

Private Sub FECHA_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
FECHA_Click
End Sub

Private Sub Form_Activate()
On Error Resume Next
RT.LoadFile App.Path & "\EMPRESA.TXT"
EMPRESA.Caption = UCase(RT.Text)
FECHA2.Text = Date
FECHA_Click
End Sub

Private Sub Form_Load()
FRMLOGON.Show vbModal

FECHA.Value = Date
FECHA_Click
End Sub

Private Sub Form_Resize()

IMA.Top = 0
IMA.Left = 0
IMA.Height = Me.Height
IMA.Width = Me.Width
End Sub



Private Sub Form_Unload(Cancel As Integer)
End
End Sub

