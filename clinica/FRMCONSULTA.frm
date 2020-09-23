VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FRMCONSULTA 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGISTRO DE CONSULTAS"
   ClientHeight    =   9915
   ClientLeft      =   915
   ClientTop       =   810
   ClientWidth     =   13710
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
   ScaleHeight     =   9915
   ScaleWidth      =   13710
   Begin VB.ComboBox CODCITA 
      Height          =   315
      Left            =   5580
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   930
      Width           =   2265
   End
   Begin VB.TextBox CANCELADO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9585
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   6750
      Width           =   2220
   End
   Begin VB.TextBox CONCEPTO 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   3645
      Left            =   4395
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2640
      Width           =   7695
   End
   Begin VB.ComboBox PACIENTE 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5580
      TabIndex        =   6
      Top             =   585
      Width           =   6405
   End
   Begin VB.ComboBox DOCTOR 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5580
      TabIndex        =   3
      Top             =   1275
      Width           =   6405
   End
   Begin MSComCtl2.DTPicker FECHA 
      Height          =   315
      Left            =   5250
      TabIndex        =   1
      Top             =   165
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   8454143
      Format          =   23658497
      CurrentDate     =   38798
   End
   Begin MSMask.MaskEdBox PRECIO 
      Bindings        =   "FRMCONSULTA.frx":0000
      DataMember      =   "IMPORTE"
      Height          =   300
      Left            =   5925
      TabIndex        =   10
      Top             =   6750
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16761024
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox IMPORTECAN 
      Bindings        =   "FRMCONSULTA.frx":0012
      DataMember      =   "IMPORTE"
      Height          =   300
      Left            =   9600
      TabIndex        =   13
      Top             =   7335
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16761087
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox SALDO 
      Bindings        =   "FRMCONSULTA.frx":0024
      DataMember      =   "IMPORTE"
      Height          =   300
      Left            =   9600
      TabIndex        =   15
      Top             =   7770
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      BackColor       =   12632319
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
      TabIndex        =   21
      Top             =   9645
      Width           =   13710
      _ExtentX        =   24183
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11827
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11827
         EndProperty
      EndProperty
   End
   Begin VB.Label CITA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "ESPECIALIDAD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5580
      TabIndex        =   20
      Top             =   1635
      Width           =   6390
   End
   Begin VB.Label ESPECIALIDAD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "ESPECIALIDAD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5565
      TabIndex        =   18
      Top             =   1950
      Width           =   6390
   End
   Begin VB.Image Image1 
      Height          =   3675
      Left            =   870
      Picture         =   "FRMCONSULTA.frx":0036
      Stretch         =   -1  'True
      Top             =   2190
      Width           =   2355
   End
   Begin VB.Label CODIGO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7290
      TabIndex        =   16
      Top             =   60
      Width           =   4740
   End
   Begin VB.Line Line4 
      X1              =   11865
      X2              =   4410
      Y1              =   7170
      Y2              =   7170
   End
   Begin VB.Line Line3 
      X1              =   11850
      X2              =   6300
      Y1              =   7695
      Y2              =   7695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SALDO"
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
      Index           =   8
      Left            =   8835
      TabIndex        =   14
      Top             =   7815
      Width           =   720
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
      Index           =   7
      Left            =   7185
      TabIndex        =   12
      Top             =   7365
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CANCELADO"
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
      Left            =   8235
      TabIndex        =   11
      Top             =   6780
      Width           =   1320
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
      Index           =   5
      Left            =   5055
      TabIndex        =   9
      Top             =   6780
      Width           =   810
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO CITA"
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
      Left            =   4095
      TabIndex        =   8
      Top             =   945
      Width           =   1425
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3930
      X2              =   12180
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DATOS DE LA CONSULTA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   4395
      TabIndex        =   5
      Top             =   2385
      Width           =   2595
   End
   Begin VB.Line Line2 
      X1              =   5535
      X2              =   5535
      Y1              =   585
      Y2              =   1665
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
      Index           =   2
      Left            =   4410
      TabIndex        =   4
      Top             =   615
      Width           =   1065
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3900
      X2              =   12150
      Y1              =   510
      Y2              =   510
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
      Index           =   1
      Left            =   4590
      TabIndex        =   2
      Top             =   1305
      Width           =   885
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA"
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
      Left            =   4455
      TabIndex        =   0
      Top             =   195
      Width           =   735
   End
End
Attribute VB_Name = "FRMCONSULTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CANCELADO_Change()
SALDO.Text = Val(PRECIO.Text) - Val(CANCELADO.Text)
End Sub

Private Sub CANCELADO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If DOCTOR.Tag = "" Then MsgBox "INGRESE EL DOCTOR:": Exit Sub
    ssql = "SELECT * FROM DOCTOR WHERE DOCTOR='" & DOCTOR.Text & "';"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        DOCTOR.Tag = tbl!CODDOCTOR
    End If
    If PACIENTE.Tag = "" Then MsgBox "INGRESE EL PACIENTE": Exit Sub
    ssql = "SELECT * FROM CONSULTA ORDER BY DOCUMENTO DESC;"
    Set tbl = conn.Execute(ssql)
    COD = 1
    If tbl.EOF = False Then
        COD = Val(tbl!DOCUMENTO) + 1
    End If
    ssql = "SELECT * FROM CONSULTA WHERE CODCITA=" & CODCITA.Text & ";"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then MsgBox "NO SE PUEDE REGISTRAR DOS VECES LA MISMA CONSULTA": Exit Sub
    ssql = "INSERT INTO CONSULTA VALUES(" & COD & "," & FECHAS(FECHA, True) & "," & DOCTOR.Tag & "," & PACIENTE.Tag & ",'" & CONCEPTO.Text & "'," & Val(IMPORTECAN.Text) & "," & CODCITA.Text & ");"
    Set tbl = conn.Execute(ssql)
    COD1 = COD

    If Val(SALDO.Text) > 0 Then
        ssql = "SELECT * FROM CAJA ORDER BY CODCAJA DESC;"
        Set tbl = conn.Execute(ssql)
        COD = 0
        If tbl.EOF = False Then
            COD = Val(tbl!CODCAJA) + 1
        End If
        
        ssql = "INSERT INTO CAJA VALUES(" & COD & "," & FECHAS(FECHA, True) & ",'CS" & COD1 & "','INGRESO DE CREDITO POR CITA-CONSULTA-PACIENTE: " & PACIENTE.Text & "',2," & Val(SALDO.Text) & ");"
        Set tbl = conn.Execute(ssql)
        
        A3 = MsgBox("   CREDITO         . CONTADO", vbYesNo)
        If A3 = vbNo Then GoTo 67
        ssql = "SELECT * FROM TIPO WHERE CODTIPO=4;"
        Set tbl = conn.Execute(ssql)
        If tbl.EOF = True Then MsgBox "NO PUEDO BORRAR EL TIPO DE DOCUMENTO 4 QUE ES CREDITOS": Exit Sub
        FRMCAJA.TIPO.Tag = 4
        ssql = "SELECT * FROM TIPO WHERE CODTIPO=4;"
        Set tbl = conn.Execute(ssql)
        If tbl.EOF = True Then MsgBox "NO PUEDO BORRAR EL TIPO DE DOCUMENTO 4 QUE ES CREDITOS": Exit Sub
        FRMCAJA.TIPO.Text = tbl!TIPO
        FRMCAJA.IMPORTE.Text = SALDO.Text
        FRMCAJA.CONCEPTO.Text = "CREDITO " & FRMCAJA.TIPO.Text & ":" & PACIENTE.Tag & "-" & PACIENTE.Text
        'FRMCAJA.IMPORTE.SetFocus
        FRMCAJA.Show vbModal
67
    End If
    
    Unload Me
End If
End Sub

Private Sub CODCITA_Click()
ssql = "SELECT CITA.CODCITA, CITA.FECHA, CITA.HORA, CITA.CODPACIENTE, CITA.CODDOCTOR, CITA.CANCELADO, CITA.IMPORTE, CITA.PRECIO, CITA.CODESPECIALIDAD, ESPECIALIDAD.ESPECIALIDAD, DOCTOR.DOCTOR, PACIENTE.PACIENTE " & _
        "FROM PACIENTE INNER JOIN (DOCTOR INNER JOIN (CITA INNER JOIN ESPECIALIDAD ON CITA.CODESPECIALIDAD = ESPECIALIDAD.CODESPECIALIDAD) ON DOCTOR.CODDOCTOR = CITA.CODDOCTOR) ON PACIENTE.CODPACIENTE = CITA.CODPACIENTE " & _
        "WHERE CITA.CODCITA=" & CODCITA.Text & ";"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    PACIENTE.Tag = tbl!CODPACIENTE
    PACIENTE.Text = tbl!PACIENTE
    DOCTOR.Text = tbl!DOCTOR
    DOCTOR.Tag = tbl!DOCTOR
    PRECIO.Text = tbl!PRECIO
    ESPECIALIDAD.Caption = tbl!ESPECIALIDAD
    ESPECIALIDAD.Tag = tbl!CODESPECIALIDAD
        ssql = "SELECT * FROM CAJA WHERE NDOCUMENTO='C" & CODCITA.Text & "';"
        Set TB = conn.Execute(ssql)
        TOT = 0
        Do Until TB.EOF
            TOT = Val(TOT) + Val(TB!IMPORTE)
            TB.MoveNext
        Loop


        CANCELADO.Text = TOT
    ssql = "SELECT * FROM CONSULTA WHERE CODCITA=" & CODCITA.Text & ";"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        CITA.Caption = "ATENDIDO"
    Else
        CITA.Caption = "SIN ATENDER"
    End If
Else
    PACIENTE.Text = ""
    PACIENTE.Tag = ""
    DOCTOR.Text = ""
    DOCTOR.Tag = ""
    PRECIO.Text = ""
    
End If

End Sub

Private Sub CODCITA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DOCTOR.SetFocus
End Sub

Private Sub CONCEPTO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CANCELADO.SetFocus
End If
End Sub

Private Sub DOCTOR_Click()
ssql = "SELECT * FROM DOCTOR WHERE DOCTOR='" & DOCTOR.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    DOCTOR.Tag = tbl!CODDOCTOR
    ssql = "SELECT * FROM CITA WHERE FECHA=" & FECHAS(FECHA, True) & " AND CODDOCTOR=" & Val(DOCTOR.Tag) & " AND CODPACIENTE=" & Val(PACIENTE.Tag) & ";"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        CODCITA.Text = tbl!CODCITA
        PRECIO.Text = tbl!PRECIO
        CANCELADO.Text = tbl!IMPORTE
    End If
Else
    DOCTOR.Tag = ""
End If

End Sub

Private Sub DOCTOR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then PACIENTE.SetFocus
End Sub

Private Sub FECHA_Change()
PACIENTE_Click
End Sub

Private Sub FECHA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    DOCTOR.SetFocus
End If
End Sub

Private Sub Form_Load()
ssql = "SELECT * FROM CONSULTA ORDER BY DOCUMENTO DESC;"
Set tbl = conn.Execute(ssql)
    COD = 1
    If tbl.EOF = False Then
        COD = Val(tbl!DOCUMENTO) + 1
    End If
    CODIGO.Caption = "CODIGO CONSULTA: " & COD
FECHA.Value = Date
ssql = "sELECT * FROM DOCTOR ORDER BY DOCTOR;"
Set tbl = conn.Execute(ssql)
DOCTOR.Clear
Do Until tbl.EOF
    DOCTOR.AddItem tbl!DOCTOR
    tbl.MoveNext
Loop
ssql = "SELECT * FROM PACIENTE ORDER BY PACIENTE;"
Set tbl = conn.Execute(ssql)
PACIENTE.Clear
Do Until tbl.EOF
    PACIENTE.AddItem tbl!PACIENTE
    tbl.MoveNext
Loop
CONCEPTO.Text = ""
CODCITA.Text = ""
PRECIO.Text = ""
CANCELADO.Text = ""
IMPORTECAN.Text = ""
SALDO.Text = ""
End Sub

Private Sub IMPORTECAN_Change()
SALDO.Text = Val(PRECIO.Text) - Val(CANCELADO.Text) - Val(IMPORTECAN.Text)

End Sub

Private Sub IMPORTECAN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If DOCTOR.Tag = "" Then MsgBox "INGRESE EL DOCTOR:": Exit Sub
    If PACIENTE.Tag = "" Then MsgBox "INGRESE EL PACIENTE": Exit Sub
    ssql = "SELECT * FROM CONSULTA ORDER BY DOCUMENTO DESC;"
    Set tbl = conn.Execute(ssql)
    COD = 1
    If tbl.EOF = False Then
        COD = Val(tbl!DOCUMENTO) + 1
    End If
    ssql = "SELECT * FROM CONSULTA WHERE CODPACIENTE=" & PACIENTE.Tag & " AND CODDOCTOR=" & DOCTOR.Tag & " AND FECHA=" & FECHAS(FECHA, True) & ";"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then MsgBox "NO PUEDEN EXISTIR DOS CONSULTAS DEL MISMO PACIENTE Y DOCTOR EN UN MISMO DIA": Exit Sub
    ssql = "INSERT INTO CONSULTA VALUES(" & COD & "," & FECHAS(FECHA, True) & "," & DOCTOR.Tag & "," & PACIENTE.Tag & ",'" & CONCEPTO.Text & "'," & Val(IMPORTECAN.Text) & ");"
    Set tbl = conn.Execute(ssql)
   
    COD1 = COD
    If Val(IMPORTECAN.Text) > 0 Then
        ssql = "SELECT * FROM CAJA ORDER BY CODCAJA DESC;"
        Set tbl = conn.Execute(ssql)
        COD = 0
        If tbl.EOF = False Then
            COD = Val(tbl!CODCAJA) + 1
        End If
        ssql = "INSERT INTO CAJA VALUES(" & COD & "," & FECHAS(FECHA, True) & "," & COD1 & ",'INGRESO DE CANCELACION POR CITA-CONSULTA-PACIENTE: " & PACIENTE.Text & "',2," & Val(IMPORTECAN.Text) & ");"
        Set tbl = conn.Execute(ssql)
    End If
    If Val(SALDO.Text) > 0 Then
        FRMCAJA.NDOCUMENTO.Text = COD1
        ssql = "SELECT * FROM TIPO WHERE CODTIPO=4;"
        Set tbl = conn.Execute(ssql)
        If tbl.EOF = True Then MsgBox "NO PUEDO BORRAR EL TIPO DE DOCUMENTO 4 QUE ES CREDITOS": Exit Sub
        FRMCAJA.TIPO.Tag = 4
        FRMCAJA.TIPO.Text = tbl!TIPO
        FRMCAJA.IMPORTE.Text = SALDO.Text
        FRMCAJA.CONCEPTO.Text = "CREDITO " & FRMCAJA.TIPO.Text & ":" & PACIENTE.Tag & "-" & PACIENTE.Text
        'FRMCAJA.IMPORTE.SetFocus
        FRMCAJA.Show vbModal
        
    End If
    Unload Me
End If
    
End Sub

Private Sub PACIENTE_Click()
Dim TB As New ADODB.Recordset
ssql = "SELECT * FROM PACIENTE WHERE PACIENTE='" & PACIENTE.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    PACIENTE.Tag = tbl!CODPACIENTE
    ssql = "SELECT * FROM CITA WHERE CODPACIENTE=" & PACIENTE.Tag & " AND FECHA=" & FECHAS(FECHA, True) & ";"
    Set tbl = conn.Execute(ssql)
    CODCITA.Clear
    Do Until tbl.EOF
        CODCITA.AddItem tbl!CODCITA
        tbl.MoveNext
    Loop
Else
CODCITA.Clear

End If
End Sub

Private Sub paciente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CODCITA.SetFocus
End Sub

