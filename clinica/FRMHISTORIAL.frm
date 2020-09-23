VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRMHISTORIAL 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HISTORIALCLINICO"
   ClientHeight    =   8235
   ClientLeft      =   15
   ClientTop       =   1440
   ClientWidth     =   15300
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
   ScaleHeight     =   8235
   ScaleWidth      =   15300
   Begin MSFlexGridLib.MSFlexGrid F 
      Height          =   6750
      Left            =   3975
      TabIndex        =   4
      Top             =   1215
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   11906
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   8454143
      BackColorFixed  =   16761024
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.TextBox CODPACIENTE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2295
      TabIndex        =   1
      Top             =   285
      Width           =   2895
   End
   Begin VB.ComboBox PACIENTE 
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2295
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   675
      Width           =   5610
   End
   Begin LVbuttons.LaVolpeButton ELIMINAR 
      Height          =   465
      Left            =   13395
      TabIndex        =   5
      Top             =   720
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
      MICON           =   "FRMHISTORIAL.frx":0000
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
   Begin LVbuttons.LaVolpeButton IMPRIMIR 
      Height          =   465
      Left            =   11460
      TabIndex        =   6
      Top             =   720
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "IMPRIMIR"
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
      MICON           =   "FRMHISTORIAL.frx":001C
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
      Top             =   7965
      Width           =   15300
      _ExtentX        =   26988
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13229
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13229
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   3975
      X2              =   3975
      Y1              =   1215
      Y2              =   8220
   End
   Begin VB.Line Line1 
      X1              =   15
      X2              =   15285
      Y1              =   1215
      Y2              =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO PACIENTE"
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
      Left            =   225
      TabIndex        =   3
      Top             =   285
      Width           =   1995
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
      Left            =   1170
      TabIndex        =   2
      Top             =   720
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   -210
      Picture         =   "FRMHISTORIAL.frx":0038
      Top             =   4005
      Width           =   4500
   End
End
Attribute VB_Name = "FRMHISTORIAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CODESPECIALIDAD_Change()

End Sub

Private Sub CODPACIENTE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
F.FormatString = "DOCUMENTO|FECHA|DOCTOR|DATOS CONSULTA"
F.Rows = 1
F.ColWidth(0) = 1
F.ColWidth(1) = 1500
F.ColWidth(2) = 2500
F.ColWidth(3) = 7000
    ssql = "SELECT * FROM PACIENTE WHERE CODPACIENTE=" & Val(CODPACIENTE.Text) & ";"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        PACIENTE.Text = tbl!PACIENTE
        ssql = "SELECT CONSULTA.DOCUMENTO, CONSULTA.FECHA, CONSULTA.CODDOCTOR, CONSULTA.CODPACIENTE, CONSULTA.CONCEPTO, CONSULTA.IMPORTE, DOCTOR.DOCTOR " & _
                "FROM CONSULTA INNER JOIN DOCTOR ON CONSULTA.CODDOCTOR = DOCTOR.CODDOCTOR " & _
                "WHERE CONSULTA.CODPACIENTE=" & Val(CODPACIENTE.Text) & " ORDER BY CONSULTA.FECHA;"
        Set tbl = conn.Execute(ssql)
        Do Until tbl.EOF
            F.AddItem tbl!DOCUMENTO & vbTab & FECHAS(tbl, False) & vbTab & tbl!DOCTOR & vbTab & UCase(tbl!CONCEPTO)
            tbl.MoveNext
        Loop
    Else
        PACIENTE.Text = ""
    End If
End If
End Sub

Private Sub ELIMINAR_Click()
A = MsgBox("ESTA SEGURO QUE DESEA ELMINAR", vbYesNo)
If A = vbNo Then Exit Sub

ssql = "DELETE FROM CONSULTA WHERE DOCUMENTO=" & F.TextMatrix(F.Row, 0) & ";"
Set tbl = conn.Execute(ssql)
ssql = "DELETE FROM CAJA WHERE NDOCUMENTO='CS" & F.TextMatrix(F.Row, 0) & "' AND CONCEPTO LIKE '%CONSULTA%' AND CODTIPO=2;"
Set tbl = conn.Execute(ssql)
Form_Load

F.Rows = 1

End Sub

Private Sub Form_Load()
F.FormatString = "DOCUMENTO|FECHA|DOCTOR|DATOS CONSULTA"
F.ColWidth(0) = 1
F.ColWidth(1) = 1500
F.ColWidth(2) = 2500
F.ColWidth(3) = 7000

F.Rows = 1
ssql = "SELECT * FROM PACIENTE ORDER BY PACIENTE;"
Set tbl = conn.Execute(ssql)
PACIENTE.Clear
Do Until tbl.EOF
    PACIENTE.AddItem UCase(tbl!PACIENTE)
    tbl.MoveNext
Loop

End Sub

Private Sub IMPRIMIR_Click()
Dim DT As New ADODB.Recordset
DT.Fields.Append "NDOCUMENTO", adChar, 255
DT.Fields.Append "FECHA", adChar, 255
DT.Fields.Append "DOCTOR", adChar, 255
DT.Fields.Append "CONCEPTO", adLongVarWChar, 5000
DT.Open
For I = 1 To F.Rows - 1
    DT.AddNew
    DT!NDOCUMENTO = F.TextMatrix(I, 0)
    DT!FECHA = F.TextMatrix(I, 1)
    DT!DOCTOR = F.TextMatrix(I, 2)
    DT!CONCEPTO = F.TextMatrix(I, 3)
    DT.Update
Next
Set DTHISTORIAL.DataSource = DT
DTHISTORIAL.TopMargin = 200
DTHISTORIAL.BottomMargin = 200
DTHISTORIAL.LeftMargin = 200
DTHISTORIAL.RightMargin = 200
DTHISTORIAL.Sections("TIT").Controls("TITULO").Caption = "CODHISTORIAL: " & CODPACIENTE.Text & Chr(13) & "PACIENTE: " & PACIENTE.Text
DTHISTORIAL.Show vbModal
End Sub

Private Sub PACIENTE_Click()
ssql = "sELECT * FROM PACIENTE WHERE PACIENTE='" & PACIENTE.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    CODPACIENTE.Text = tbl!CODPACIENTE
    CODPACIENTE_KeyPress 13
End If

    
End Sub
