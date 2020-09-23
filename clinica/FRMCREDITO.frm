VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRMCREDITO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "REGISTRO DE CREDITOS"
   ClientHeight    =   10335
   ClientLeft      =   510
   ClientTop       =   420
   ClientWidth     =   14640
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
   ScaleHeight     =   10335
   ScaleWidth      =   14640
   Begin VB.ComboBox PACIENTE 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3270
      TabIndex        =   0
      Top             =   195
      Width           =   6405
   End
   Begin LVbuttons.LaVolpeButton IMPRIMIR 
      Height          =   465
      Left            =   10785
      TabIndex        =   2
      Top             =   9585
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
      MICON           =   "FRMCREDITO.frx":0000
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
      Height          =   8400
      Left            =   3495
      TabIndex        =   4
      Top             =   1125
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   14817
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   12648447
      BackColorFixed  =   16761024
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      MousePointer    =   99
      MouseIcon       =   "FRMCREDITO.frx":001C
   End
   Begin LVbuttons.LaVolpeButton SALIR 
      Cancel          =   -1  'True
      Height          =   465
      Left            =   12705
      TabIndex        =   5
      Top             =   9585
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
      MICON           =   "FRMCREDITO.frx":0336
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
      TabIndex        =   6
      Top             =   10065
      Width           =   14640
      _ExtentX        =   25823
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12647
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12647
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   3300
      X2              =   14550
      Y1              =   9540
      Y2              =   9540
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RELACION DE CREDITOS Y PAGOS"
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
      Left            =   3270
      TabIndex        =   3
      Top             =   735
      Width           =   3585
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
      Left            =   2145
      TabIndex        =   1
      Top             =   210
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   5250
      Left            =   -300
      Picture         =   "FRMCREDITO.frx":0352
      Top             =   4815
      Width           =   3480
   End
End
Attribute VB_Name = "FRMCREDITO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DOCTOR_Change()

End Sub

Private Sub Form_Load()
F.Rows = 1
F.FormatString = "cocaja|FECHA|NDOCUMENTO|CREDITO|PAGO"
F.ColWidth(0) = 1
F.ColWidth(1) = 1500
F.ColWidth(2) = 3500
F.ColWidth(3) = 2000
F.ColWidth(4) = 2000
ssql = "SELECT * FROM PACIENTE ORDER BY PACIENTE;"
Set tbl = conn.Execute(ssql)
PACIENTE.Clear
Do Until tbl.EOF
    PACIENTE.AddItem tbl!PACIENTE
    tbl.MoveNext
Loop

End Sub

Private Sub IMPRIMIR_Click()
Dim DT As New ADODB.Recordset
DT.Fields.Append "FECHA", adChar, 255
DT.Fields.Append "NDOCUMENTO", adChar, 255
DT.Fields.Append "CREDITO", adChar, 255
DT.Fields.Append "PAGO", adChar, 255
DT.Open
For I = 0 To F.Rows - 1
    DT.AddNew
    DT!FECHA = F.TextMatrix(I, 1)
    DT!NDOCUMENTO = F.TextMatrix(I, 2)
    DT!credito = F.TextMatrix(I, 3)
    DT!pago = F.TextMatrix(I, 4)
    DT.Update
Next
Set DTCREDITO.DataSource = DT
DTCREDITO.LeftMargin = 300
DTCREDITO.RightMargin = 300
DTCREDITO.Sections("TIT").Controls("TITULO").Caption = PACIENTE.Text
DTCREDITO.Show vbModal
End Sub

Private Sub PACIENTE_Click()
ssql = "SELECT * FROM PACIENTE WHERE PACIENTE='" & PACIENTE.Text & "';"
Set tbl = conn.Execute(ssql)


If tbl.EOF = True Then Exit Sub
PACIENTE.Tag = tbl!CODPACIENTE
ssql = "SELECT * FROM CAJA WHERE (CODTIPO=5 OR CODTIPO=4) AND CONcEPTO LIKE '%" & PACIENTE.Tag & "%' order by FECHA,CODTIPO ;"
Set tbl = conn.Execute(ssql)
credito = 0
pago = 0
F.Rows = 1

Do Until tbl.EOF
    If Val(tbl!CODTIPO) = 4 Then
        F.AddItem tbl!CODCAJA & vbTab & FECHAS(tbl, False) & vbTab & "CREDITO Nª:" & tbl!NDOCUMENTO & vbTab & Format(tbl!IMPORTE, "###,###,##0.00") & vbTab & "0.00"
        credito = Val(credito) + Val(tbl!IMPORTE)
    ElseIf Val(tbl!CODTIPO) = 5 Then
        F.AddItem tbl!CODCAJA & vbTab & FECHAS(tbl, False) & vbTab & "PAGO Nª: " & tbl!NDOCUMENTO & vbTab & "0.00" & vbTab & Format(tbl!IMPORTE, "###,###,##0.00")
        pago = Val(pago) + Val(tbl!IMPORTE)
        F.AddItem ""
        F.AddItem vbTab & FECHAS(tbl, False) & vbTab & "SALDO HASTA LA FECHA" & vbTab & Format(Val(credito) - Val(pago), "###,###,##0.00")

        F.Row = F.Rows - 1
        For I = 0 To F.Cols - 1
            F.Col = I
            F.CellBackColor = &HFFC0C0
        Next
        F.AddItem ""
    End If
    tbl.MoveNext
Loop
F.AddItem ""
F.AddItem vbTab & vbTab & "SALDO GENERAL " & vbTab & Format(Val(credito) - Val(pago), "###,###,##0.00")
        F.Row = F.Rows - 1
        For I = 0 To F.Cols - 1
            F.Col = I
            F.CellBackColor = &HFFC0C0
        Next

End Sub

Private Sub SALIR_Click()
Unload Me
End Sub
