VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FRMCAJA 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "REGISTRO DE CAJA"
   ClientHeight    =   9345
   ClientLeft      =   1440
   ClientTop       =   1020
   ClientWidth     =   11910
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
   ScaleHeight     =   9345
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.TextBox TOTALCAJA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   525
      Left            =   12165
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   9390
      Width           =   2865
   End
   Begin VB.TextBox CONCEPTO 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2205
      MaxLength       =   255
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   1575
      Width           =   10110
   End
   Begin VB.TextBox NDOCUMENTO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2190
      MaxLength       =   255
      TabIndex        =   4
      Top             =   1005
      Width           =   3885
   End
   Begin VB.ComboBox TIPO 
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
      Height          =   465
      Left            =   2190
      TabIndex        =   0
      Top             =   255
      Width           =   9615
   End
   Begin MSComCtl2.MonthView FECHA 
      Height          =   2310
      Left            =   12060
      TabIndex        =   1
      Top             =   3015
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      StartOfWeek     =   23724033
      CurrentDate     =   38798
   End
   Begin MSMask.MaskEdBox IMPORTE 
      Bindings        =   "FRMCAJA.frx":0000
      DataMember      =   "IMPORTE"
      Height          =   420
      Left            =   8580
      TabIndex        =   7
      Top             =   2580
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   741
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid F 
      Height          =   6900
      Left            =   1965
      TabIndex        =   9
      Top             =   3045
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   12171
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
   End
   Begin LVbuttons.LaVolpeButton ELIMINAR 
      Height          =   465
      Left            =   8055
      TabIndex        =   10
      Top             =   10005
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
      MICON           =   "FRMCAJA.frx":0012
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
      Left            =   9975
      TabIndex        =   11
      Top             =   10005
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
      MICON           =   "FRMCAJA.frx":002E
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
      Left            =   6135
      TabIndex        =   12
      Top             =   10005
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
      MICON           =   "FRMCAJA.frx":004A
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
      TabIndex        =   15
      Top             =   9075
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10239
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10239
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL CAJA S/."
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
      Index           =   3
      Left            =   12270
      TabIndex        =   14
      Top             =   9030
      Width           =   2580
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTE S/."
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
      Left            =   7080
      TabIndex        =   8
      Top             =   2640
      Width           =   1365
   End
   Begin VB.Line Line2 
      X1              =   150
      X2              =   8550
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION"
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
      Left            =   165
      TabIndex        =   5
      Top             =   1635
      Width           =   1485
   End
   Begin VB.Line Line1 
      X1              =   165
      X2              =   11835
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. DOCUMENTO"
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
      Left            =   165
      TabIndex        =   3
      Top             =   1035
      Width           =   1875
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO DOCUMENTO"
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
      Left            =   210
      TabIndex        =   2
      Top             =   285
      Width           =   1935
   End
   Begin VB.Image IM 
      Height          =   9000
      Left            =   1080
      Picture         =   "FRMCAJA.frx":0066
      Top             =   9540
      Width           =   12000
   End
End
Attribute VB_Name = "FRMCAJA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

End Sub

Private Sub CONCEPTO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then IMPORTE.SetFocus
End Sub

Private Sub ELIMINAR_Click()
ssql = "DELETE FROM CAJA WHERE CODCAJA=" & F.TextMatrix(F.Row, 0) & ";"
Set tbl = conn.Execute(ssql)
Form_Load
End Sub

Private Sub FECHA_Click()

F.Rows = 0
F.Cols = 4
F.ColWidth(0) = 1
F.ColWidth(1) = 1500
F.ColWidth(2) = 6000
F.ColWidth(3) = 2000
ssql = "SELECT * FROM TIPO ORDER BY D,TIPO;"
Set tbl = conn.Execute(ssql)
Dim TB As New ADODB.Recordset
Do Until tbl.EOF
    ssql = "SELECT CAJA.CODCAJA, CAJA.FECHA, CAJA.NDOCUMENTO, CAJA.CONCEPTO, CAJA.CODTIPO, CAJA.IMPORTE, TIPO.TIPO, TIPO.D " & _
            "FROM CAJA INNER JOIN TIPO ON CAJA.CODTIPO = TIPO.CODTIPO " & _
            "WHERE CAJA.FECHA=" & FECHAS(FECHA, True) & " AND CAJA.CODTIPO=" & tbl!CODTIPO & " ORDER BY TIPO.D,TIPO.TIPO,NDOCUMENTO DESC;"
    Set TB = conn.Execute(ssql)
    If TB.EOF = False Then
        F.AddItem ""
        If Val(TB!D) = 0 Then
            F.AddItem vbTab & vbTab & UCase(tbl!TIPO)
            F.Row = Val(F.Rows) - 1
            For I = 0 To F.Cols - 1
                F.Col = I
                F.CellBackColor = &HFFC0C0
            Next
            F.AddItem vbTab & "NDOCUMENTO" & vbTab & "CONCEPTO" & vbTab & "IMPORTE"
            F.Row = Val(F.Rows) - 1
            For I = 0 To F.Cols - 1
                F.Col = I
                F.CellBackColor = &HFFC0C0
            Next
        Else
            F.AddItem vbTab & vbTab & UCase(tbl!TIPO) & "(-)"
            F.Row = Val(F.Rows) - 1
            For I = 0 To F.Cols - 1
                F.Col = I
                F.CellBackColor = 12632319
            Next
            F.AddItem vbTab & "NDOCUMENTO" & vbTab & "CONCEPTO" & vbTab & "IMPORTE"
            F.Row = Val(F.Rows) - 1
            For I = 0 To F.Cols - 1
                F.Col = I
                F.CellBackColor = 12632319
            Next
        End If
        
        
    End If
    Do Until TB.EOF
        If Val(TB!D) = 0 Then
            F.AddItem TB!CODCAJA & vbTab & TB!NDOCUMENTO & vbTab & UCase(TB!CONCEPTO) & vbTab & Format(TB!IMPORTE, "###,###,##0.00")
            TOT = Val(TOT) + Val(TB!IMPORTE)
        Else
            F.AddItem TB!CODCAJA & vbTab & TB!NDOCUMENTO & vbTab & UCase(TB!CONCEPTO) & vbTab & Format(Val(TB!IMPORTE) * -1, "###,###,##0.00")
            TOT = Val(TOT) - Val(TB!IMPORTE)

            
        End If
        TB.MoveNext
    Loop
    If Val(TOT) > 0 Then
        TOT1 = TOT1 + Val(TOT)
        F.AddItem vbTab & vbTab & "TOTAL" & vbTab & Format(TOT, "###,###,##0.00")
        F.Row = Val(F.Rows) - 1
        For I = 0 To F.Cols - 1
            F.Col = I
            F.CellBackColor = &HFFC0C0
        Next
        TOT = 0
    ElseIf Val(TOT) < 0 Then
        TOT1 = TOT1 + Val(TOT)
        F.AddItem vbTab & vbTab & "TOTAL" & vbTab & Format(TOT, "###,###,##0.00")
        F.Row = Val(F.Rows) - 1
        For I = 0 To F.Cols - 1
            F.Col = I
            F.CellBackColor = 12632319
        Next
        TOT = 0
    End If
    tbl.MoveNext
Loop
TOTALCAJA.Text = Format(TOT1, "###,###,##0.00")
    F.AddItem ""
    F.AddItem ""
    F.AddItem vbTab & vbTab & "TOTAL NETO" & vbTab & Format(TOT1, "###,###,##0.00")
    F.Row = Val(F.Rows) - 1
    For I = 0 To F.Cols - 1
        F.Col = I
        F.CellBackColor = &HFFC0C0
    Next
    
    TOT1 = 0

End Sub

Private Sub FECHA_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
FECHA_Click
End Sub

Private Sub Form_Load()
ssql = "SELECT * FROM TIPO WHERE D<2 ORDER BY TIPO;"
Set tbl = conn.Execute(ssql)
TIPO.Clear
Do Until tbl.EOF
    TIPO.AddItem tbl!TIPO
    tbl.MoveNext
Loop
FECHA.Value = Date
NDOCUMENTO.Text = ""
CONCEPTO.Text = ""
F.Rows = 0
IMPORTE.Text = 0
FECHA_Click
End Sub

Private Sub Form_Resize()
IM.Top = 0
IM.Left = 0
End Sub

Private Sub IMPORTE_GotFocus()
IMPORTE.SelStart = 0
IMPORTE.SelLength = Len(IMPORTE.Text)
End Sub

Private Sub IMPORTE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    A = MsgBox("ESTA SEGURO QUE DESEA GRABAR", vbYesNo)
If A = vbNo Then Exit Sub
If TIPO.Tag = "" Then MsgBox "INGRESE EL TIPO DE DOCUMENTO": Exit Sub
If NDOCUMENTO.Text = "" Then MsgBox "INGRESE EL NUMERO DE DOCUMENTO": Exit Sub
If CONCEPTO.Text = "" Then MsgBox "INGRESE EL CONCEPTO": Exit Sub
If IMPORTE.Text = "" Then MsgBox "INGRESE EL IMPORTE ": Exit Sub
ssql = "SELECT * FROM CAJA ORDER BY CODCAJA DESC;"
Set tbl = conn.Execute(ssql)
COD = 0
If tbl.EOF = False Then
    COD = Val(tbl!CODCAJA) + 1
End If
'MsgBox CONCEPTO.Text
ssql = "INSERT INTO CAJA VALUES(" & COD & "," & FECHAS(FECHA, True) & ",'" & NDOCUMENTO.Text & "','" & CONCEPTO.Text & "'," & TIPO.Tag & "," & IMPORTE.Text & ");"
'MsgBox ssql
ssql = "INSERT INTO CAJA VALUES(" & COD & "," & FECHAS(FECHA, True) & ",'" & NDOCUMENTO.Text & "','" & CONCEPTO.Text & "'," & TIPO.Tag & "," & IMPORTE.Text & ");"

Set tbl = conn.Execute(ssql)
Unload Me
End If
End Sub

Private Sub IMPRIMIR_Click()
Dim DT As New ADODB.Recordset
DT.Fields.Append "C1", adChar, 255
DT.Fields.Append "C2", adChar, 255
DT.Fields.Append "C3", adChar, 255
DT.Open
For I = 0 To F.Rows - 1
    DT.AddNew
    DT!C1 = F.TextMatrix(I, 1)
    DT!C2 = F.TextMatrix(I, 2)
    DT!C3 = F.TextMatrix(I, 3)
    DT.Update
Next
Set DTCAJA.DataSource = DT
DTCAJA.Sections("TIT").Controls("TITULO").Caption = "REPORTE DE CAJA " & Chr(13) & "FECHA: " & FECHA.Value
DTCAJA.LeftMargin = 200
DTCAJA.RightMargin = 200
DTCAJA.Show vbModal
End Sub

Private Sub NDOCUMENTO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CONCEPTO.SetFocus
End Sub

Private Sub SALIR_Click()
Unload Me
End Sub

Private Sub TIPO_Click()
ssql = "sELECT * FROM TIPO WHERE TIPO='" & TIPO.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    TIPO.Tag = tbl!CODTIPO
    If Val(tbl!CODTIPO) = 5 Or Val(tbl!CODTIPO) = 4 Then
        FRMICREDITO.Show vbModal
        NDOCUMENTO.SetFocus
    End If
Else
    TIPO.Tag = ""
End If
End Sub

Private Sub TIPO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then NDOCUMENTO.SetFocus
End Sub
