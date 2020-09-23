VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRMBUSCADOR 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "BUSCADOR DE DOCUMENTOS"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   1545
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
   ScaleHeight     =   9135
   ScaleWidth      =   15240
   Begin VB.ComboBox PROD 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6720
      TabIndex        =   16
      Top             =   720
      Width           =   8475
   End
   Begin VB.TextBox NDOCUMENTO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4440
      TabIndex        =   4
      Top             =   735
      Width           =   1620
   End
   Begin VB.ComboBox TIPO 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5370
      TabIndex        =   0
      Top             =   75
      Width           =   4185
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   8865
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
   Begin MSComCtl2.DTPicker DESDE 
      Height          =   315
      Left            =   5370
      TabIndex        =   6
      Top             =   390
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   8454143
      Format          =   64749569
      CurrentDate     =   38798
   End
   Begin MSComCtl2.DTPicker HASTA 
      Height          =   315
      Left            =   7860
      TabIndex        =   8
      Top             =   390
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   8454143
      Format          =   64749569
      CurrentDate     =   38798
   End
   Begin MSFlexGridLib.MSFlexGrid F 
      Height          =   6915
      Left            =   3045
      TabIndex        =   9
      Top             =   1905
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   12197
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
      MouseIcon       =   "FRMBUSCADOR.frx":0000
   End
   Begin LVbuttons.LaVolpeButton BUSCAR 
      Height          =   465
      Left            =   3030
      TabIndex        =   10
      Top             =   1275
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "BUSCAR"
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
      MICON           =   "FRMBUSCADOR.frx":031A
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   465
      Left            =   4110
      TabIndex        =   11
      Top             =   1275
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "ABRIR"
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
      MICON           =   "FRMBUSCADOR.frx":0336
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   465
      Left            =   5190
      TabIndex        =   12
      Top             =   1275
      Width           =   1065
      _ExtentX        =   1879
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
      MICON           =   "FRMBUSCADOR.frx":0352
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   465
      Left            =   6810
      TabIndex        =   13
      Top             =   1275
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "IMPRIMIR RESUMEN"
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
      MICON           =   "FRMBUSCADOR.frx":036E
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton4 
      Cancel          =   -1  'True
      Height          =   465
      Left            =   14115
      TabIndex        =   14
      Top             =   120
      Width           =   1065
      _ExtentX        =   1879
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
      MICON           =   "FRMBUSCADOR.frx":038A
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton5 
      Height          =   465
      Left            =   9225
      TabIndex        =   17
      Top             =   1275
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "IMPRIMIR DETALLE"
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
      MICON           =   "FRMBUSCADOR.frx":03A6
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
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PROD"
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
      Left            =   6105
      TabIndex        =   15
      Top             =   750
      Width           =   585
   End
   Begin VB.Line Line1 
      X1              =   9765
      X2              =   15420
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HASTA"
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
      Left            =   7110
      TabIndex        =   7
      Top             =   420
      Width           =   720
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DESDE"
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
      Left            =   4620
      TabIndex        =   5
      Top             =   420
      Width           =   705
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
      Index           =   0
      Left            =   3045
      TabIndex        =   3
      Top             =   765
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   5205
      Left            =   -15
      Picture         =   "FRMBUSCADOR.frx":03C2
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO DE DOCUMENTO"
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
      Left            =   3045
      TabIndex        =   1
      Top             =   105
      Width           =   2280
   End
End
Attribute VB_Name = "FRMBUSCADOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BUSCAR_Click()
ssql = "select * from tipo where tipo='" & TIPO.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    If NDOCUMENTO.Text = "" Then
        If PROD.Text = "" Then
            ssql = "select * from movimiento where codtipo=" & tbl!CODTIPO & " and fecha between " & FECHAS(DESDE, True) & " and " & FECHAS(HASTA, True) & " order by fecha"
        Else
            ssql = "SELECT MOVIMIENTO.CODMOVIMIENTO, MOVIMIENTO.FECHA, MOVIMIENTO.CODTIPO, MOVIMIENTO.CONCEPTO, MOVIMIENTO.TOTALS, MOVIMIENTO.IGV, MOVIMIENTO.TOTALG, DETMOVIMIENTO.CODPRODUCTO " & _
                    "FROM MOVIMIENTO INNER JOIN DETMOVIMIENTO ON MOVIMIENTO.CODMOVIMIENTO = DETMOVIMIENTO.CODMOVIMIENTO " & _
                    "WHERE (DETMOVIMIENTO.CODPRODUCTO='" & PROD.Tag & "' AND MOVIMIENTO.CODTIPO=" & tbl!CODTIPO & ") AND FECHA BETWEEN " & FECHAS(DESDE, True) & " AND " & FECHAS(HASTA, True) & " ORDER BY MOVIMIENTO.FECHA;"
        End If

    Else
        ssql = "select * from movimiento where CODMOVIMIENTO  ='" & NDOCUMENTO.Text & "';"
    End If
    
    Set tbl = conn.Execute(ssql)
    F.Rows = 1
    Do Until tbl.EOF
        F.AddItem FECHAS(tbl, False) & vbTab & tbl!CODMOVIMIENTO & vbTab & UCase(tbl!CONCEPTO) & vbTab & Format(tbl!TOTALS, "###,###,##0.00") & vbTab & Format(tbl!IGV, "###,###,##0.00") & vbTab & Format(tbl!TOTALG, "###,###,##0.00")
        tbl.MoveNext
    Loop
End If
End Sub

Private Sub Form_Load()
F.Rows = 1
F.FormatString = "FECHA|Nª DOCUMENTO|CONCEPTO|TOTAL|IGV|TOTAL GENERAL"
F.ColWidth(0) = 1500
F.ColWidth(1) = 1500
F.ColWidth(2) = 3500
F.ColWidth(3) = 2000
F.ColWidth(4) = 1000
F.ColWidth(5) = 2000
ssql = "SELECT * FROM TIPO ORDER BY TIPO;"
Set tbl = conn.Execute(ssql)
TIPO.Clear
Do Until tbl.EOF
    TIPO.AddItem UCase(tbl!TIPO)
    tbl.MoveNext
Loop
DESDE.Value = Date
HASTA.Value = Date
End Sub

Private Sub LaVolpeButton1_Click()
FRMMOVIMIENTO.NDOCUMENTO.Text = F.TextMatrix(F.Row, 1)
FRMMOVIMIENTO.NDOCUMENTO_KeyPress 13
FRMMOVIMIENTO.Show vbModal
End Sub

Private Sub LaVolpeButton2_Click()
A = MsgBox("ESTA SEGURO QUE DESEA ELIMINAR", vbYesNo)
If A = vbNo Then Exit Sub
ssql = "DELETE FROM MOVIMIENTO WHERE CODMOVIMIENTO='" & F.TextMatrix(F.Row, 1) & "';"
Set tbl = conn.Execute(ssql)
ssql = "DELETE FROM CAJA WHERE NDOCUMENTO='MM" & F.TextMatrix(F.Row, 1) & "'"
Set tbl = conn.Execute(ssql)
BUSCAR_Click
End Sub

Private Sub LaVolpeButton3_Click()
Dim DT As New ADODB.Recordset
DT.Fields.Append "C1", adChar, 255
DT.Fields.Append "C2", adChar, 255
DT.Fields.Append "C3", adChar, 255
DT.Fields.Append "C4", adChar, 255
DT.Fields.Append "C5", adChar, 255
DT.Fields.Append "C6", adChar, 255
DT.Open
For I = 0 To F.Rows - 1
    DT.AddNew
    DT!C1 = F.TextMatrix(I, 0)
    DT!C2 = F.TextMatrix(I, 1)
    DT!C3 = F.TextMatrix(I, 2)
    DT!C4 = F.TextMatrix(I, 3)
    DT!C5 = F.TextMatrix(I, 4)
    DT!C6 = F.TextMatrix(I, 5)
    DT.Update
Next
Set DTREPORTE.DataSource = DT
DTREPORTE.Sections("TIT").Controls("TITULO").Caption = "REPORTE  DE MOVIMIENTOS " & Chr(13) & TIPO.Text & Chr(13) & "DESDE: " & DESDE.Value & "-HASTA: " & HASTA.Value
DTREPORTE.LeftMargin = 200
DTREPORTE.RightMargin = 200
DTREPORTE.Show vbModal
End Sub

Private Sub LaVolpeButton4_Click()
Unload Me
End Sub

Private Sub LaVolpeButton5_Click()
Dim DT As New ADODB.Recordset
DT.Fields.Append "C1", adChar, 255
DT.Fields.Append "C2", adChar, 255
DT.Fields.Append "C3", adChar, 255
DT.Fields.Append "C4", adChar, 255
DT.Fields.Append "C5", adChar, 255
DT.Fields.Append "C6", adChar, 255
DT.Open
F.FormatString = "FECHA|Nª DOCUMENTO|CONCEPTO|PRECIO|CANTIDAD|TOTAL"
For I = 0 To F.Rows - 1
    If I > 0 Then
        ssql = "SELECT MOVIMIENTO.CODMOVIMIENTO, MOVIMIENTO.FECHA, PRODUCTO.PRODUCTO, DETMOVIMIENTO.CODPRODUCTO, DETMOVIMIENTO.CANTIDAD, DETMOVIMIENTO.PRECIO, MOVIMIENTO.CODTIPO, MOVIMIENTO.CONCEPTO, MOVIMIENTO.TOTALS, MOVIMIENTO.IGV, MOVIMIENTO.TOTALG " & _
                "FROM PRODUCTO INNER JOIN (MOVIMIENTO INNER JOIN DETMOVIMIENTO ON MOVIMIENTO.CODMOVIMIENTO = DETMOVIMIENTO.CODMOVIMIENTO) ON PRODUCTO.CODPRODUCTO = DETMOVIMIENTO.CODPRODUCTO " & _
                "WHERE MOVIMIENTO.CODMOVIMIENTO='" & F.TextMatrix(I, 1) & "';"
        Set tbl = conn.Execute(ssql)
        Do Until tbl.EOF
            DT.AddNew
            DT!C1 = FECHAS(tbl, False)
            DT!C2 = tbl!CODMOVIMIENTO
            DT!C3 = tbl!PRODUCTO
            DT!C4 = tbl!PRECIO
            DT!C5 = tbl!CANTIDAD
            DT!C6 = Format(Val(tbl!PRECIO) * Val(tbl!CANTIDAD), "###,##0.00")
            DT.Update
            tbl.MoveNext
        Loop
    Else
        DT.AddNew
        DT!C1 = F.TextMatrix(I, 0)
        DT!C2 = F.TextMatrix(I, 1)
        DT!C3 = F.TextMatrix(I, 2)
        DT!C4 = F.TextMatrix(I, 3)
        DT!C5 = F.TextMatrix(I, 4)
        DT!C6 = F.TextMatrix(I, 5)
        DT.Update
    End If
    
Next
Set DTREPORTE.DataSource = DT
DTREPORTE.Sections("TIT").Controls("TITULO").Caption = "REPORTE  DE MOVIMIENTOS " & Chr(13) & TIPO.Text & Chr(13) & "DESDE: " & DESDE.Value & "-HASTA: " & HASTA.Value
DTREPORTE.LeftMargin = 200
DTREPORTE.RightMargin = 200
DTREPORTE.Show vbModal

End Sub

Private Sub PROD_Click()
ssql = "SELECT * FROM PRODUCTO WHERE PRODUCTO='" & PROD.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    PROD.Tag = tbl!CODPRODUCTO
Else
    PROD.Tag = ""
End If
End Sub

Private Sub PROD_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 8
Case 13
Case Else
    ssql = "SELECT * FROM PRODUCTO WHERE PRODUCTO LIKE '%" & PROD.Text & "%' ORDER BY PRODUCTO;"
    Set tbl = conn.Execute(ssql)
    A = PROD.Text
    PROD.Clear
    Do Until tbl.EOF
        PROD.AddItem tbl!PRODUCTO
        tbl.MoveNext
    Loop
    
    res = SendMessageLong(PROD.hwnd, &H14F, True, 0)
    PROD.Text = A
    PROD.SelStart = Len(A)
    F.Refresh
    Me.Refresh
End Select

End Sub

Private Sub TIPO_Click()
ssql = "sELECT * FROM TIPO WHERE TIPO='" & TIPO.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    TIPO.Tag = tbl!CODTIPO
Else
    TIPO.Tag = ""
End If
End Sub
