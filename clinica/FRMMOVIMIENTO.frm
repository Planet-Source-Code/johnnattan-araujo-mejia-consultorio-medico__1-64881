VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FRMMOVIMIENTO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "REGISTRO DE MOVIMIENTOS, COMPRAS Y VENTAS"
   ClientHeight    =   10740
   ClientLeft      =   2175
   ClientTop       =   360
   ClientWidth     =   11835
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
   ScaleHeight     =   10740
   ScaleWidth      =   11835
   Begin VB.ComboBox CONCEPTO 
      Height          =   315
      Left            =   4440
      TabIndex        =   27
      Top             =   1770
      Width           =   7185
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   240
      Left            =   6555
      TabIndex        =   26
      Top             =   2640
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CheckBox paciente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "PACIENTES"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6885
      TabIndex        =   24
      Top             =   1410
      Width           =   1995
   End
   Begin VB.ListBox BUSQUEDA 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   2370
      Left            =   240
      TabIndex        =   9
      Top             =   2925
      Width           =   11400
   End
   Begin VB.TextBox PRODUCTO 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1260
      TabIndex        =   8
      Top             =   2250
      Width           =   7620
   End
   Begin VB.TextBox NDOCUMENTO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4440
      TabIndex        =   0
      Top             =   1005
      Width           =   2220
   End
   Begin VB.TextBox TIPO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4305
      TabIndex        =   2
      Text            =   "DFSDF"
      Top             =   15
      Width           =   7485
   End
   Begin MSMask.MaskEdBox FECHA 
      Bindings        =   "FRMMOVIMIENTO.frx":0000
      DataMember      =   "IMPORTE"
      Height          =   300
      Left            =   4440
      TabIndex        =   6
      Top             =   1410
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   16761024
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox CANTIDAD 
      Bindings        =   "FRMMOVIMIENTO.frx":0012
      DataMember      =   "IMPORTE"
      Height          =   330
      Left            =   3675
      TabIndex        =   11
      Top             =   5475
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   16761024
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
   Begin MSMask.MaskEdBox PRECIO 
      Bindings        =   "FRMMOVIMIENTO.frx":0024
      DataMember      =   "IMPORTE"
      Height          =   330
      Left            =   6795
      TabIndex        =   13
      Top             =   5460
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   16761024
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
   Begin MSMask.MaskEdBox TOTAL 
      Bindings        =   "FRMMOVIMIENTO.frx":0036
      DataMember      =   "IMPORTE"
      Height          =   330
      Left            =   9615
      TabIndex        =   15
      Top             =   5460
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   16761024
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
   Begin MSFlexGridLib.MSFlexGrid F 
      Height          =   3180
      Left            =   330
      TabIndex        =   17
      Top             =   6015
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   5609
      _Version        =   393216
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin MSMask.MaskEdBox TOTALS 
      Bindings        =   "FRMMOVIMIENTO.frx":0048
      DataMember      =   "IMPORTE"
      Height          =   300
      Left            =   9675
      TabIndex        =   18
      Top             =   9240
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   16761024
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox IGV 
      Bindings        =   "FRMMOVIMIENTO.frx":005A
      DataMember      =   "IMPORTE"
      Height          =   300
      Left            =   9675
      TabIndex        =   20
      Top             =   9585
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12632319
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TOTALG 
      Bindings        =   "FRMMOVIMIENTO.frx":006C
      DataMember      =   "IMPORTE"
      Height          =   300
      Left            =   9675
      TabIndex        =   22
      Top             =   10065
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   16761024
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   28
      Top             =   10470
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10160
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10160
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STOCK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9720
      TabIndex        =   29
      Top             =   2550
      Width           =   1605
   End
   Begin VB.Label STOCK 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Height          =   300
      Left            =   9720
      TabIndex        =   25
      Top             =   2265
      Width           =   1605
   End
   Begin VB.Line Line4 
      X1              =   11715
      X2              =   8190
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   11
      Left            =   8655
      TabIndex        =   23
      Top             =   10095
      Width           =   555
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IGV"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   10
      Left            =   8655
      TabIndex        =   21
      Top             =   9615
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   9
      Left            =   8655
      TabIndex        =   19
      Top             =   9270
      Width           =   555
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
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
      Left            =   8865
      TabIndex        =   16
      Top             =   5520
      Width           =   675
   End
   Begin VB.Line Line3 
      X1              =   11640
      X2              =   8460
      Y1              =   5850
      Y2              =   5850
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
      Index           =   7
      Left            =   5955
      TabIndex        =   14
      Top             =   5535
      Width           =   810
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
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
      Left            =   2520
      TabIndex        =   12
      Top             =   5535
      Width           =   1110
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RESULTADOS DE LA BUSQUEDA"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   10
      Top             =   2670
      Width           =   2775
   End
   Begin VB.Line Line2 
      X1              =   11685
      X2              =   2460
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTO"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   2325
      Width           =   990
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   3810
      TabIndex        =   5
      Top             =   1455
      Width           =   570
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CONCEPTO"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   3405
      TabIndex        =   4
      Top             =   1830
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº DOCUMENTO"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   2985
      TabIndex        =   3
      Top             =   1065
      Width           =   1395
   End
   Begin VB.Line Line1 
      X1              =   11835
      X2              =   2970
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "TIPO DE DOCUMENTO"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   2340
      TabIndex        =   1
      Top             =   150
      Width           =   1905
   End
   Begin VB.Shape Shape1 
      Height          =   2130
      Left            =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   -30
      Picture         =   "FRMMOVIMIENTO.frx":007E
      Top             =   -30
      Width           =   2325
   End
End
Attribute VB_Name = "FRMMOVIMIENTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BUSQUEDA_Click()
ssql = "SELECT * FROM PRODUCTO WHERE PRODUCTO='" & BUSQUEDA.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    PRECIO.Text = tbl!PRECIO
    BUSQUEDA.Tag = tbl!CODPRODUCTO
    'STOCK
    
    ssql = "sELECT * FROM PRODUCTO WHERE PRODUCTO='" & BUSQUEDA.Text & "' ORDER BY CODPRODUCTO;"
    Set tbl = conn.Execute(ssql)
    Dim TB As New ADODB.Recordset
    Do Until tbl.EOF
    DoEvents
        PB.Max = tbl.RecordCount
        PB.Value = tbl.Bookmark
        ssql = "SELECT DETMOVIMIENTO.CODMOVIMIENTO, DETMOVIMIENTO.CODPRODUCTO, DETMOVIMIENTO.CANTIDAD, DETMOVIMIENTO.PRECIO, MOVIMIENTO.CODTIPO, TIPO.T " & _
                "FROM TIPO INNER JOIN (MOVIMIENTO INNER JOIN DETMOVIMIENTO ON MOVIMIENTO.CODMOVIMIENTO = DETMOVIMIENTO.CODMOVIMIENTO) ON TIPO.CODTIPO = MOVIMIENTO.CODTIPO " & _
                "WHERE DETMOVIMIENTO.CODPRODUCTO='" & tbl!CODPRODUCTO & "';"
        Set TB = conn.Execute(ssql)
        STOCK1 = 0
        Do Until TB.EOF
            If Val(TB!T) = 1 Then
                STOCK1 = Val(STOCK1) + Val(TB!CANTIDAD)
            Else
                STOCK1 = Val(STOCK1) - Val(TB!CANTIDAD)
            End If
            TB.MoveNext
        Loop
        STOCK.Caption = Round(STOCK1, 2)
        tbl.MoveNext
    Loop
        PRODUCTO.Text = BUSQUEDA.Text
Else
    BUSQUEDA.Tag = ""
    PRECIO.Text = ""
End If

End Sub

Private Sub BUSQUEDA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    PRODUCTO.Text = BUSQUEDA.Text
    CANTIDAD.SetFocus
    PRODUCTO.Tag = BUSQUEDA.Tag
End If
End Sub

Private Sub CANTIDAD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    PRECIO.SetFocus
End If
End Sub

Private Sub CONCEPTO_Click()
ssql = "SELECT * FROM PACIENTE WHERE PACIENTE='" & CONCEPTO.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    CONCEPTO.Tag = tbl!CODPACIENTE
Else
    CONCEPTO.Tag = ""
End If
End Sub

Private Sub CONCEPTO_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 13
    If paciente.Value = 1 Then
        ssql = "select * from paciente where paciente LIKE '%" & CONCEPTO.Text & "%';"
        Set tbl = conn.Execute(ssql)
        If tbl.EOF = True Then MsgBox "NO EXISTE ESE PACIENTE": CONCEPTO.Text = "": Exit Sub
        PRODUCTO.SetFocus
    Else
        PRODUCTO.SetFocus
    End If
End Select
End Sub

Private Sub F_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    If F.Rows > 1 Then
        F.RemoveItem F.Row
    Else
        F.Rows = 1
    End If
End If
End Sub

Private Sub FECHA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then paciente.SetFocus
End Sub

Private Sub Form_Load()
F.Rows = 1
F.FormatString = "CODIGO|PRODUCTO|CANTIDAD|PRECIO|TOTAL"
F.ColWidth(0) = 1
F.ColWidth(1) = 6500
F.ColWidth(2) = 1500
F.ColWidth(3) = 1500
F.ColWidth(4) = 1500
FECHA.Text = Date

End Sub


Private Sub IGV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TOTALG.Text = Val(TOTALS.Text) + Val(IGV.Text)
    TOTALG.SetFocus
End If
End Sub

Public Sub NDOCUMENTO_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Dim TB As New ADODB.Recordset
    ssql = "SELECT * FROM MOVIMIENTO WHERE CODMOVIMIENTO='" & NDOCUMENTO.Text & "';"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        FECHA.Text = FECHAS(tbl, False)
        NDOCUMENTO.Text = tbl!CODMOVIMIENTO
        CONCEPTO.Text = tbl!CONCEPTO
        TOTALS.Text = tbl!TOTALS
        IGV.Text = tbl!IGV
        TOTALG.Text = tbl!TOTALG
        ssql = "SELECT * FROM TIPO WHERE CODTIPO=" & tbl!CODTIPO & ";"
        Set tbl = conn.Execute(ssql)
        If tbl.EOF = False Then
            TIPO.Text = tbl!TIPO
            TIPO.Tag = tbl!CODTIPO
        End If
        ssql = "SELECT DETMOVIMIENTO.CODMOVIMIENTO, DETMOVIMIENTO.CODPRODUCTO, DETMOVIMIENTO.CANTIDAD, DETMOVIMIENTO.PRECIO, PRODUCTO.PRODUCTO " & _
                "FROM PRODUCTO INNER JOIN DETMOVIMIENTO ON PRODUCTO.CODPRODUCTO = DETMOVIMIENTO.CODPRODUCTO " & _
                " WHERE DETMOVIMIENTO.CODMOVIMIENTO='" & NDOCUMENTO.Text & "' ORDER BY PRODUCTO.PRODUCTO;"
        Set tbl = conn.Execute(ssql)
        F.Rows = 1
        Do Until tbl.EOF
            F.AddItem tbl!CODPRODUCTO & vbTab & tbl!PRODUCTO & vbTab & Format(tbl!CANTIDAD, "###,###,##0.00") & vbTab & Format(tbl!PRECIO, "###,###,##0.00") & vbTab & Format(Val(tbl!CANTIDAD) * Val(tbl!PRECIO), "###,###,##0.00")
            tbl.MoveNext
        Loop
    End If
        

FECHA.SetFocus
End If
End Sub

Private Sub PACIENTE_Click()
If paciente.Value = 1 Then
        ssql = "SELECT * FROM PACIENTE ORDER BY PACIENTE;"
        Set tbl = conn.Execute(ssql)
        CONCEPTO.Clear
        Do Until tbl.EOF
            CONCEPTO.AddItem tbl!paciente
            tbl.MoveNext
        Loop
Else
    CONCEPTO.Clear
End If
End Sub

Private Sub paciente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CONCEPTO.SetFocus
End Sub

Private Sub PRECIO_Change()
TOTAL.Text = Val(PRECIO.Text) * Val(CANTIDAD.Text)
End Sub

Private Sub PRECIO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TOTAL.SetFocus
TOTAL.Text = Val(PRECIO.Text) * Val(CANTIDAD.Text)
End If
End Sub

Private Sub PRODUCTO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then BUSQUEDA.SetFocus
End Sub

Private Sub PRODUCTO_KeyUp(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
Select Case KeyCode


Case 8
Case Else
    ssql = "SELECT * FROM PRODUCTO WHERE PRODUCTO LIKE '%" & PRODUCTO.Text & "%' ORDER BY PRODUCTO"
    Set tbl = conn.Execute(ssql)
    BUSQUEDA.Clear
    Do Until tbl.EOF
        BUSQUEDA.AddItem UCase(tbl!PRODUCTO)
        
        tbl.MoveNext
    Loop

End Select
    
End Sub

Private Sub TOTAL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(CANTIDAD.Text) = False Then MsgBox "EL VALOR DE LA CANTIDAD NO ES NUMERICO": CANTIDAD.SetFocus: Exit Sub
    If IsNumeric(PRECIO.Text) = False Then MsgBox "EL VALOR DE EL PRECIO NO ES NUMERICO": PRECIO.SetFocus: Exit Sub
    If IsNumeric(TOTAL.Text) = False Then MsgBox "EL VALOR DE EL TOTAL NO ES NUMERICO": TOTAL.SetFocus: Exit Sub
    If PRODUCTO.Text = "" Then MsgBox "INGRESE EL PRODUCTO": Exit Sub
    If CANTIDAD.Text = "" Then MsgBox "INGRESE LA CANTIDAD": Exit Sub
    If PRECIO.Text = "" Then MsgBox "INGRESE EL PRECIO": Exit Sub
    If TOTAL.Text = "" Then MsgBox "INGRESE EL TOTAL": Exit Sub
    ssql = "SELECT * FROM TIPO WHERE TIPO='" & TIPO.Text & "';"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        If Val(tbl!T) = 2 Then
            If Val(CANTIDAD.Text) > Val(STOCK.Caption) Then MsgBox "NO SE PUEDE REGISTRAR MAS QUE EL STOCK QUE TIENE": Exit Sub
        End If
    End If
    ssql = "SELECT * FROM PRODUCTO WHERE PRODUCTO='" & PRODUCTO.Text & "';"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        PRODUCTO.Tag = tbl!CODPRODUCTO
    Else
        MsgBox "INGRESE BIEN EL PRODUCTO": Exit Sub
    End If
    
    For I = 1 To F.Rows - 1
    If UCase(F.TextMatrix(I, 1)) = UCase(PRODUCTO.Text) Then
        F.TextMatrix(I, 2) = CANTIDAD.Text
        F.TextMatrix(I, 3) = PRECIO.Text
        F.TextMatrix(I, 4) = TOTAL.Text
        GoTo 56
    End If
    Next
    F.AddItem PRODUCTO.Tag & vbTab & PRODUCTO.Text & vbTab & Format(CANTIDAD.Text, "0.00") & vbTab & Format(PRECIO.Text, "0.00") & vbTab & Format(TOTAL.Text, "0.00")
56
    PRODUCTO.Text = ""
    PRODUCTO.Tag = ""
    PRECIO.Text = ""
    CANTIDAD.Text = ""
    TOTAL.Text = ""
    TOTALS.Text = 0
    For I = 1 To F.Rows - 1
        TOTALS.Text = Val(TOTALS.Text) + Val(F.TextMatrix(I, 4))
    Next
    A = MsgBox("DESEA SEGUIR INGRESANDO PRODUCTO", vbYesNo)
    If A = vbYes Then
        PRODUCTO.SetFocus
    Else
        TOTALS.SetFocus
    End If
End If
End Sub

Private Sub TOTALG_KeyPress(KeyAscii As Integer)
'On Error Resume Next
If KeyAscii = 13 Then
    ssql = "SELECT * FROM MOVIMIENTO WHERE CODMOVIMIENTO='" & NDOCUMENTO.Text & "';"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        MsgBox "YA EXISTE UN DOCUMENTO CON EL NUMERO :" & NDOCUMENTO.Text
        NDOCUMENTO.SelStart = 1
        NDOCUMENTO.SelLength = Len(NDOCUMENTO.Text)
        NDOCUMENTO.SetFocus
        Exit Sub
    End If
    ssql = "INSERT INTO MOVIMIENTO VALUES('" & NDOCUMENTO.Text & "'," & FECHAS(FECHA, True) & "," & TIPO.Tag & ",'" & CONCEPTO.Text & "'," & Val(TOTALS.Text) & "," & Val(IGV.Text) & "," & Val(TOTALG.Text) & ");"
'MsgBox ssql
    Set tbl = conn.Execute(ssql)
    ssql = "DELETE FROM DETMOVIMIENTO WHERE CODMOVIMIENTO='" & NDOCUMENTO.Text & "';"
    Set tbl = conn.Execute(ssql)
    For I = 1 To F.Rows - 1
        ssql = "INSERT INTO DETMOVIMIENTO VALUES('" & NDOCUMENTO.Text & "','" & F.TextMatrix(I, 0) & "'," & Val(F.TextMatrix(I, 2)) & "," & Val(F.TextMatrix(I, 3)) & ");"
        'MsgBox ssql
        Set tbl = conn.Execute(ssql)
    Next
ssql = "SELECT * FROM TIPO WHERE TIPO='" & TIPO.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    If Val(tbl!D) < 2 Then
        ssql = "SELECT * FROM CAJA ORDER BY CODCAJA DESC;"
        Set tbl = conn.Execute(ssql)
        COD = 0
        If tbl.EOF = False Then
            COD = Val(tbl!CODCAJA)
        End If
        COD = Val(COD) + 1
        ssql = "INSERT INTO CAJA VALUES(" & COD & "," & FECHAS(FECHA, True) & ",'MM" & NDOCUMENTO.Text & "','" & CONCEPTO.Text & "'," & TIPO.Tag & "," & TOTALG.Text & ")"
        Set tbl = conn.Execute(ssql)
    End If
End If
ssql = "SELECT * FROM MOVIMIENTO WHERE CODMOVIMIENTO='" & NDOCUMENTO.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = True Then MsgBox "NO SE LOGRO GRABAR"
If Err Then MsgBox Err.Description
Unload Me
End If

End Sub

Private Sub TOTALS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        ssql = "SELECT * FROM IGV;"
        Set tbl = conn.Execute(ssql)
        If tbl.EOF = False Then
            IGV.Text = Val(TOTALS.Text) * Val(tbl!IGV)
        Else
            MsgBox "NO ESTA CONFIGURADO EL IGV": Exit Sub
        End If
            IGV.SetFocus
End If
End Sub
