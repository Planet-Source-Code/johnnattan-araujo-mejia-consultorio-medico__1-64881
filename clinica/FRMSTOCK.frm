VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRMSTOCK 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK DE PRODUCTOS"
   ClientHeight    =   10560
   ClientLeft      =   1230
   ClientTop       =   465
   ClientWidth     =   12510
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
   ScaleHeight     =   10560
   ScaleWidth      =   12510
   Begin VB.TextBox PRODUCTO 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2280
      TabIndex        =   4
      Top             =   60
      Width           =   7620
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   6780
      Left            =   930
      TabIndex        =   3
      Top             =   765
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   11959
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid F 
      Height          =   9120
      Left            =   3240
      TabIndex        =   0
      Top             =   345
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   16087
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
      MouseIcon       =   "FRMSTOCK.frx":0000
   End
   Begin LVbuttons.LaVolpeButton ACEPTAR 
      Height          =   465
      Left            =   7740
      TabIndex        =   1
      Top             =   9630
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
      MICON           =   "FRMSTOCK.frx":031A
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
      Left            =   9675
      TabIndex        =   2
      Top             =   9630
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
      MICON           =   "FRMSTOCK.frx":0336
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
      Top             =   10290
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10769
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10769
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   -435
      Picture         =   "FRMSTOCK.frx":0352
      Top             =   7695
      Width           =   4500
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CRITERIO BUSQ."
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
      Index           =   0
      Left            =   465
      TabIndex        =   5
      Top             =   105
      Width           =   1755
   End
   Begin VB.Line Line1 
      X1              =   2295
      X2              =   11685
      Y1              =   9540
      Y2              =   9540
   End
End
Attribute VB_Name = "FRMSTOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ACEPTAR_Click()
Dim DT As New ADODB.Recordset
DT.Fields.Append "CODIGO", adChar, 255
DT.Fields.Append "PRODUCTO", adChar, 255
DT.Fields.Append "STOCK", adChar, 255
DT.Open
For I = 1 To F.Rows - 1
    DT.AddNew
    DT!CODIGO = F.TextMatrix(I, 0)
    DT!PRODUCTO = F.TextMatrix(I, 1)
    DT!STOCK = F.TextMatrix(I, 2)
Next
Set DTSTOCK.DataSource = DT
DTSTOCK.LeftMargin = 200
DTSTOCK.RightMargin = 200
DTSTOCK.Show vbModal
End Sub

Private Sub Form_Load()
'Me.Visible = True
PRODUCTO_KeyPress 13
            

End Sub

Private Sub PRODUCTO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
F.Rows = 1
F.FormatString = "CODIGO|PRODUCTO|STOCK"
F.ColWidth(0) = 1500
F.ColWidth(1) = 5500
F.ColWidth(2) = 1500
ssql = "sELECT * FROM PRODUCTO WHERE PRODUCTO LIKE '%" & PRODUCTO.Text & "%' ORDER BY CODPRODUCTO;"
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
    STOCK = 0
    Do Until TB.EOF
        If Val(TB!T) = 1 Then
            STOCK = Val(STOCK) + Val(TB!CANTIDAD)
        Else
            STOCK = Val(STOCK) - Val(TB!CANTIDAD)
        End If
        TB.MoveNext
    Loop
    If Val(STOCK) > 0 Then
        F.AddItem UCase(tbl!CODPRODUCTO & vbTab & tbl!PRODUCTO & vbTab & Format(STOCK, "###,###,##0.00"))
    End If
    tbl.MoveNext
Loop
End If

End Sub

Private Sub SALIR_Click()
Unload Me
End Sub

