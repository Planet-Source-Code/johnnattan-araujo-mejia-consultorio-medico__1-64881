VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRMICREDITO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESOS DE CREDITOS"
   ClientHeight    =   1260
   ClientLeft      =   4515
   ClientTop       =   3330
   ClientWidth     =   8370
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
   ScaleHeight     =   1260
   ScaleWidth      =   8370
   Begin VB.ComboBox PACIENTE 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   90
      Width           =   7125
   End
   Begin LVbuttons.LaVolpeButton ACEPTAR 
      Default         =   -1  'True
      Height          =   465
      Left            =   6420
      TabIndex        =   2
      Top             =   435
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "ACEPTAR"
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
      MICON           =   "FRMICREDITO.frx":0000
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
      TabIndex        =   3
      Top             =   990
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
      EndProperty
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
      Left            =   75
      TabIndex        =   1
      Top             =   105
      Width           =   1065
   End
End
Attribute VB_Name = "FRMICREDITO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ACEPTAR_Click()
ssql = "SELECT * FROM PACIENTE where paciente='" & PACIENTE.Text & "' ORDER BY PACIENTE;"
Set tbl = conn.Execute(ssql)
Dim TB As New ADODB.Recordset

Do Until tbl.EOF
    ssql = "SELECT * FROM CAJA WHERE (CODTIPO=5 OR CODTIPO=4) AND CONcEPTO LIKE '%" & tbl!CODPACIENTE & "%' order by FECHA,CODTIPO ;"
    Set TB = conn.Execute(ssql)
    credito = 0
    pago = 0

    
    Do Until TB.EOF
        If Val(TB!CODTIPO) = 4 Then
            credito = Val(credito) + Val(TB!IMPORTE)
        ElseIf Val(TB!CODTIPO) = 5 Then
            pago = Val(pago) + Val(TB!IMPORTE)
        End If
        TB.MoveNext
    Loop
    If Val(credito) - Val(pago) > 0 Then
    FRMCAJA.IMPORTE.Text = Val(credito) - Val(pago)
    End If

tbl.MoveNext
Loop
ssql = "select * from paciente where paciente='" & PACIENTE.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = True Then Exit Sub
PACIENTE.Tag = tbl!CODPACIENTE
FRMCAJA.CONCEPTO.Text = "CREDITO " & FRMCAJA.TIPO.Text & ":" & PACIENTE.Tag & "-" & PACIENTE.Text
Unload Me
End Sub

Private Sub Form_Load()
If Val(FRMCAJA.TIPO.Tag) = 5 Then

    ssql = "SELECT * FROM PACIENTE ORDER BY PACIENTE;"
    Set tbl = conn.Execute(ssql)
    Dim TB As New ADODB.Recordset
    
    Do Until tbl.EOF
        ssql = "SELECT * FROM CAJA WHERE (CODTIPO=5 OR CODTIPO=4) AND CONCEPTO LIKE '%" & tbl!CODPACIENTE & "%' order by FECHA,CODTIPO ;"
        Set TB = conn.Execute(ssql)
        credito = 0
        pago = 0
        
        Do Until TB.EOF
            If Val(TB!CODTIPO) = 4 Then
                credito = Val(credito) + Val(TB!IMPORTE)
            ElseIf Val(TB!CODTIPO) = 5 Then
                pago = Val(pago) + Val(TB!IMPORTE)
            End If
            TB.MoveNext
        Loop
        If Val(credito) - Val(pago) > 0 Then
        PACIENTE.AddItem tbl!PACIENTE
        End If
    
    tbl.MoveNext
    Loop
Else
    ssql = "SELECT * FROM PACIENTE ORDER BY PACIENTE;"
    Set tbl = conn.Execute(ssql)
    PACIENTE.Clear
    Do Until tbl.EOF
        PACIENTE.AddItem tbl!PACIENTE
        tbl.MoveNext
    Loop
End If
End Sub

