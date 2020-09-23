VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FRMDOCTOR 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGISTRO DE DOCTORES"
   ClientHeight    =   8235
   ClientLeft      =   2040
   ClientTop       =   2025
   ClientWidth     =   11355
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
   ScaleWidth      =   11355
   Begin VB.ListBox ESPECIALIDAD 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   1980
      Left            =   4605
      TabIndex        =   19
      Top             =   4905
      Width           =   4110
   End
   Begin VB.ComboBox ESPE 
      Height          =   315
      Left            =   4605
      TabIndex        =   18
      Top             =   4455
      Width           =   4125
   End
   Begin MSMask.MaskEdBox INI 
      Height          =   375
      Left            =   6210
      TabIndex        =   14
      Top             =   3780
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   16711680
      Format          =   "hh:mm"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox DOCTOR 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4770
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1305
      Width           =   6450
   End
   Begin VB.TextBox TELEFONO1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   4770
      TabIndex        =   5
      Top             =   1845
      Width           =   3105
   End
   Begin VB.TextBox TELEFONO2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   4770
      TabIndex        =   4
      Top             =   2325
      Width           =   3105
   End
   Begin VB.TextBox DIRECCION 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   4770
      TabIndex        =   3
      Top             =   2805
      Width           =   6450
   End
   Begin VB.TextBox CODDOCTOR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   7215
      TabIndex        =   0
      Top             =   60
      Width           =   2895
   End
   Begin LVbuttons.LaVolpeButton ELIMINAR 
      Height          =   465
      Left            =   9435
      TabIndex        =   2
      Top             =   6960
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
      MICON           =   "FRMDOCTOR.frx":0000
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
   Begin LVbuttons.LaVolpeButton ACEPTAR 
      Height          =   465
      Left            =   7515
      TabIndex        =   7
      Top             =   6960
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
      MICON           =   "FRMDOCTOR.frx":001C
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
      Left            =   9420
      TabIndex        =   8
      Top             =   7455
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
      MICON           =   "FRMDOCTOR.frx":0038
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
   Begin MSMask.MaskEdBox FIN 
      Height          =   375
      Left            =   9345
      TabIndex        =   16
      Top             =   3780
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   16711680
      Format          =   "hh:mm"
      PromptChar      =   "_"
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   345
      Left            =   8745
      TabIndex        =   20
      Top             =   4425
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
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
      MICON           =   "FRMDOCTOR.frx":0054
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
      Height          =   345
      Left            =   8745
      TabIndex        =   21
      Top             =   4905
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "QUITAR"
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
      MICON           =   "FRMDOCTOR.frx":0070
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
   Begin MSMask.MaskEdBox DIA 
      Height          =   375
      Left            =   6210
      TabIndex        =   22
      Top             =   3330
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   16711680
      PromptChar      =   "_"
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   24
      Top             =   7965
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9737
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9737
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DIAS LABORABLES:"
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
      Index           =   9
      Left            =   4110
      TabIndex        =   23
      Top             =   3390
      Width           =   2070
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
      Index           =   8
      Left            =   2865
      TabIndex        =   17
      Top             =   4500
      Width           =   1590
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HORARIO FIN:"
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
      Left            =   7800
      TabIndex        =   15
      Top             =   3840
      Width           =   1515
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   4695
      X2              =   11295
      Y1              =   3270
      Y2              =   3270
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HORARIO INI:"
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
      Left            =   4710
      TabIndex        =   13
      Top             =   3840
      Width           =   1470
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
      Left            =   3585
      TabIndex        =   12
      Top             =   1335
      Width           =   885
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TELEFONO 1"
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
      Left            =   3135
      TabIndex        =   11
      Top             =   1890
      Width           =   1320
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TELEFONO 2"
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
      Left            =   3135
      TabIndex        =   10
      Top             =   2355
      Width           =   1320
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   2835
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   4590
      X2              =   4590
      Y1              =   870
      Y2              =   3105
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO DOCTOR"
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
      Left            =   5295
      TabIndex        =   1
      Top             =   75
      Width           =   1815
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4665
      X2              =   10185
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Image Image1 
      Height          =   5250
      Left            =   -285
      Picture         =   "FRMDOCTOR.frx":008C
      Top             =   2790
      Width           =   3480
   End
End
Attribute VB_Name = "FRMDOCTOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ACEPTAR_Click()
On Error Resume Next
ssql = "SELECT * FROM DOCTOR WHERE CODDOCTOR=" & Val(CODDOCTOR.Text) & ";"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    ssql = "UPDATE DOCTOR SET DOCTOR='" & DOCTOR & "',TELEFONO1='" & TELEFONO1.Text & "',TELEFONO2='" & TELEFONO2.Text & "',DIRECCION='" & DIRECCION.Text & "',DIA=" & Val(DIA.Text) & ",HORAINI='" & INI.Text & "',HORAFIN='" & FIN.Text & "' WHERE CODDOCTOR=" & CODDOCTOR.Text & ""
Else
    ssql = "INSERT INTO DOCTOR VALUES(" & CODDOCTOR.Text & ",'" & DOCTOR.Text & "','" & TELEFONO1.Text & "','" & TELEFONO2.Text & "','" & DIRECCION.Text & "'," & Val(DIA.Text) & ",'" & INI.Text & "','" & FIN.Text & "');"
End If
'MsgBox ssql
Set tbl = conn.Execute(ssql)
ssql = "DELETE FROM DETDOCTOR WHERE CODDOCTOR=" & Val(CODDOCTOR.Text) & ";"
Set tbl = conn.Execute(ssql)
For I = 0 To Val(ESPECIALIDAD.ListCount) - 1
    ssql = "SELECT * FROM ESPECIALIDAD WHERE ESPECIALIDAD='" & ESPECIALIDAD.List(I) & "';"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        ssql = "INSERT INTO DETDOCTOR VALUES(" & CODDOCTOR.Text & "," & Val(tbl!CODESPECIALIDAD) & ");"
        Set tbl = conn.Execute(ssql)
    End If
Next
If Err Then MsgBox Err.Description
Form_Load
End Sub

Private Sub CODDOCTOR_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    ssql = "select * from doctor where coddoctor=" & CODDOCTOR.Text & ";"
    Set tbl = conn.Execute(ssql)
    If tbl.EOF = False Then
        DOCTOR.Text = tbl!DOCTOR
        TELEFONO1.Text = tbl!TELEFONO1
        TELEFONO2.Text = tbl!TELEFONO2
        DIRECCION.Text = tbl!DIRECCION
        INI.Text = tbl!HORAINI
        FIN.Text = tbl!HORAFIN
        DIA.Text = tbl!DIA
        ssql = "SELECT DETDOCTOR.CODDOCTOR, DETDOCTOR.CODESPECIALIDAD, ESPECIALIDAD.ESPECIALIDAD " & _
                "FROM ESPECIALIDAD INNER JOIN DETDOCTOR ON ESPECIALIDAD.CODESPECIALIDAD = DETDOCTOR.CODESPECIALIDAD " & _
                "where DETDOCTOR.coddoctor=" & tbl!CODDOCTOR & " ORDER BY ESPECIALIDAD.ESPECIALIDAD;"
        Set tbl = conn.Execute(ssql)
        ESPECIALIDAD.Clear
        Do Until tbl.EOF
            ESPECIALIDAD.AddItem tbl!ESPECIALIDAD
            tbl.MoveNext
        Loop
        DOCTOR.SetFocus
    Else
        DOCTOR.Text = ""
        'DNI.Text = ""
        TELEFONO1.Text = ""
        TELEFONO2.Text = ""
        DIRECCION.Text = ""
        INI.Text = ""
        FIN.Text = ""
        ESPECIALIDAD.Clear
        DOCTOR.SetFocus
    End If
End If
If Err Then MsgBox Err.Description
End Sub

Private Sub DIA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then INI.SetFocus
End Sub

Private Sub DIRECCION_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DIA.SetFocus
End Sub

Private Sub DNI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TELEFONO1.SetFocus
End Sub

Private Sub DOCTOR_Click()
ssql = "SELECT * FROM DOCTOR WHERE DOCTOR='" & DOCTOR.Text & "';"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    CODDOCTOR.Text = tbl!CODDOCTOR
End If
    CODDOCTOR_KeyPress 13
    
End Sub

Private Sub DOCTOR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TELEFONO1.SetFocus
    
End Sub

Private Sub ELIMINAR_Click()
ssql = "DELETE FROM DOCTOR WHERE CODDOCTOR=" & Val(CODDOCTOR.Text) & ";"
Set tbl = conn.Execute(ssql)
Form_Load
End Sub

Private Sub ESPE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    For I = 0 To Val(ESPECIALIDAD.ListCount) - 1
        If ESPE.Text = ESPECIALIDAD.List(I) Then Exit Sub
    Next
    ESPECIALIDAD.AddItem ESPE.Text
    A = MsgBox("DESEA SEGUIR AGREGANDO ESPECIALIDADES", vbYesNo)
    If A = vbNo Then ACEPTAR_Click
End If
End Sub

Private Sub FIN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ESPE.SetFocus
End Sub

Private Sub Form_Load()
CODDOCTOR.Text = ""
DOCTOR.Text = ""
TELEFONO1.Text = ""
TELEFONO2.Text = ""
DIRECCION.Text = ""
ESPE.Clear
ESPECIALIDAD.Clear
INI.Text = ""
FIN.Text = ""
ssql = "SELECT * FROM DOCTOR ORDER BY DOCTOR;"
Set tbl = conn.Execute(ssql)
DOCTOR.Clear
Do Until tbl.EOF
    DOCTOR.AddItem tbl!DOCTOR
    tbl.MoveNext
Loop
ssql = "select * from especialidad order by especialidad;"
Set tbl = conn.Execute(ssql)
ESPE.Clear
Do Until tbl.EOF
    ESPE.AddItem tbl!ESPECIALIDAD
    tbl.MoveNext
Loop
End Sub

Private Sub INI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then FIN.SetFocus
End Sub

Private Sub LaVolpeButton1_Click()
FRMESPECIALIDAD.Show vbModal
ssql = "select * from especialidad order by especialidad;"
Set tbl = conn.Execute(ssql)
ESPE.Clear
Do Until tbl.EOF
    ESPE.AddItem tbl!ESPECIALIDAD
    tbl.MoveNext
Loop

End Sub

Private Sub LaVolpeButton2_Click()
If Val(ESPECIALIDAD.ListIndex) > 0 Then
    ESPECIALIDAD.RemoveItem ESPECIALIDAD.ListIndex
Else
    ESPECIALIDAD.Clear
End If
End Sub

Private Sub SALIR_Click()
Unload Me
End Sub

Private Sub TELEFONO1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TELEFONO2.SetFocus
End Sub

Private Sub TELEFONO2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DIRECCION.SetFocus
End Sub
