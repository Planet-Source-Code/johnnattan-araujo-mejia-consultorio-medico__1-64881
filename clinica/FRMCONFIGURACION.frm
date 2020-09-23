VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRMCONFIGURACION 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONFIGURACION"
   ClientHeight    =   9330
   ClientLeft      =   2115
   ClientTop       =   1050
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   12480
   Begin VB.CommandButton IGV 
      Caption         =   "IGV"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7200
      TabIndex        =   6
      Top             =   4815
      Width           =   2040
   End
   Begin LVbuttons.LaVolpeButton CB 
      Height          =   405
      Index           =   0
      Left            =   4905
      TabIndex        =   0
      Top             =   3855
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "PACIENTE"
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
      MICON           =   "FRMCONFIGURACION.frx":0000
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
      Index           =   1
      Left            =   4905
      TabIndex        =   1
      Top             =   4320
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "DOCTOR"
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
      MICON           =   "FRMCONFIGURACION.frx":031A
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
      Index           =   2
      Left            =   4920
      TabIndex        =   2
      Top             =   4800
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "ESPECIALIDAD"
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
      MICON           =   "FRMCONFIGURACION.frx":0634
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
      Index           =   3
      Left            =   7155
      TabIndex        =   3
      Top             =   3855
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "TIPO DOCUMENTO"
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
      MICON           =   "FRMCONFIGURACION.frx":094E
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
      Index           =   4
      Left            =   7155
      TabIndex        =   4
      Top             =   4320
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "PRODUCTO"
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
      MICON           =   "FRMCONFIGURACION.frx":0C68
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
      Left            =   9180
      TabIndex        =   5
      Top             =   5640
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
      MICON           =   "FRMCONFIGURACION.frx":0F82
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
      Top             =   9060
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10742
            Text            =   "SISTEMA CLINICO"
            TextSave        =   "SISTEMA CLINICO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10742
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton CB 
      Height          =   405
      Index           =   5
      Left            =   4920
      TabIndex        =   8
      Top             =   5280
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "USUARIO"
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
      MICON           =   "FRMCONFIGURACION.frx":0F9E
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
   Begin VB.Shape Shape1 
      Height          =   2955
      Left            =   3975
      Top             =   3195
      Width           =   6315
   End
   Begin VB.Image Image1 
      Height          =   5205
      Left            =   -45
      Picture         =   "FRMCONFIGURACION.frx":12B8
      Top             =   -105
      Width           =   3000
   End
End
Attribute VB_Name = "FRMCONFIGURACION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB_Click(Index As Integer)
Select Case Index
Case 5
    FRMUSUARIO.Show vbModal
Case 0
    FRMPACIENTE.Show vbModal
Case 1
    FRMDOCTOR.Show vbModal
Case 2
    FRMESPECIALIDAD.Show vbModal
Case 3
    frmdocumento.Show vbModal
Case 4
    FRMPRODUCTO.Show vbModal
End Select
End Sub

Private Sub IGV_Click()
ssql = "select * from igv;"
Set tbl = conn.Execute(ssql)
If tbl.EOF = False Then
    G = Val(tbl!IGV) * 100
Else
    G = ""
End If
A = InputBox("INGRESE EL PROCENTAJE DEL IGV", "SISTEMA CLINICO", G)
If IsNumeric(A) = False Then Exit Sub
ssql = "DELETE FROM IGV;"
Set tbl = conn.Execute(ssql)
ssql = "INSERT INTO IGV VALUES(" & Val(A) / 100 & ");"
Set tbl = conn.Execute(ssql)
End Sub

Private Sub LaVolpeButton4_Click()
Unload Me
End Sub
