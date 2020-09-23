Attribute VB_Name = "Module1"
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public conn As New ADODB.Connection
Public tbl As New ADODB.Recordset
Public ssql As String
Public s As String
Public s1 As String
Public s2 As String

Sub main()
'If DateDiff("d", Date, "17/09/2005") = 0 Then End
conn.ConnectionString = "Provider=msdatashape;" & _
"data Provider=Microsoft.Jet.OLEDB.4.0; " & _
     "Data Source=" & App.Path & "\CLINICA.mdb" & ";" & _
     "Jet OLEDB:Database" 'Password=456"
conn.Open
FRMPRINCIPAL.Show

End Sub
Public Function FECHAS(OBJETO As Object, EXTRAER As Boolean)
If EXTRAER = True Then
    If TypeName(OBJETO) = "MaskEdBox" Then
        FECHAS = Format(Right(OBJETO.Text, 4), "0000") & Format(Mid(OBJETO.Text, 4, 2), "00") & Format(Left(OBJETO.Text, 2), "00")
    Else
        FECHAS = Format(OBJETO.Year, "0000") & Format(OBJETO.Month, "00") & Format(OBJETO.Day, "00")
    End If
Else
    
    FECHAS = Right(OBJETO!FECHA, 2) & "/" & Mid(OBJETO!FECHA, 5, 2) & "/" & Left(OBJETO!FECHA, 4)
End If
End Function
Public Sub Limpiar(FORMULARIO As Form)
On Error Resume Next
Dim x As Object
For Each x In FORMULARIO.Controls
    If (TypeName(x) = "TextBox" Or TypeName(x) = "ComboBox" Or TypeName(x) = "MaskEdBox") Then
        x.Text = ""
    ElseIf TypeName(x) = "DTPicker" Then
        x.Value = Now
    End If
Next
End Sub

Public Sub Activar(ESTADO As Boolean, FORMULARIO As Form)
    If ESTADO = True Then
        FORMULARIO.CB(0).Enabled = True
        FORMULARIO.CB(1).Enabled = True
        FORMULARIO.CB(2).Enabled = False
        FORMULARIO.CB(3).Enabled = False
        FORMULARIO.CB(4).Enabled = True
    Else
        FORMULARIO.CB(0).Enabled = False
        FORMULARIO.CB(1).Enabled = False
        FORMULARIO.CB(2).Enabled = True
        FORMULARIO.CB(3).Enabled = True
        FORMULARIO.CB(4).Enabled = False
    End If
End Sub
Public Sub PROBAR(FORMULARIO As Form, respuesta As Long)
On Error Resume Next
Dim x As Object
For Each x In FORMULARIO.Controls
    If TypeName(x) = "TextBox" Or TypeName(x) = "ComboBox" Then
        If Val(x.Tag) = 1 Then
            'MsgBox x.Name
            If x.Text = "" Then
            
                MsgBox "ES NECESARIO QUE INGRESE EL SIGUIENTE DATO", vbInformation
                x.SetFocus
                respuesta = 1
                Exit Sub
            End If
        Else
            If x.Text = "" Then
            x.Text = "_"
            End If
        End If
    End If
Next
respuesta = 0
End Sub

