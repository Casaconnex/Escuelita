Attribute VB_Name = "Funciones"
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub ConexionODBC(nombre As String)
    Dim llamar As Double
    On Error Resume Next
    llamar = Shell(nombre, 5)
End Sub

Public Sub ConexionBD(Formulario As Object, consulta As String)
'esta funcion utiliza odbc para conectarse a la bd
On Error Resume Next
Formulario.bd.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PROYECTO.mdb;Mode=Share Deny None;Persist Security Info=False;Jet OLEDB:Database Password=810205;Jet OLEDB:Engine Type=4"
Formulario.bd.RecordSource = consulta
Formulario.bd.Refresh
End Sub
Public Sub ConexionBD3(Formulario As Object, consulta As String)
'esta funcion utiliza odbc para conectarse a la bd
On Error Resume Next
Formulario.bd3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PROYECTO.mdb;Mode=Share Deny None;Persist Security Info=False;Jet OLEDB:Database Password=810205;Jet OLEDB:Engine Type=4"
Formulario.bd3.RecordSource = consulta
Formulario.bd3.Refresh
End Sub
Public Sub ConexionDocu(Formulario As Object, consulta As String)
'esta funcion utiliza odbc para conectarse a la bd
On Error Resume Next
Formulario.docu.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PROYECTO.mdb;Mode=Share Deny None;Persist Security Info=False;Jet OLEDB:Database Password=810205;Jet OLEDB:Engine Type=4"
Formulario.docu.RecordSource = consulta
Formulario.docu.Refresh
End Sub
Public Sub ConexionBD1(Formulario As Object, consulta As String)
'esta funcion utiliza odbc para conectarse a la bd
On Error Resume Next
Formulario.bd1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PROYECTO.mdb;Mode=Share Deny None;Persist Security Info=False;Jet OLEDB:Database Password=810205;Jet OLEDB:Engine Type=4"
Formulario.bd1.RecordSource = consulta
Formulario.bd1.Refresh
End Sub
Public Sub ConexionBD2(Formulario As Object, consulta As String)
'esta funcion utiliza odbc para conectarse a la bd
On Error Resume Next
Formulario.bd2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PROYECTO.mdb;Mode=Share Deny None;Persist Security Info=False;Jet OLEDB:Database Password=810205;Jet OLEDB:Engine Type=4"
Formulario.bd2.RecordSource = consulta
Formulario.bd2.Refresh
End Sub
Public Sub CargarUsuarios()
'cargar usuarios registrados
acceso.bd.Recordset.MoveFirst
For i = 1 To acceso.bd.Recordset.RecordCount
    acceso.user.AddItem acceso.bd.Recordset!usuario
    acceso.bd.Recordset.MoveNext
Next i

End Sub

Public Sub BuscarU(userb As String)
'busca el usuario en la bd
acceso.bd.Recordset.MoveFirst
For i = 0 To acceso.bd.Recordset.RecordCount
    If acceso.bd.Recordset!usuario = userb Then
        passw = acceso.bd.Recordset!clave
        perfil = acceso.bd.Recordset!perfil
        Exit For
    End If
    acceso.bd.Recordset.MoveNext
Next i

End Sub
Public Sub BuscarUs(userb As String)
'busca el usuario en la bd
Perfiles.bd1.Recordset.MoveFirst
For i = 0 To Perfiles.bd1.Recordset.RecordCount
    If Perfiles.bd1.Recordset!usuario = userb Then
        perfilito = Perfiles.bd1.Recordset!perfil
        Exit For
    End If
    Perfiles.bd1.Recordset.MoveNext
Next i

End Sub

Public Sub CargarCuentas()
Perfiles.cuentas.Clear
Perfiles.bd.Recordset.MoveFirst
For i = 1 To Perfiles.bd.Recordset.RecordCount
    Perfiles.cuentas.AddItem Perfiles.bd.Recordset!usuario
    Perfiles.bd.Recordset.MoveNext
Next i
End Sub
Public Sub ConfigurarCuenta(NombreCuenta As String)
If NombreCuenta = "Administrador" Then
    Perfiles.cnombre.Locked = True
End If
If Perfiles.cuentas.ListIndex = -1 Then
    MsgBox "Debe seleccionar una cuenta para configurarla!", vbInformation, "Configurar"
    Exit Sub
End If
Perfiles.nuevo.Enabled = False
Perfiles.eliminar.Enabled = False
Perfiles.Height = 7155
Perfiles.frame3.Visible = True
Perfiles.frame2.Visible = False
Perfiles.cnombre.Text = NombreCuenta
CambiarNomUsuario = NombreCuenta
Perfiles.cpw.SetFocus

End Sub

Public Sub EliminarCuenta(NombreCuenta As String)
If NombreCuenta = "Administrador" Then
    MsgBox "Imposible eliminar la cuenta de Administrador", vbCritical, "Eliminar cuentas"
    Exit Sub
End If
If NombreCuenta = usuario Then
    MsgBox "Imposible eliminar esta cuenta porque se encuentra es uso!", vbCritical, "Eliminar cuentas"
    Exit Sub
End If
If Perfiles.cuentas.ListIndex = -1 Then
    MsgBox "Debe seleccionar una cuenta para eliminarla!", vbInformation, "Eliminar"
    Exit Sub
End If
If perfilito = 1 Then
    Perfiles.bd.Recordset.MoveFirst
    For i = 0 To Perfiles.bd.Recordset.RecordCount
        If Perfiles.bd.Recordset!usuario = NombreCuenta Then
            Perfiles.bd.Recordset.Delete adAffectCurrent
            CargarCuentas
            Exit For
        End If
        Perfiles.bd.Recordset.MoveNext
    Next i
Else
    MsgBox "No tiene permisos para realizar esta acción!", vbExclamation, "Configuración de Usuarios"
End If
End Sub

Public Sub CuentaNueva()
On Error GoTo Error
Perfiles.bd.Recordset.AddNew
Perfiles.bd.Recordset!usuario = Perfiles.nombre.Text
Perfiles.bd.Recordset!clave = Perfiles.pw.Text
If Perfiles.admin.Value = True Then
    Perfiles.bd.Recordset!perfil = 1
ElseIf Perfiles.normal.Value = True Then
    Perfiles.bd.Recordset!perfil = 2
End If
Perfiles.bd.Recordset.Update
CargarCuentas
SalirError:
    Exit Sub
Error:
    If Err.Number Then
        valor = Mid(des, InStr(des, "[Administrador de controladores ODBC]") + 38)
       MsgBox valor, vbInformation, "Error"
    End If

End Sub

Public Sub CambiarCuenta()
Perfiles.bd.Recordset.MoveFirst
For i = 0 To Perfiles.bd.Recordset.RecordCount
    If Perfiles.bd.Recordset!usuario = CambiarNomUsuario Then
        Perfiles.bd.Recordset.Delete adAffectCurrent
        Perfiles.bd.Recordset.AddNew
        Perfiles.bd.Recordset!usuario = Perfiles.cnombre.Text
        Perfiles.bd.Recordset!clave = Perfiles.cpw.Text
        If Perfiles.admin1.Visible = False And Perfiles.normal1.Visible = False Then
            Perfiles.bd.Recordset!perfil = 1
        End If
        If Perfiles.admin1.Value = True Then
            Perfiles.bd.Recordset!perfil = 1
        ElseIf Perfiles.normal1.Value = True Then
            Perfiles.bd.Recordset!perfil = 2
        End If
        Perfiles.bd.Recordset.Update
        CargarCuentas
        Exit For
    End If
    Perfiles.bd.Recordset.MoveNext
Next i

End Sub

Public Function Horizontal(Newform As Form, Colour1 As ColorConstants, Colour2 As ColorConstants)
'apariencia de la barra de carga al estilo xp
    Dim X As Integer
    Dim VR, VG, VB As Single
    Dim Color1, Color2 As Long
    Dim R, G, b, R2, G2, B2 As Integer
    Dim temp As Long

    Color1 = Colour1
    Color2 = Colour2

    temp = (Color1 And 255)
    R = temp And 255
    temp = Int(Color1 / 256)
    G = temp And 255
    temp = Int(Color1 / 65536)
    b = temp And 255
    temp = (Color2 And 255)
    R2 = temp And 255
    temp = Int(Color2 / 256)
    G2 = temp And 255
    temp = Int(Color2 / 65536)
    B2 = temp And 255

    VR = Abs(R - R2) / Newform.ScaleWidth
    VG = Abs(G - G2) / Newform.ScaleWidth
    VB = Abs(b - B2) / Newform.ScaleWidth

    If R2 < R Then VR = -VR
    If G2 < G Then VG = -VG
    If B2 < b Then VB = -VB

    For X = 0 To Newform.ScaleWidth
        R2 = R + VR * X
        G2 = G + VG * X
        B2 = b + VB * X
        Newform.Line (X, 0)-(X, Newform.ScaleHeight), RGB(R2, G2, B2)
    Next X
End Function

Public Function Validar_letra(Caracter As Integer) As Integer
If Caracter <> 13 Then
    If Caracter <> 8 Then
        If Caracter <> 32 Then
            If (Caracter = 241 Or Caracter = 209) Then
            Validar_letra = Caracter
            Exit Function
            End If
            If Caracter >= 65 And Caracter <= 90 Then
            
                Validar_letra = Caracter
            ElseIf Caracter >= 97 And Caracter <= 122 Then
                Validar_letra = Caracter
            Else
                Validar_letra = 0
            End If
        Else
            Validar_letra = Caracter
        End If
    Else
        Validar_letra = Caracter
    End If
Else
    Validar_letra = Caracter
End If
End Function

Public Function Validar_numero(numero As Integer) As Integer
If numero <> 13 Then
    If numero <> 8 Then
        If numero >= 48 And numero <= 57 Then
            Validar_numero = numero
        Else
            Validar_numero = 0
        End If
    Else
        Validar_numero = numero
    End If
Else
    Validar_numero = numero
End If
End Function

Public Function mes(MesId As Integer) As String
Select Case MesId
    Case 1: mes = "Enero"
    Case 2: mes = "Febrero"
    Case 3: mes = "Marzo"
    Case 4: mes = "Abril"
    Case 5: mes = "Mayo"
    Case 6: mes = "Junio"
    Case 7: mes = "Julio"
    Case 8: mes = "Agosto"
    Case 9: mes = "Septiembre"
    Case 10: mes = "Octubre"
    Case 11: mes = "Noviembre"
    Case 12: mes = "Diciembre"
End Select
End Function

Public Sub cajas(Formulario As Form)
Dim vari
For Each vari In Formulario.Controls
    If TypeOf vari Is TextBox Or TypeOf vari Is ComboBox Then
        If vari.Tag = "" Then
            vari.Text = ""
            vari.Locked = False
        End If
    End If
Next
End Sub
Public Sub cajasM(Formulario As Form)
Dim vari
For Each vari In Formulario.Controls
    If TypeOf vari Is TextBox Or TypeOf vari Is ComboBox Then
        If vari.Tag = "" Or vari.Tag <> "" Then
            vari.Text = ""
            vari.Locked = False
        End If
    End If
Next
End Sub
Public Sub cajasl(Formulario As Form)
Dim vari
For Each vari In Formulario.Controls
    If TypeOf vari Is TextBox Or TypeOf vari Is ComboBox Then
        If vari.Tag <> "1" Then
           vari.Text = ""
        End If
        vari.Locked = False
    End If
Next
End Sub
Public Sub Deshabilitarl(Formulario As Form)
Dim vari
For Each vari In Formulario.Controls
    If TypeOf vari Is TextBox Or TypeOf vari Is ComboBox Then
        vari.Locked = True
    End If
Next
End Sub
Public Sub Deshabilitari(Formulario As Form)
Dim vari
For Each vari In Formulario.Controls
    If TypeOf vari Is ComboBox Then
        vari.Locked = True
    End If
Next
End Sub
Public Sub Habilitei(Formulario As Form)
Dim vari
For Each vari In Formulario.Controls
    If TypeOf vari Is ComboBox Or TypeOf vari Is TextBox Then
        vari.Locked = False
    End If
Next
End Sub
Public Sub Habilitarl(Formulario As Form)
Dim vari
For Each vari In Formulario.Controls
    If TypeOf vari Is TextBox Or TypeOf vari Is ComboBox Then
        vari.Locked = False
    End If
Next
End Sub

Public Sub ModificarCajas(Formulario As Form)
Dim vari
For Each vari In Formulario.Controls
    If TypeOf vari Is TextBox Or TypeOf vari Is ComboBox Then
        If vari.Tag = 1 Then
            vari.Locked = False
            vari.Enabled = True
        End If
    End If
Next
End Sub
Public Sub Deshabilitar(Formulario As Form)
Dim vari
For Each vari In Formulario.Controls
    If TypeOf vari Is TextBox Or TypeOf vari Is ComboBox Then
        If vari.Tag = 2 Then
            vari.Locked = False
        Else
            vari.Locked = True
        End If
    End If
Next
End Sub
Public Sub focus(obj As Object)
obj.SetFocus
End Sub
