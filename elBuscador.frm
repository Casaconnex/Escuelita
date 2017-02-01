VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form elBuscador 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "elBuscador de SISJACE"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5835
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
   ScaleHeight     =   4185
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc bd2 
      Height          =   330
      Left            =   4200
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc docu 
      Height          =   330
      Left            =   2640
      Top             =   3720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc BD1 
      Height          =   330
      Left            =   2640
      Top             =   3360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin JeweledBut.JeweledButton buscar 
      Default         =   -1  'True
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      TX              =   "Buscar..."
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
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "elBuscador.frx":0000
      BC              =   8438015
      FC              =   0
      Picture         =   "elBuscador.frx":001C
   End
   Begin VB.TextBox termino 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.ListBox lista 
      Height          =   1425
      ItemData        =   "elBuscador.frx":0176
      Left            =   120
      List            =   "elBuscador.frx":0178
      TabIndex        =   1
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Ingrese el termino que desea buscar:"
      Height          =   435
      Left            =   2640
      TabIndex        =   7
      Top             =   2280
      Width           =   3000
   End
   Begin VB.Label numero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   60
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Buscar en:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   930
   End
   Begin VB.Label buscaren 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   1920
      Width           =   60
   End
   Begin VB.Label Label1 
      Caption         =   "Para buscar por favor seleccione un parametro en la lista, digite el termino y luego de click en el botón etiquetado ""Buscar""."
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4035
   End
   Begin VB.Image Image3 
      Height          =   1845
      Left            =   3840
      Picture         =   "elBuscador.frx":017A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   120
      Picture         =   "elBuscador.frx":3004
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "elBuscador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buscar_Click()
If termino.Text = "" Then
    MsgBox "No ha ingresado el termino a buscar!", vbExclamation, "elBuscador"
    termino.SetFocus
    Exit Sub
End If
Select Case lista.Tag
    'listado de espera
    Case 1:
        Select Case lista.ListIndex
            Case 0:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from listadodeespera where numdoc='" & termino.Text & "'"
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposL
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            Case 1:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from listadodeespera where priape='" & termino.Text & "'"
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposL
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
        End Select
    'inscripciones
    Case 2:
        Select Case lista.ListIndex
            Case 0:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from inscripciones where numdoc='" & termino.Text & "'"
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposI
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            Case 1:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from inscripciones where numins=" & termino.Text
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposI
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
        End Select
    'matriculas
    Case 3:
        Select Case lista.ListIndex
            Case 0:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from matricula where numdoc='" & termino.Text & "'"
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposM
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            Case 1:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from matricula where numfor=" & termino.Text
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposM
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
        End Select
    'empleado
    Case 4:
        Select Case lista.ListIndex
            Case 0:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from empleado where numdoc=" & termino.Text
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposE
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            Case 1:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from empleado where nom='" & termino.Text & "'"
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposE
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            Case 2:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from empleado where cor='" & termino.Text & "'"
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposE
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            Case 3:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from empleado where car='" & termino.Text & "'"
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposE
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
        End Select
    'pagos
    Case 5:
        Select Case lista.ListIndex
            Case 0:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from pagos where numdoc='" & termino.Text & "'"
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposP
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            Case 1:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from pagos where tipopago='" & termino.Text & "'"
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposP
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            Case 2:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from pagos where numcon=" & termino.Text
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposP
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
        End Select
    'material existente
    Case 6:
        Select Case lista.ListIndex
            Case 0:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from material where ref=" & termino.Text
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposMae
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            Case 1:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from material where nommat='" & termino.Text & "'"
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposMae
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
        End Select
    'material despachado
    Case 7:
        Select Case lista.ListIndex
            Case 0:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from materialdespachado where numdes=" & termino.Text
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposMad
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            Case 1:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from materialdespachado where numdoc=" & termino.Text
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposMad
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            Case 2:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from materialdespachado where sede='" & termino.Text & "'"
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposMad
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            End Select
        'compras
    Case 8:
        Select Case lista.ListIndex
            Case 0:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from compras where numcom=" & termino.Text
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposCom
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
            Case 1:
                On Error Resume Next
                ConexionBD1 elBuscador, "select * from compras where refmat=" & termino.Text
                If bd1.Recordset.RecordCount > 0 Then
                    numero = "Encontrados: " & bd1.Recordset.RecordCount & " coincidencias."
                    MostrarCamposCom
                    Unload Me
                Else
                    MsgBox "No se encontró ninguna coincidencia!", vbInformation, "elBuscador"
                    Exit Sub
                End If
    End Select
End Select
End Sub
Function MostrarCamposCom()
Compras.numcom = bd1.Recordset!numcom
Compras.numfac.Text = bd1.Recordset!numfac
Compras.fecfac.Value = bd1.Recordset!fecfac
Compras.can.Text = bd1.Recordset!can
Compras.refmat.Text = bd1.Recordset!refmat
Compras.valor.Text = bd1.Recordset!valor
Compras.med.Visible = True
Compras.nmm.Visible = False
End Function
Function MostrarCamposMad()
mdespachado.numreg = bd1.Recordset.AbsolutePosition & " registro."
mdespachado.numdes.Text = bd1.Recordset!numdes
mdespachado.fecdes = bd1.Recordset!fecdes
mdespachado.ref.Text = bd1.Recordset!ref
mdespachado.can.Text = bd1.Recordset!can
mdespachado.numdoc.Text = bd1.Recordset!numdoc
mdespachado.sede.Text = bd1.Recordset!sede
ConexionBD2 elBuscador, "select can from material where ref=" & bd1.Recordset!ref
If bd2.Recordset.RecordCount > 0 Then
    mdespachado.saldo = bd2.Recordset!can
End If
End Function
Function MostrarCamposMae()
material.numreg = bd1.Recordset.AbsolutePosition & " registro."
material.ref.Text = bd1.Recordset!ref
material.nommat.Text = bd1.Recordset!nommat
material.can.Text = bd1.Recordset!can
End Function
Function MostrarCamposP()
Pagos.numreg = bd1.Recordset.AbsolutePosition & " registro."
Pagos.numdoc.Text = bd1.Recordset!numdoc
Pagos.pagmat.Text = bd1.Recordset!tipopago
If IsNull(bd1.Recordset!otro) = False Then
    Pagos.otro.Text = bd1.Recordset!otro
End If
If IsNull(bd1.Recordset!mes) = False Then
    Pagos.mes.Text = bd1.Recordset!mes
End If
Pagos.valcon.Text = bd1.Recordset!valcon
Pagos.feccon.Value = bd1.Recordset!feccon
Pagos.numcon.Text = bd1.Recordset!numcon
End Function
Function MostrarCamposE()
empleado.numreg = bd1.Recordset.AbsolutePosition & " registro."
empleado.numdoc.Text = bd1.Recordset!numdoc
empleado.exp.Text = bd1.Recordset!exp
empleado.nom.Text = bd1.Recordset!nom
empleado.ape.Text = bd1.Recordset!ape
empleado.dir.Text = bd1.Recordset!dir
empleado.bar.Text = bd1.Recordset!bar
If IsNull(bd1.Recordset!tel) = False Then
    empleado.tel.Text = bd1.Recordset!tel
End If
If IsNull(bd1.Recordset!cor) = False Then
    empleado.cor.Text = bd1.Recordset!cor
End If
If IsNull(bd1.Recordset!cel) = False Then
    empleado.cel.Text = bd1.Recordset!cel
End If
If IsNull(bd1.Recordset!pro) = False Then
    empleado.pro.Text = bd1.Recordset!pro
End If
If IsNull(bd1.Recordset!car) = False Then
    empleado.car.Text = bd1.Recordset!car
End If
If IsNull(bd1.Recordset!fecvin) = False Then
    empleado.fecvin.Value = bd1.Recordset!fecvin
End If
If IsNull(bd1.Recordset!fecnac) = False Then
    empleado.fecnac.Value = bd1.Recordset!fecnac
End If
If IsNull(bd1.Recordset!nivestfor) = False Then
    empleado.nivestfor.Text = bd1.Recordset!nivestfor
End If
If IsNull(bd1.Recordset!niveestnofor) = False Then
    empleado.nivestnofor.Text = bd1.Recordset!niveestnofor
End If
End Function
Function MostrarCamposM()
'muestra todos los campos de matricula
    matricula.numreg = bd1.Recordset.AbsolutePosition & " registro."
    matricula.numfor = bd1.Recordset!numfor
    matricula.col.Text = bd1.Recordset!col
    matricula.uniope.Text = bd1.Recordset!uniope
    matricula.modal.Text = bd1.Recordset!modal
    matricula.submod.Text = bd1.Recordset!submod
    matricula.fecmat = bd1.Recordset!fecmat
    matricula.persolser.Text = bd1.Recordset!persolser
    matricula.rempor.Text = bd1.Recordset!rempor
    matricula.prorem.Text = bd1.Recordset!prorem
    matricula.entrem.Text = bd1.Recordset!entrem
    matricula.numdoc.Text = bd1.Recordset!numdoc
    matricula.depnac.Text = bd1.Recordset!depnac
    matricula.munnac.Text = bd1.Recordset!munnac
    matricula.Painac.Text = bd1.Recordset!Painac
    matricula.tipdisest.Text = bd1.Recordset!tipdisest
    matricula.nivestalc.Text = bd1.Recordset!niveduben
    matricula.asiactcenedu.Text = bd1.Recordset!asiactcenedu
    matricula.proaso.Text = bd1.Recordset!proaso
    matricula.afisegsocfam.Text = bd1.Recordset!afisegsocben
    matricula.regsegsocfam.Text = bd1.Recordset!regsegsocben
    matricula.calbenfam.Text = bd1.Recordset!calben
    matricula.vinsecsalfam.Text = bd1.Recordset!vinsecsalben
    numficsis.Text = bd1.Recordset!numficsis
    punsis.Text = bd1.Recordset!punsis
    matricula.loc.Text = bd1.Recordset!loc
    matricula.forpagviv.Text = bd1.Recordset!forpagviv
    matricula.dep.Text = bd1.Recordset!dptoprofam
    matricula.mun.Text = bd1.Recordset!munprofam
    matricula.pais.Text = bd1.Recordset!paiprofam
    matricula.feclle.Value = bd1.Recordset!fecllebogfam
    matricula.ninvivpapmam.Text = bd1.Recordset!ninvivpapmam
    matricula.ninvivperpadmadotr.Text = bd1.Recordset!ninvivperpadmadotr
    matricula.vivpermpadmad.Text = bd1.Recordset!vivperpadmad
    matricula.edaninvivpapmad.Text = bd1.Recordset!edaninvivpapmad
    matricula.cuinindurdia.Text = bd1.Recordset!cuinindurdia
    matricula.graasp.Text = bd1.Recordset!graasp
    If IsNull(bd1.Recordset!nomdilform) = False Then
        matricula.nomdilfor.Text = bd1.Recordset!nomdilform
    End If
    matricula.nomfundighojsir.Text = bd1.Recordset!nomfundighojsir
    matricula.fecdighojsir.Value = bd1.Recordset!fecdighojsir
    matricula.obs.Text = bd1.Recordset!obs
    matricula.SSTab1.Tab = 0
    'muestra los campos de listado de espera
    ConexionBD2 elBuscador, "select * from listadodeespera where numdoc='" & bd1.Recordset!numdoc & "'"
    If bd2.Recordset.RecordCount > 0 Then
        matricula.tipdocnin.Text = bd2.Recordset!tipdoc
        matricula.nomnin.Text = bd2.Recordset!prinom & " " & bd2.Recordset!segnom
        matricula.apenin.Text = bd2.Recordset!priape & " " & bd2.Recordset!segape
        matricula.sexo.Text = bd2.Recordset!sex
        matricula.fecnac.Value = bd2.Recordset!fecnac
        matricula.edad.Text = bd2.Recordset!eda
        matricula.parjeffam.Text = bd2.Recordset!parfam
        matricula.dir.Text = bd2.Recordset!dir
        matricula.bar.Text = bd2.Recordset!bar
        matricula.tel.Text = bd2.Recordset!tel
    End If
        'conectamos bd para cargar datos de inscripciones
        ConexionBD2 matricula, "select * from inscripciones where numdoc='" & bd1.Recordset!numdoc & "'"
        If bd2.Recordset.RecordCount > 0 Then
            matricula.tipviv.Text = bd2.Recordset!tipviv
            matricula.conviv.Text = bd2.Recordset!conviv
            matricula.tenviv.Text = bd2.Recordset!tenviv
        End If
End Function
Function MostrarCamposI()
inscripciones.numreg = bd1.Recordset.AbsolutePosition & " registro."
'mostrar listado de documentos en listado de espera
inscripciones.numdoc.Clear
ConexionDocu elBuscador, "select * from listadodeespera"
If docu.Recordset.RecordCount > 0 Then
    docu.Recordset.MoveFirst
    For i = 1 To docu.Recordset.RecordCount
        inscripciones.numdoc.AddItem Trim$(docu.Recordset!numdoc)
        docu.Recordset.MoveNext
    Next i
End If
inscripciones.numdoc.Text = bd1.Recordset!numdoc
inscripciones.lugnac = bd1.Recordset!lugnac

If bd1.Recordset!prealgenf = "NO" Then
    inscripciones.xpradiobutton2.Value = True
    inscripciones.xpradiobutton1.Value = False
    inscripciones.prealgenf.Visible = False
ElseIf bd1.Recordset!prealgenf <> "SI" Then
    inscripciones.prealgenf = bd1.Recordset!prealgenf
    inscripciones.prealgenf.Visible = True
    inscripciones.xpradiobutton1.Value = True
    inscripciones.xpradiobutton2.Value = False
End If
If IsNull(bd1.Recordset!numins) = False Then
    inscripciones.numins = bd1.Recordset!numins
End If
inscripciones.numher = bd1.Recordset!numher
inscripciones.lugocufam = bd1.Recordset!lugocufam
inscripciones.ninviv = bd1.Recordset!ninviv

inscripciones.nompad.Text = bd1.Recordset!nompad
inscripciones.ocupad.Text = bd1.Recordset!ocupad
inscripciones.ingmenpad.Text = bd1.Recordset!ingmenpad
inscripciones.edapad.Text = bd1.Recordset!edapad
inscripciones.nomemppad.Text = bd1.Recordset!nomemppad
inscripciones.telemppad.Text = bd1.Recordset!telemppad
inscripciones.nivacapad.Text = bd1.Recordset!nivacapad
inscripciones.otringpad.Text = bd1.Recordset!otringpad
inscripciones.nommad.Text = bd1.Recordset!nommad
inscripciones.ocumad.Text = bd1.Recordset!ocumad
inscripciones.nomempmad.Text = bd1.Recordset!nomempmad
inscripciones.ingmenmad.Text = bd1.Recordset!ingmenmad
inscripciones.edamad.Text = bd1.Recordset!edamad
If bd1.Recordset!telempmad <> Null Then
    inscripciones.telempmad.Text = bd1.Recordset!telempmad
End If
If bd1.Recordset!nivacamad <> Null Then
    inscripciones.nivacamad.Text = bd1.Recordset!nivacamad
End If
If bd1.Recordset!otringmad <> Null Then
    inscripciones.otringmad.Text = bd1.Recordset!otringmad
End If
If bd1.Recordset!tenviv <> Null Then
    inscripciones.tenviv.Text = bd1.Recordset!tenviv
End If
If bd1.Recordset!tipviv <> Null Then
    inscripciones.tipviv.Text = bd1.Recordset!tipviv
End If
If bd1.Recordset!conviv <> Null Then
    inscripciones.conviv.Text = bd1.Recordset!conviv
End If
If bd1.Recordset!estviv <> Null Then
    inscripciones.estviv.Text = bd1.Recordset!estviv
End If
If bd1.Recordset!serpub <> Null Then
    inscripciones.serpub.Text = bd1.Recordset!serpub
End If
inscripciones.SSTab1.Tab = 0
inscripciones.numdoc.SetFocus
End Function
Function MostrarCamposL()
listespera.numreg = bd1.Recordset.AbsolutePosition & " registro."
listespera.fecins = bd1.Recordset!fecins
listespera.tipdoc = bd1.Recordset!tipdoc
listespera.numdoc = bd1.Recordset!numdoc
listespera.priape = bd1.Recordset!priape
listespera.segape = bd1.Recordset!segape
listespera.prinom = bd1.Recordset!prinom
listespera.segnom = bd1.Recordset!segnom
listespera.sex = bd1.Recordset!sex
If IsNull(bd1.Recordset!fecnac) = False Then
    listespera.fecnac = bd1.Recordset!fecnac
    listespera.fecnac.Enabled = True
Else
    listespera.fecnac.Enabled = False
End If
listespera.eda = bd1.Recordset!eda
If IsNull(bd1.Recordset!edaano) = False Then
    listespera.edaano.Text = bd1.Recordset!edaano
End If
If IsNull(bd1.Recordset!edames) = False Then
    listespera.edames.Text = bd1.Recordset!edames
End If
listespera.parfam = bd1.Recordset!parfam
listespera.dir = bd1.Recordset!dir
listespera.bar = bd1.Recordset!bar
listespera.tel = bd1.Recordset!tel
listespera.sal = bd1.Recordset!sal
listespera.sec = bd1.Recordset!secrsal
listespera.nivel = bd1.Recordset!niv
End Function

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
Select Case MB.Formulario
    Case "listespera":
        'muestra descripcion busqueda
        buscaren.Caption = MB.Descripcion
        'genera la lista de parametros
        lista.Clear
        lista.AddItem "Num. Documento"
        lista.AddItem "Primer Apellido"
        'controla cajatexto y cajafecha
        lista.Tag = 1
    Case "inscripciones":
        'muestra descripcion busqueda
        buscaren.Caption = MB.Descripcion
        'genera la lista de parametros
        lista.Clear
        lista.AddItem "Num. Documento"
        lista.AddItem "Num. Inscripción"
        'controla cajatexto y cajafecha
        lista.Tag = 2
    Case "matricula":
        'muestra descripcion busqueda
        buscaren.Caption = MB.Descripcion
        'genera la lista de parametros
        lista.Clear
        lista.AddItem "Num. Documento"
        lista.AddItem "Num. Matricula"
        'controla cajatexto y cajafecha
        lista.Tag = 3
    Case "empleado":
        'muestra descripcion busqueda
        buscaren.Caption = MB.Descripcion
        'genera la lista de parametros
        lista.Clear
        lista.AddItem "Num. Documento"
        lista.AddItem "Nombres"
        lista.AddItem "Correo Electrónico"
        lista.AddItem "Cargo"
        'controla cajatexto y cajafecha
        lista.Tag = 4
    Case "Pagos":
        'muestra descripcion busqueda
        buscaren.Caption = MB.Descripcion
        'genera la lista de parametros
        lista.Clear
        lista.AddItem "Num. Documento"
        lista.AddItem "Concepto Pago"
        lista.AddItem "Num. Consignación"
        'controla cajatexto y cajafecha
        lista.Tag = 5
    Case "material":
        'muestra descripcion busqueda
        buscaren.Caption = MB.Descripcion
        'genera la lista de parametros
        lista.Clear
        lista.AddItem "Referencia"
        lista.AddItem "Nombre Material"
        'controla cajatexto y cajafecha
        lista.Tag = 6
    Case "mdespachado":
        'muestra descripcion busqueda
        buscaren.Caption = MB.Descripcion
        'genera la lista de parametros
        lista.Clear
        lista.AddItem "Num. Despacho"
        lista.AddItem "Num. Doc. Empleado"
        lista.AddItem "Sede"
        'controla cajatexto y cajafecha
        lista.Tag = 7
    Case "Compras":
        buscaren.Caption = MB.Descripcion
        'genera la lista de parametros
        lista.Clear
        lista.AddItem "Num. Compra"
        lista.AddItem "Num. Referencia Material"
        'controla cajatexto y cajafecha
        lista.Tag = 8
End Select
End Sub

Private Sub lista_Click()
termino.SetFocus
End Sub

Private Sub termino_GotFocus()
termino.SelStart = 0
termino.SelLength = Len(termino.Text)
End Sub
