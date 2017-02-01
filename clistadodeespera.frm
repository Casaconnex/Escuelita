VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form consultas 
   Caption         =   "Consultas"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "clistadodeespera.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin Jardin.xphelp xphelp1 
      Height          =   315
      Left            =   9960
      Top             =   840
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   120
   End
   Begin MSDataGridLib.DataGrid consulta 
      Bindings        =   "clistadodeespera.frx":0E42
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12726
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc bd 
      Height          =   330
      Left            =   360
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin JeweledBut.JeweledButton salir 
      Height          =   375
      Left            =   10080
      TabIndex        =   1
      Top             =   7320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Cerrar"
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
      MPTR            =   99
      MICON           =   "clistadodeespera.frx":0E53
      BC              =   8438015
      FC              =   0
      Picture         =   "clistadodeespera.frx":0FC1
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C5FAFE&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   10080
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   1575
      Left            =   10320
      Picture         =   "clistadodeespera.frx":111B
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C5FAFE&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   9960
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "consultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bd_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub Form_Activate()
FormularioActivo = True
salir.SetFocus
End Sub

Private Sub Form_Load()
If Tabla = "listadodeespera" Then
    menu.estado.Panels(4).Text = "Listado de Espera"
    On Error Resume Next
    ConexionBD consultas, "select * from " & Tabla
    consulta.Caption = "LISTADO DE NIÑOS EN ESPERA DE CUPO"
    consulta.Columns(0).Caption = "Fecha Incripción"
    consulta.Columns(1).Caption = "Tipo Documento"
    consulta.Columns(2).Caption = "No. Documento"
    consulta.Columns(3).Caption = "Primer Apellido"
    consulta.Columns(4).Caption = "Segundo Apellido"
    consulta.Columns(5).Caption = "Primer Nombre"
    consulta.Columns(6).Caption = "Segundo Nombre"
    consulta.Columns(7).Caption = "Fecha Nacimiento"
    consulta.Columns(8).Caption = "Edad"
    consulta.Columns(9).Caption = "Sexo"
    consulta.Columns(10).Caption = "Parentesco Familiar"
    consulta.Columns(11).Caption = "Dirección"
    consulta.Columns(12).Caption = "Barrio"
    consulta.Columns(13).Caption = "Telefono"
    consulta.Columns(14).Caption = "Tipo Salud"
    consulta.Columns(15).Caption = "Nivel Salud"
    consulta.Columns(16).Caption = "Afiliado Sec. Salud"
    consulta.Enabled = False
ElseIf Tabla = "inscripciones" Then
    menu.estado.Panels(4).Text = "Listado de Incripciones"
    On Error Resume Next
    ConexionBD consultas, "select * from " & Tabla
    consulta.Caption = "LISTADO DE NIÑOS INSCRITOS"
    consulta.Columns(0).Caption = "No. Documento"
    consulta.Columns(1).Caption = "Lugar Nacimiento"
    consulta.Columns(2).Caption = "Enfermedad"
    consulta.Columns(3).Caption = "No. Hermanos"
    consulta.Columns(4).Caption = "Lugar en la Familia"
    consulta.Columns(5).Caption = "Vive con"
    consulta.Columns(6).Caption = "Nombre Padre"
    consulta.Columns(7).Caption = "Ocupación"
    consulta.Columns(8).Caption = "Ingresos"
    consulta.Columns(9).Caption = "Edad"
    consulta.Columns(10).Caption = "Empresa"
    consulta.Columns(11).Caption = "Telefono"
    consulta.Columns(12).Caption = "Nivel Académico"
    consulta.Columns(13).Caption = "Otros Ingresos"
    consulta.Columns(14).Caption = "Nombre Madre"
    consulta.Columns(15).Caption = "Ocupación"
    consulta.Columns(16).Caption = "Empresa"
    consulta.Columns(17).Caption = "Ingresos"
    consulta.Columns(18).Caption = "Edad"
    consulta.Columns(19).Caption = "Telefono"
    consulta.Columns(20).Caption = "Nivel Académico"
    consulta.Columns(21).Caption = "Otros Ingresos"
    consulta.Columns(22).Caption = "Tenencia Vivienda"
    consulta.Columns(23).Caption = "Tipo"
    consulta.Columns(24).Caption = "Condición"
    consulta.Columns(25).Caption = "Estado"
    consulta.Columns(26).Caption = "Servicios Públicos"
    consulta.Enabled = False
ElseIf Tabla = "matricula" Then
    menu.estado.Panels(4).Text = "Listado de Matriculas"
    On Error Resume Next
    ConexionBD consultas, "select * from " & Tabla
    consulta.Caption = "LISTADO DE NIÑOS MATRICULADOS"
    consulta.Columns(0).Caption = "No. Documento"
    consulta.Columns(1).Caption = "Fecha Matricula"
    consulta.Columns(2).Caption = "Nombre Institución"
    consulta.Columns(3).Caption = "Lugar Institución"
    consulta.Columns(4).Caption = "Tipo Modalidad"
    consulta.Columns(5).Caption = "Tipo Submodalidad"
    consulta.Columns(6).Caption = "No. Formulario"
    consulta.Columns(7).Caption = "Persona que Solicita Servicio"
    consulta.Columns(8).Caption = "Entidad Reporta Beneficiario"
    consulta.Columns(9).Caption = "Entidad Remite"
    consulta.Columns(10).Caption = "Proyecto Remite"
    consulta.Columns(11).Caption = "Depto. Naciemiento"
    consulta.Columns(12).Caption = "Municipio"
    consulta.Columns(13).Caption = "País"
    consulta.Columns(14).Caption = "Discapacidad Beneficiario"
    consulta.Columns(15).Caption = "Tipo Discapacidad"
    consulta.Columns(16).Caption = "Nivel Educación"
    consulta.Columns(17).Caption = "Asistencia Plantel Educativo"
    consulta.Columns(18).Caption = "Problema Asociados"
    consulta.Columns(19).Caption = "Seguridad Social"
    consulta.Columns(20).Caption = "Regimen Seguridad Social"
    consulta.Columns(21).Caption = "Calidad"
    consulta.Columns(22).Caption = "Vinculación Sec. Salud"
    consulta.Columns(23).Caption = "No. Ficha Sisben"
    consulta.Columns(24).Caption = "Puntaje Sisben"
    consulta.Columns(25).Caption = "Localidad Vivienda"
    consulta.Columns(26).Caption = "Estrato"
    consulta.Columns(27).Caption = "Forma Pago Vivienda"
    consulta.Columns(28).Caption = "Depto. Jefe Hogar"
    consulta.Columns(29).Caption = "Municipio"
    consulta.Columns(30).Caption = "País"
    consulta.Columns(31).Caption = "Fecha Llegada a Bogotá"
    consulta.Columns(32).Caption = "Niño vive con"
    consulta.Columns(33).Caption = "Beneficiario vive con"
    consulta.Columns(34).Caption = "Edad Beneficiario"
    consulta.Columns(35).Caption = "Cuidado durante el día"
    consulta.Columns(36).Caption = "Grado Aspiración"
    consulta.Columns(37).Caption = "No. Doc. Jefe Hogar"
    consulta.Columns(38).Caption = "Tipo Documento"
    consulta.Columns(39).Caption = "Sexo Jefe Hogar"
    consulta.Columns(40).Caption = "Fecha Nacimiento"
    consulta.Columns(41).Caption = "Estado Civil"
    consulta.Columns(42).Caption = "Tipo Discapacidad"
    consulta.Columns(43).Caption = "Parentesco"
    consulta.Columns(44).Caption = "Nivel Estudio"
    consulta.Columns(45).Caption = "Años de Educación"
    consulta.Columns(46).Caption = "Asiste Centro Educativo"
    consulta.Columns(47).Caption = "Ocupación"
    consulta.Columns(48).Caption = "Posición Ocupacional"
    consulta.Columns(49).Caption = "Forma Ingresos"
    consulta.Columns(50).Caption = "Afiliado Seguridad Social"
    consulta.Columns(51).Caption = "Regimen Seguridad Social"
    consulta.Columns(52).Caption = "Calidad Entidad Jefe Hogar"
    consulta.Columns(53).Caption = "Vinculación Sec. Salud Distrital"
    consulta.Columns(54).Caption = "Violencia Intrafamliar"
    consulta.Columns(55).Caption = "Nombre Funcionario"
    consulta.Columns(56).Caption = "Fecha Digitación Hoja Sirbe"
    consulta.Columns(57).Caption = "Nombre Funcionario que digito"
    consulta.Columns(58).Caption = "Edad del Familiar"
    consulta.Columns(59).Caption = "Observaciones Generales"
    consulta.Enabled = False
ElseIf Tabla = "Empleado" Then
    menu.estado.Panels(4).Text = "Listado de Empleados"
    On Error Resume Next
    ConexionBD consultas, "select * from " & Tabla
    consulta.Caption = "LISTADO DE EMPLEADOS DEL JARDIN"
    consulta.Columns(0).Caption = "No. Documento"
    consulta.Columns(1).Caption = "Lugar Expedición Doc."
    consulta.Columns(2).Caption = "Nombre Empleado"
    consulta.Columns(3).Caption = "Apellido"
    consulta.Columns(4).Caption = "Dirección"
    consulta.Columns(5).Caption = "Barrio"
    consulta.Columns(6).Caption = "Telefono"
    consulta.Columns(7).Caption = "E-mail"
    consulta.Columns(8).Caption = "Celular"
    consulta.Columns(9).Caption = "Profesión"
    consulta.Columns(10).Caption = "Cargo"
    consulta.Columns(11).Caption = "Fecha Vinculación"
    consulta.Columns(12).Caption = "Fecha Nacimiento"
    consulta.Columns(13).Caption = "Nivel Educación Formal"
    consulta.Columns(14).Caption = "Nivel Educación No Formal"
    consulta.Enabled = False
ElseIf Tabla = "Pagos" Then
    menu.estado.Panels(4).Text = "Listado de Pagos"
    On Error Resume Next
    ConexionBD consultas, "select * from " & Tabla
    consulta.Caption = "LISTADO DE LOS PAGOS REALIZADOS"
    consulta.Columns(0).Caption = "No. Documento"
    consulta.Columns(1).Caption = "Pago de"
    consulta.Columns(2).Caption = "Concepto"
    consulta.Columns(3).Caption = "Valor Consignación"
    consulta.Columns(4).Caption = "Fecha Consignación"
    consulta.Columns(5).Caption = "No. Consignación"
    consulta.Enabled = False
ElseIf Tabla = "Material" Then
    menu.estado.Panels(4).Text = "Listado de Material Didáctico"
    On Error Resume Next
    ConexionBD consultas, "select * from " & Tabla
    consulta.Caption = "LISTADO DEL MATERIAL DIDACTICO"
    consulta.Columns(0).Caption = "Referencia"
    consulta.Columns(1).Caption = "Nombre"
    consulta.Columns(2).Caption = "Cantidad"
    consulta.Enabled = False
ElseIf Tabla = "Materialdespachado" Then
    menu.estado.Panels(4).Text = "Listado de Material Didáctico Despachado"
    On Error Resume Next
    ConexionBD consultas, "select * from " & Tabla
    consulta.Caption = "LISTADO DE MATERIAL DESPACHADO"
    consulta.Columns(0).Caption = "No. Despacho"
    consulta.Columns(1).Caption = "Referencia"
    consulta.Columns(2).Caption = "Cantidad"
    consulta.Columns(3).Caption = "Despachado Personal"
    consulta.Enabled = False
End If

If Err.Number <> 0 Then
    MsgBox Err.Description, vbCritical, "Error en Consultas"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormularioActivo = False
menu.estado.Panels(4).Text = "Menú Principal"
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
consulta.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "consultas.htm"
End Sub
