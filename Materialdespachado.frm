VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form mdespachado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Despachado"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Materialdespachado.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6750
   Begin MSAdodcLib.Adodc bd2 
      Height          =   330
      Left            =   5160
      Top             =   4080
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
   Begin Jardin.xphelp xphelp1 
      Height          =   315
      Left            =   6120
      Top             =   3360
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin MSAdodcLib.Adodc bd1 
      Height          =   330
      Left            =   5040
      Top             =   4920
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
   Begin MSAdodcLib.Adodc bd 
      Height          =   330
      Left            =   3720
      Top             =   4920
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
   Begin Jardin.xpgroupbox frame 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7011
      Caption         =   "Material"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin VB.ComboBox ref 
         Height          =   315
         Left            =   2400
         TabIndex        =   28
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox sede 
         Height          =   315
         ItemData        =   "Materialdespachado.frx":058A
         Left            =   2400
         List            =   "Materialdespachado.frx":0594
         TabIndex        =   16
         Top             =   3000
         Width           =   1095
      End
      Begin VB.ComboBox numdoc 
         Height          =   315
         ItemData        =   "Materialdespachado.frx":059E
         Left            =   2400
         List            =   "Materialdespachado.frx":05A0
         TabIndex        =   15
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox numdes 
         Height          =   285
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox can 
         Height          =   285
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   1
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label fecdes 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2400
         TabIndex        =   30
         Top             =   720
         Width           =   60
      End
      Begin VB.Label nom 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2400
         TabIndex        =   27
         Top             =   1680
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Material"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   1410
      End
      Begin VB.Label saldo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2400
         TabIndex        =   24
         Top             =   3600
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Material"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sede"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento Empleado"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de despacho"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha despacho"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   915
      End
   End
   Begin Jardin.xpgroupbox xpgroupbox2 
      Height          =   2175
      Left            =   4920
      TabIndex        =   6
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   5530
      Caption         =   "Opciones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin JeweledBut.JeweledButton nuevo 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         TX              =   "&Nuevo"
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
         MICON           =   "Materialdespachado.frx":05A2
         BC              =   8438015
         FC              =   0
         Picture         =   "Materialdespachado.frx":0710
      End
      Begin JeweledBut.JeweledButton busqueda 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         TX              =   "&Buscar"
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
         MICON           =   "Materialdespachado.frx":321A
         BC              =   8438015
         FC              =   0
         Picture         =   "Materialdespachado.frx":3388
      End
      Begin JeweledBut.JeweledButton eliminar 
         Height          =   135
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   238
         TX              =   "Eli&minar"
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
         MICON           =   "Materialdespachado.frx":34E2
         BC              =   8438015
         FC              =   0
         Picture         =   "Materialdespachado.frx":3650
      End
      Begin JeweledBut.JeweledButton guardar 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         TX              =   "&Guardar"
         ENAB            =   0   'False
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
         MICON           =   "Materialdespachado.frx":3BEA
         BC              =   8438015
         FC              =   0
         Picture         =   "Materialdespachado.frx":3D58
      End
      Begin JeweledBut.JeweledButton cancelar 
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         TX              =   "Cancelar"
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
         MICON           =   "Materialdespachado.frx":9986
         BC              =   8438015
         FC              =   0
         Picture         =   "Materialdespachado.frx":9AF4
      End
   End
   Begin JeweledBut.JeweledButton salir 
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   3840
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
      MICON           =   "Materialdespachado.frx":9C4E
      BC              =   8438015
      FC              =   0
      Picture         =   "Materialdespachado.frx":9DBC
   End
   Begin Jardin.xpgroupbox xpgroupbox3 
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2143
      Caption         =   "Navegación"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin JeweledBut.JeweledButton primero 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Primer Registro"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         TX              =   "Primero"
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
         MICON           =   "Materialdespachado.frx":9F16
         BC              =   8438015
         FC              =   0
         Picture         =   "Materialdespachado.frx":A084
      End
      Begin JeweledBut.JeweledButton siguiente 
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         ToolTipText     =   "Siguiente Registro"
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         TX              =   "Siguiente"
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
         MICON           =   "Materialdespachado.frx":A1DE
         BC              =   8438015
         FC              =   0
         Picture         =   "Materialdespachado.frx":A34C
      End
      Begin JeweledBut.JeweledButton ultimo 
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         ToolTipText     =   "Ultimo Registro"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         TX              =   "Ultimo"
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
         MICON           =   "Materialdespachado.frx":A4A6
         BC              =   8438015
         FC              =   0
         Picture         =   "Materialdespachado.frx":A614
      End
      Begin JeweledBut.JeweledButton anterior 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Anterior Registro"
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         TX              =   "Anterior"
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
         MICON           =   "Materialdespachado.frx":A76E
         BC              =   8438015
         FC              =   0
         Picture         =   "Materialdespachado.frx":A8DC
      End
   End
   Begin VB.Label numreg 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5040
      TabIndex        =   25
      Top             =   2640
      Width           =   60
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Formato Fecha: dia/mes/año"
      Height          =   195
      Left            =   3720
      TabIndex        =   17
      Top             =   4560
      Width           =   2475
   End
End
Attribute VB_Name = "mdespachado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim nuevomd As Boolean


Private Sub anterior_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MovePrevious
    If bd.Recordset.BOF Then
        bd.Recordset.MoveFirst
    End If
    mostrarcampos
End If
numdes.SetFocus
End Sub

Private Sub bd_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub bd1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub bd2_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub busqueda_Click()
MB.Formulario = Me.Name
MB.Descripcion = "Material Didáctico Despachado"
elBuscador.Show
End Sub

Private Sub can_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    If ref.Text <> "" Then
        ConexionBD1 mdespachado, "SELECT * FROM material WHERE ref=" & ref.Text
        If BD1.Recordset.RecordCount > 0 Then
            If BD1.Recordset!can < Val(can.Text) Then
                MsgBox "No se puede despachar esta cantidad de material" & vbCrLf & "En inventario solo exiten: " & BD1.Recordset!can, vbInformation, "Material Despachado"
                Exit Sub
            End If
        End If
    End If
    numdoc.SetFocus
End If
End Sub
Private Sub desemp_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub

Private Sub desper_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    numdoc.SetFocus
End If
End Sub

Private Sub cancelar_Click()
If bd.Recordset.RecordCount > 0 Then
mostrarcampos
End If
If NuevoRegI = True Then
    nuevo.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    anterior.Enabled = True
    siguiente.Enabled = True
    ultimo.Enabled = True
    guardar.Enabled = False
    busqueda.Enabled = True
    NuevoRegI = False
End If
End Sub

Private Sub eliminar_Click()
If bd.Recordset.RecordCount > 0 Then
If MsgBox("Está seguro de querer eliminar el registro?", vbYesNo + vbQuestion, "Eliminar Registro") = vbYes Then
   bd.Recordset.Delete
   If bd.Recordset.RecordCount > 0 Then
    bd.Recordset.MoveFirst
    mostrarcampos
    Else
        Unload Me
        mdespachado.Show
   End If
End If
End If
End Sub


Private Sub LlenarCombos()
menu.estado.Panels(4).Text = "Cargando..."
'llenar sede
ConexionBD1 mdespachado, "select * from parametrizacion where tippar=22" & " order by dato;"
sede.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        sede.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar referencia
ConexionBD1 mdespachado, "select ref from material"
ref.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        ref.AddItem BD1.Recordset!ref
        BD1.Recordset.MoveNext
    Next i
End If
End Sub



Private Sub Form_Activate()
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height
FormularioActivo = True

numdes.Locked = True
ref.Locked = True
can.Locked = True
numdoc.Locked = True
sede.Locked = True
ConexionBD mdespachado, "select * from materialdespachado"
If bd.Recordset.RecordCount > 0 Then
    mostrarcampos
End If
ConexionBD2 mdespachado, "select numdoc from empleado"
numdoc.Clear
If bd2.Recordset.RecordCount > 0 Then
    bd2.Recordset.MoveFirst
    For i = 1 To bd2.Recordset.RecordCount
        numdoc.AddItem bd2.Recordset!numdoc
        bd2.Recordset.MoveNext
    Next i
End If
numdes.SetFocus

End Sub

Function mostrarcampos()
numreg = bd.Recordset.AbsolutePosition & " registro."
numdes.Text = bd.Recordset!numdes
fecdes = bd.Recordset!fecdes
ref.Text = bd.Recordset!ref
can.Text = bd.Recordset!can
numdoc.Text = bd.Recordset!numdoc
sede.Text = bd.Recordset!sede
ConexionBD1 mdespachado, "select * from material where ref=" & ref.Text
If BD1.Recordset.RecordCount > 0 Then
    saldo = BD1.Recordset!can
    nom = BD1.Recordset!nommat
End If
End Function
Function avanzar()
 If tecla = 13 Then
  SendKeys "{tab}"
  KeyAscii = 0: tecla = 0
 End If
End Function

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
LlenarCombos
menu.estado.Panels(4).Text = "Control de Material Didáctico Despachado"
End Sub

Private Sub Form_Resize()
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)
FormularioActivo = False
menu.estado.Panels(4).Text = "Menú Principal"
bd.Recordset.Close
Set bd.Recordset = Nothing
End Sub

Private Sub guardar_Click()
If MsgBox("Esta seguro(a) de guardar el despacho?" & vbCrLf & "Recuerde que una vez guardado no se puede eliminar!", vbYesNo + vbQuestion, "Material Didáctico Despachado") = vbYes Then
    guardarregistro
    ConexionBD1 mdespachado, "select can from material where ref=" & ref.Text
    If BD1.Recordset.RecordCount > 0 Then
        saldo = BD1.Recordset!can
    End If
    busqueda.Enabled = True
    nuevomd = False
End If
End Sub
Function guardarregistro()
If numdes.Text = "" Or ref.Text = "" Or can.Text = "" Or numdoc.Text = "" Or sede.Text = "" Then
    MsgBox "Falta datos por ingresar!", vbInformation, "Material Despachado"
    Exit Function
End If

'validar datos
'On Error Resume Next
ConexionBD1 mdespachado, "SELECT * FROM materialdespachado WHERE numdes=" & Val(numdes.Text) '& "';"


If BD1.Recordset.RecordCount = 1 Then
    MsgBox "Este número de despacho ya existe!", vbInformation, "Material Despachado"
    Exit Function
End If

BD1.Recordset.Close
'Si cantidad exece la de inventario entonces
ConexionBD1 mdespachado, "SELECT * FROM material WHERE ref=" & ref.Text
'bd1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=PROYECTO.mdb;Mode=Share Deny None;Persist Security Info=False;Jet OLEDB:Database Password=810205;Jet OLEDB:Engine Type=4"
'bd1.RecordSource = "SELECT * FROM material WHERE ref=" & ref.Text
'bd1.Refresh

If BD1.Recordset.RecordCount = 0 Then
    MsgBox "No existe material de esta referencia!", vbInformation, "Material Despachado"
    Exit Function
ElseIf BD1.Recordset.RecordCount > 0 Then
    If BD1.Recordset!can < Val(can.Text) Then
        MsgBox "No se puede despachar esta cantidad de material" & vbCrLf & "En inventario solo exiten: " & BD1.Recordset!can, vbInformation, "Material Despachado"
        Exit Function
    End If
End If

'ConexionBD1 mdespachado, "SELECT * FROM empleado WHERE numdoc=" & numdoc.Text

'If bd1.Recordset.RecordCount = 0 Then
 '   MsgBox "Este número de documento no existe!", vbInformation, "Material Despachado"
    'Exit Function
'End If



On Error GoTo ERROR_GUARDAR
bd.Recordset.AddNew
bd.Recordset!numdes = Val(numdes.Text)
bd.Recordset!fecdes = fecdes
bd.Recordset!ref = ref.Text
bd.Recordset!can = Val(can.Text)
'bd.Recordset!desper = desper.Text
bd.Recordset!numdoc = Val(numdoc.Text)
bd.Recordset!sede = sede.Text
bd.Recordset.Update

'descuenta del material existente
ConexionBD1 mdespachado, "SELECT * FROM material WHERE ref=" & ref.Text
Dim editar

editar = BD1.Recordset.EditMode
BD1.Recordset!can = Val(BD1.Recordset!can) - Val(can.Text)
BD1.Recordset.Update


nuevo.Enabled = True

eliminar.Enabled = True
primero.Enabled = True
anterior.Enabled = True
siguiente.Enabled = True
ultimo.Enabled = True
guardar.Enabled = False

numdes.Locked = True
ref.Locked = True
can.Locked = True

SALIR_GUARDAR:
    Exit Function
ERROR_GUARDAR:
    MsgBox Err.Description, vbCritical, "Error"
    Resume SALIR_GUARDAR

End Function



Private Sub nuevo_Click()
'deshabilitar controles
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False
nuevo.Enabled = False
guardar.Enabled = True

eliminar.Enabled = False
busqueda.Enabled = False
numdes.Locked = False
ref.Locked = False
can.Locked = False
'desper.Locked = False
numdoc.Locked = False
sede.Locked = False
nuevomd = True

ref = ""
can = ""
'desper.Text = ""
numdoc.Text = ""
sede.Text = ""
'genera el autonumerico para numero de despacho
ConexionBD1 mdespachado, "select * from materialdespachado"
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveLast
    numdes.Text = BD1.Recordset!numdes + 1
ElseIf BD1.Recordset.RecordCount = 0 Then
    numdes.Text = 1
End If
busqueda.Enabled = False
ref.SetFocus
saldo = ""
nom = ""
fecdes = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub numdes_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    ref.SetFocus
End If
End Sub

Private Sub numdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    sede.SetFocus
End If
End Sub

Private Sub primero_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveFirst
    mostrarcampos
End If
numdes.SetFocus
End Sub

Private Sub ref_Click()
On Error Resume Next
ConexionBD1 mdespachado, "select * from material where ref=" & Val(ref.Text)
If BD1.Recordset.RecordCount > 0 Then
    nom = BD1.Recordset!nommat
    saldo = BD1.Recordset!can
End If
End Sub

Private Sub ref_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    If ref.Text <> "" Then
        ConexionBD1 mdespachado, "SELECT * FROM material WHERE ref=" & ref.Text
    End If
    If BD1.Recordset.RecordCount = 0 Then
        MsgBox "No existe material de esta referencia!", vbInformation, "Material Despachado"
        Exit Sub
    Else
        can.SetFocus
    End If
    ConexionBD1 mdespachado, "select * from material where ref=" & Val(ref.Text)
    If BD1.Recordset.RecordCount > 0 Then
        nom = BD1.Recordset!nommat
        saldo = BD1.Recordset!can
    End If
End If
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Private Sub sede_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
End Sub

Private Sub siguiente_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveNext
    If bd.Recordset.EOF Then
        bd.Recordset.MoveLast
    End If
    mostrarcampos
End If
numdes.SetFocus
End Sub

Private Sub ultimo_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveLast
    mostrarcampos
End If
numdes.SetFocus
End Sub

Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "material.htm"
End Sub
