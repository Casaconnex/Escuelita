VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form material 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Didáctico Existente"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "material didactico.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   8250
   Begin MSAdodcLib.Adodc bd1 
      Height          =   330
      Left            =   3720
      Top             =   3000
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
      Left            =   5040
      Top             =   3000
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
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3625
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
      Begin Jardin.xphelp xphelp1 
         Height          =   315
         Left            =   5760
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin VB.TextBox ref 
         Height          =   285
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox nommat 
         Height          =   285
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox can 
         Height          =   285
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   5160
         Picture         =   "material didactico.frx":058A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia del material"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del material"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad del material"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1845
      End
   End
   Begin Jardin.xpgroupbox xpgroupbox1 
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   3615
      _ExtentX        =   6376
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
         TabIndex        =   8
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
         MICON           =   "material didactico.frx":1538
         BC              =   8438015
         FC              =   0
         Picture         =   "material didactico.frx":16A6
      End
      Begin JeweledBut.JeweledButton siguiente 
         Height          =   375
         Left            =   1800
         TabIndex        =   9
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
         MICON           =   "material didactico.frx":1800
         BC              =   8438015
         FC              =   0
         Picture         =   "material didactico.frx":196E
      End
      Begin JeweledBut.JeweledButton ultimo 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
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
         MICON           =   "material didactico.frx":1AC8
         BC              =   8438015
         FC              =   0
         Picture         =   "material didactico.frx":1C36
      End
      Begin JeweledBut.JeweledButton anterior 
         Height          =   375
         Left            =   120
         TabIndex        =   11
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
         MICON           =   "material didactico.frx":1D90
         BC              =   8438015
         FC              =   0
         Picture         =   "material didactico.frx":1EFE
      End
   End
   Begin Jardin.xpgroupbox xpgroupbox2 
      Height          =   3135
      Left            =   6360
      TabIndex        =   12
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
         TabIndex        =   13
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
         MICON           =   "material didactico.frx":2058
         BC              =   8438015
         FC              =   0
         Picture         =   "material didactico.frx":21C6
      End
      Begin JeweledBut.JeweledButton busqueda 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2640
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
         MICON           =   "material didactico.frx":4CD0
         BC              =   8438015
         FC              =   0
         Picture         =   "material didactico.frx":4E3E
      End
      Begin JeweledBut.JeweledButton eliminar 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         MICON           =   "material didactico.frx":4F98
         BC              =   8438015
         FC              =   0
         Picture         =   "material didactico.frx":5106
      End
      Begin JeweledBut.JeweledButton guardar 
         Height          =   375
         Left            =   120
         TabIndex        =   16
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
         MICON           =   "material didactico.frx":56A0
         BC              =   8438015
         FC              =   0
         Picture         =   "material didactico.frx":580E
      End
      Begin JeweledBut.JeweledButton Actualizar 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         TX              =   "&Actualizar"
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
         MICON           =   "material didactico.frx":B43C
         BC              =   8438015
         FC              =   0
         Picture         =   "material didactico.frx":B5AA
      End
      Begin JeweledBut.JeweledButton modificar 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         TX              =   "&Modificar"
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
         MICON           =   "material didactico.frx":BB44
         BC              =   8438015
         FC              =   0
         Picture         =   "material didactico.frx":BCB2
      End
   End
   Begin JeweledBut.JeweledButton salir 
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      Top             =   3600
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
      MICON           =   "material didactico.frx":BE0C
      BC              =   8438015
      FC              =   0
      Picture         =   "material didactico.frx":BF7A
   End
   Begin JeweledBut.JeweledButton cancelar 
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   3600
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
      MICON           =   "material didactico.frx":C0D4
      BC              =   8438015
      FC              =   0
      Picture         =   "material didactico.frx":C242
   End
   Begin VB.Label numreg 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6480
      TabIndex        =   20
      Top             =   3600
      Width           =   60
   End
End
Attribute VB_Name = "material"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim nuevom As Boolean
Dim modim As Boolean

Private Sub Actualizar_Click()
nuevo.Enabled = True
modificar.Enabled = True
eliminar.Enabled = True
primero.Enabled = True
ultimo.Enabled = True
siguiente.Enabled = True
anterior.Enabled = True
Actualizar.Enabled = False
busqueda.Enabled = True
ref.Locked = True
nommat.Locked = True
can.Locked = True
If ModificadoMat = True Then
    bd.Recordset.Delete
    bd.Refresh
    guardarregistro
    ModificadoMat = False
End If
modim = False
End Sub

Private Sub anterior_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MovePrevious
    If bd.Recordset.BOF Then
        bd.Recordset.MoveFirst
    End If
    mostrarcampos
End If
ref.SetFocus
End Sub

Private Sub busqueda_Click()
MB.Formulario = Me.Name
MB.Descripcion = "Material Didáctico Existente"
elBuscador.Show
End Sub

Private Sub can_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
End Sub

Private Sub cancelar_Click()
If bd.Recordset.RecordCount > 0 Then
mostrarcampos
End If
If nuevom = True Then
    nuevo.Enabled = True
    modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    anterior.Enabled = True
    siguiente.Enabled = True
    ultimo.Enabled = True
    guardar.Enabled = False
    busqueda.Enabled = True
    nuevom = False
ElseIf modim = True Then
    nuevo.Enabled = True
    modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    ultimo.Enabled = True
    siguiente.Enabled = True
    anterior.Enabled = True
    Actualizar.Enabled = False
    busqueda.Enabled = True
    modim = False
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
        material.Show
   End If
End If
End If
End Sub

Private Sub feccom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    numfac.SetFocus
End If
End Sub

Private Sub Form_Activate()
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height
FormularioActivo = True

ref.Locked = True
nommat.Locked = True
can.Locked = True
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
menu.estado.Panels(4).Text = "Control de Material Didáctico"
On Error Resume Next
ConexionBD material, "select * from material"
If bd.Recordset.RecordCount > 0 Then
    mostrarcampos
End If
End Sub
Function mostrarcampos()
numreg = bd.Recordset.AbsolutePosition & " registro."
ref.Text = bd.Recordset!ref
nommat.Text = bd.Recordset!nommat
can.Text = bd.Recordset!can
End Function

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
guardarregistro
busqueda.Enabled = True
nuevom = False
End Sub
Function guardarregistro()

ConexionBD1 material, "select ref from material where ref=" & ref.Text
    If BD1.Recordset.RecordCount > 0 Then
        MsgBox "La referencia: " & ref.Text & " ya existe!", vbInformation, "Material Didáctico"
        Exit Function
    End If
nuevo.Enabled = True
modificar.Enabled = True
eliminar.Enabled = True
primero.Enabled = True
anterior.Enabled = True
siguiente.Enabled = True
ultimo.Enabled = True
guardar.Enabled = False

ref.Locked = True
nommat.Locked = True
can.Locked = True

On Error Resume Next
bd.Recordset.AddNew
bd.Recordset!ref = ref.Text
bd.Recordset!nommat = nommat.Text
bd.Recordset!can = Val(can.Text)
bd.Recordset!feccom = feccom.Value
bd.Recordset!numfac = numfac.Text
bd.Recordset.Update

End Function

Private Sub modificar_Click()
If bd.Recordset.RecordCount > 0 Then
ModificadoMat = True
ref.Locked = False
nommat.Locked = False
can.Locked = False

modificar.Enabled = False
Actualizar.Enabled = True
nuevo.Enabled = False
eliminar.Enabled = False
busqueda.Enabled = False
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False
modim = True
End If
End Sub

Private Sub nommat_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    can.SetFocus
End If
End Sub

Private Sub nuevo_Click()
'deshabilitar controles
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False
nuevo.Enabled = False
guardar.Enabled = True
modificar.Enabled = False
busqueda.Enabled = False
ref.Locked = False
nommat.Locked = False
can.Locked = False
nuevom = True


cajas material
ref.SetFocus
End Sub

Private Sub primero_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveFirst
    mostrarcampos
End If
ref.SetFocus
End Sub

Private Sub ref_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    ConexionBD1 material, "select ref from material where ref=" & ref.Text
    If BD1.Recordset.RecordCount > 0 Then
        MsgBox "La referencia: " & ref.Text & " ya existe!", vbInformation, "Material Didáctico"
        Exit Sub
    Else
        nommat.SetFocus
    End If
    
End If
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Private Sub siguiente_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveNext
    If bd.Recordset.EOF Then
        bd.Recordset.MoveLast
    End If
    mostrarcampos
End If
ref.SetFocus
End Sub

Private Sub ultimo_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveLast
    mostrarcampos
End If
ref.SetFocus
End Sub

Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "material.htm"
End Sub
