VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Pagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "pagos.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6930
   Begin Jardin.xphelp xphelp1 
      Height          =   315
      Left            =   4800
      Top             =   4920
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin MSAdodcLib.Adodc bd1 
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
   Begin MSAdodcLib.Adodc bd 
      Height          =   330
      Left            =   3840
      Top             =   3960
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
   Begin Jardin.xpgroupbox xpgroupbox1 
      Height          =   3735
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6588
      Caption         =   "Matricula"
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
      Begin VB.TextBox pagmat 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   28
         Tag             =   "1"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox mes 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "pagos.frx":030A
         Left            =   2640
         List            =   "pagos.frx":0332
         TabIndex        =   27
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox otro 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "pagos.frx":039B
         Left            =   2640
         List            =   "pagos.frx":03AB
         TabIndex        =   24
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox numdoc 
         Height          =   315
         ItemData        =   "pagos.frx":03E9
         Left            =   2640
         List            =   "pagos.frx":03EB
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker feccon 
         Height          =   375
         Left            =   2640
         TabIndex        =   21
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50528257
         CurrentDate     =   38119
      End
      Begin VB.TextBox valcon 
         Height          =   285
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   14
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox numcon 
         Height          =   285
         Left            =   2640
         MaxLength       =   15
         TabIndex        =   13
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes Pensión"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto Otros"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número Documento"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto Pago"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Consignación          $"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   2355
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha  Consignación"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   1755
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Consignación"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   3360
         Width           =   2145
      End
   End
   Begin Jardin.xpgroupbox xpgroupbox2 
      Height          =   3615
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   6376
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
         TabIndex        =   1
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
         MICON           =   "pagos.frx":03ED
         BC              =   8438015
         FC              =   0
         Picture         =   "pagos.frx":055B
      End
      Begin JeweledBut.JeweledButton busqueda 
         Height          =   375
         Left            =   120
         TabIndex        =   2
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
         MICON           =   "pagos.frx":3065
         BC              =   8438015
         FC              =   0
         Picture         =   "pagos.frx":31D3
      End
      Begin JeweledBut.JeweledButton eliminar 
         Height          =   375
         Left            =   120
         TabIndex        =   3
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
         MICON           =   "pagos.frx":332D
         BC              =   8438015
         FC              =   0
         Picture         =   "pagos.frx":349B
      End
      Begin JeweledBut.JeweledButton guardar 
         Height          =   375
         Left            =   120
         TabIndex        =   4
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
         MICON           =   "pagos.frx":3A35
         BC              =   8438015
         FC              =   0
         Picture         =   "pagos.frx":3BA3
      End
      Begin JeweledBut.JeweledButton Actualizar 
         Height          =   375
         Left            =   120
         TabIndex        =   5
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
         MICON           =   "pagos.frx":97D1
         BC              =   8438015
         FC              =   0
         Picture         =   "pagos.frx":993F
      End
      Begin JeweledBut.JeweledButton modificar 
         Height          =   375
         Left            =   120
         TabIndex        =   6
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
         MICON           =   "pagos.frx":9ED9
         BC              =   8438015
         FC              =   0
         Picture         =   "pagos.frx":A047
      End
      Begin JeweledBut.JeweledButton parametro 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   3120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         TX              =   "Parámetos"
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
         MICON           =   "pagos.frx":A1A1
         BC              =   8438015
         FC              =   0
      End
   End
   Begin Jardin.xpgroupbox xpgroupbox3 
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   4080
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
         MICON           =   "pagos.frx":A30F
         BC              =   8438015
         FC              =   0
         Picture         =   "pagos.frx":A47D
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
         MICON           =   "pagos.frx":A5D7
         BC              =   8438015
         FC              =   0
         Picture         =   "pagos.frx":A745
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
         MICON           =   "pagos.frx":A89F
         BC              =   8438015
         FC              =   0
         Picture         =   "pagos.frx":AA0D
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
         MICON           =   "pagos.frx":AB67
         BC              =   8438015
         FC              =   0
         Picture         =   "pagos.frx":ACD5
      End
   End
   Begin JeweledBut.JeweledButton salir 
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   4920
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
      MICON           =   "pagos.frx":AE2F
      BC              =   8438015
      FC              =   0
      Picture         =   "pagos.frx":AF9D
   End
   Begin JeweledBut.JeweledButton cancelar 
      Height          =   375
      Left            =   5280
      TabIndex        =   31
      Top             =   4440
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
      MICON           =   "pagos.frx":B0F7
      BC              =   8438015
      FC              =   0
      Picture         =   "pagos.frx":B265
   End
   Begin VB.Label numreg 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5160
      TabIndex        =   29
      Top             =   4080
      Width           =   60
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000018&
      Caption         =   "Formato Fecha: dia/mes/año"
      Height          =   495
      Left            =   3720
      TabIndex        =   25
      Top             =   4440
      Width           =   1395
   End
End
Attribute VB_Name = "Pagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nuevop As Boolean
Dim modip As Boolean
Dim i As Integer

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
parametro.Enabled = True
Bloquear
If ModificadoP = True Then
    bd.Recordset.Delete
    guardarregistro
    ModificadoP = False
End If
modip = False
End Sub

Private Sub anterior_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MovePrevious
    If bd.Recordset.BOF Then
        bd.Recordset.MoveFirst
    End If
    mostrarcampos
End If
numdoc.SetFocus
End Sub

Private Sub bd_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub bd1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub busqueda_Click()
MB.Formulario = Me.Name
MB.Descripcion = "Pagos (Pensión, Matrricula, Otros)"
elBuscador.Show
End Sub

Private Sub cancelar_Click()
If bd.Recordset.RecordCount > 0 Then
mostrarcampos
End If
If nuevop = True Then
    nuevo.Enabled = True
    modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    anterior.Enabled = True
    siguiente.Enabled = True
    ultimo.Enabled = True
    guardar.Enabled = False
    busqueda.Enabled = True
    nuevop = False
ElseIf modip = True Then
    nuevo.Enabled = True
    modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    ultimo.Enabled = True
    siguiente.Enabled = True
    anterior.Enabled = True
    Actualizar.Enabled = False
    busqueda.Enabled = True
    modip = False
End If
parametro.Enabled = True
Bloquear
End Sub

Private Sub eliminar_Click()
If bd.Recordset.RecordCount > 0 Then
If MsgBox("Está seguro de querer eliminar el registro?", vbYesNo + vbQuestion, "Eliminar Registro") = vbYes Then
   bd.Recordset.Delete
   If bd.Recordset.RecordCount > 0 Then
    bd.Refresh
    bd.Recordset.MoveFirst
    mostrarcampos
    descargar = False
    Else
        descargar = True
        Control = xpgroupbox1.Caption
        Unload Pagos
        Pagos.Show
   End If
End If
End If
End Sub


Private Sub feccon_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    numcon.SetFocus
End If
End Sub

Private Sub feccon_LostFocus()
Dim ANIOs As Integer
ANIOs = Format(Date, "yyyy")

If feccon.Year > ANIOs Or feccon.Year < (ANIOs - 1) Then
    MsgBox "Fecha no permitida!", vbExclamation, "Pagos"
    feccon.SetFocus
    Exit Sub
End If

End Sub

Sub Bloquear()
'bloquear cajas
numdoc.Locked = True
pagmat.Locked = True
mes.Locked = True
otro.Locked = True
valcon.Locked = True
numcon.Locked = True

End Sub
Private Sub Form_Activate()

pagmat.Text = Me.Tag
If descargar = True Then
    pagmat.Text = Control
End If
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height
FormularioActivo = True

Bloquear
numdoc.SetFocus

'carga el data con cada unos de los pagos
If pagmat.Text = "Matricula" Then
    On Error Resume Next
    ConexionBD Pagos, "select * from pagos where tipopago='Matricula';"
    If bd.Recordset.RecordCount > 0 Then
        mostrarcampos
    Else
        primero.Enabled = False
        siguiente.Enabled = False
        anterior.Enabled = False
        ultimo.Enabled = False
    End If
    busqueda.Visible = True
ElseIf pagmat.Text = "Pensión" Then
    On Error Resume Next
    ConexionBD Pagos, "select * from pagos where tipopago='Pensión';"
    If bd.Recordset.RecordCount > 0 Then
        mostrarcampos
    End If
    busqueda.Visible = False
    mes.Enabled = True
ElseIf pagmat.Text = "Otros" Then
    On Error Resume Next
    ConexionBD Pagos, "select * from pagos where tipopago='Otros';"
    If bd.Recordset.RecordCount > 0 Then
        mostrarcampos
    End If
    otro.Enabled = True
    busqueda.Visible = False
End If
On Error Resume Next
ConexionBD1 Pagos, "select * from matricula"
i = 0

If BD1.Recordset.RecordCount > 0 Then
Do Until BD1.Recordset.EOF
  numdoc.AddItem BD1.Recordset!numdoc
  i = i + 1
  BD1.Recordset.MoveNext
Loop
BD1.Recordset.MoveFirst
ElseIf BD1.Recordset.RecordCount = 0 Then
    MsgBox "No hay ningún registro en matricula!" & vbCrLf & "Para realizar algún pago debe haber algún niño matriculado.", vbExclamation, "Pagos"
    Unload Me
End If
If Para = True Then
    LlenarCombos
    Para = False
End If

End Sub
Private Sub LlenarCombos()
'llenar conceptos otros
ConexionBD1 Pagos, "select * from parametrizacion where tippar=30" & " order by dato;"
otro.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        otro.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
BD1.Recordset.Close
Set BD1.Recordset = Nothing
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
LlenarCombos
menu.estado.Panels(4).Text = "Control de Pagos (Matricula, Pensión, Otros)"

End Sub
Function guardarregistro()

'conexionbd1 pagos,"SELECT * FROM matricula WHERE numdoc='" + numdoc.Text + "';"

On Error Resume Next
bd.Recordset.AddNew
bd.Recordset!numdoc = numdoc.Text
bd.Recordset!tipopago = pagmat.Text
If pagmat.Text = "Otros" Then
    bd.Recordset!otro = otro.Text
ElseIf pagmat.Text = "Pensión" Then
    bd.Recordset!mes = mes.Text
End If
bd.Recordset!valcon = valcon.Text
bd.Recordset!feccon = feccon.Value
bd.Recordset!numcon = numcon.Text
bd.Recordset.Update
Bloquear
'desbloquear controles
nuevo.Enabled = True
modificar.Enabled = True
eliminar.Enabled = True
primero.Enabled = True
anterior.Enabled = True
siguiente.Enabled = True
ultimo.Enabled = True
guardar.Enabled = False
parametro.Enabled = True
End Function
Function mostrarcampos()
numreg = bd.Recordset.AbsolutePosition & " registro."
numdoc.Text = bd.Recordset!numdoc
pagmat.Text = bd.Recordset!tipopago
If IsNull(bd.Recordset!otro) = False Then
    otro.Text = bd.Recordset!otro
End If
If IsNull(bd.Recordset!mes) = False Then
    mes.Text = bd.Recordset!mes
End If
valcon.Text = bd.Recordset!valcon
feccon.Value = bd.Recordset!feccon
numcon.Text = bd.Recordset!numcon
End Function
Function avanzar()
 If tecla = 13 Then
  SendKeys "{tab}"
  KeyAscii = 0: tecla = 0
 End If
End Function

Private Sub Form_Resize()
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)
FormularioActivo = False
menu.estado.Panels(4).Text = "Menú Principal"
bd.Recordset.Close
End Sub

Private Sub guardar_Click()
If numdoc.Text = "" Then
    MsgBox "No ha ingresado el número de documento del niño!", vbInformation, "Pagos"
    Exit Sub
End If
If pagmat.Text = "Matricula" Then
    
    ConexionBD1 Pagos, "select numdoc from pagos where numdoc='" & numdoc.Text & "'" & " and tipopago='Matricula'"
    If BD1.Recordset.RecordCount > 0 Then
        MsgBox "Pago de matricula ya realizada!", vbInformation, "Pago Matricula"
        Exit Sub
    Else
        numdoc.SetFocus
    End If
    
    If valcon.Text = "" Or numcon.Text = "" Then
        MsgBox "Faltan datos por ingresar!", vbInformation, "Pagos"
        Exit Sub
    End If
ElseIf pagmat.Text = "Pensión" Then
    ConexionBD1 Pagos, "select numdoc from pagos where numdoc='" & numdoc.Text & "'" & " and tipopago='Pensión' and mes='" & mes.Text & "'"
    If BD1.Recordset.RecordCount = 1 Then
        MsgBox "Mes ya cancelado!", vbInformation, "Pago Pensión"
        Exit Sub
    Else
        mes.SetFocus
    End If
    If mes.Text = "" Or valcon.Text = "" Or numcon.Text = "" Then
        MsgBox "Faltan datos por ingresar!", vbInformation, "Pagos"
        Exit Sub
    End If
ElseIf pagmat.Text = "Otros" Then
    If otro.Text = "" Or valcon.Text = "" Or numcon.Text = "" Then
        MsgBox "Faltan datos por ingresar!", vbInformation, "Pagos"
        Exit Sub
    End If
End If
guardarregistro
busqueda.Enabled = True
nuevop = False
End Sub

Private Sub List1_Click()

End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    'If PagoActual = 1 Then
        ConexionBD1 Pagos, "select * from pagos where numdoc='" & numdoc.Text & "'" & " and mes='" & mes.Text & "'"
        If BD1.Recordset.RecordCount > 0 Then
            MsgBox "La pensión para el mes " & mes.Text & " ya se realizó!", vbInformation, "Pago de Pensión"
            Exit Sub
        Else
            valcon.SetFocus
        End If
    'End If
End If
End Sub

Private Sub modificar_Click()
If bd.Recordset.RecordCount > 0 Then
ModificadoP = True
modip = True
'desbloquear cajas
numdoc.Locked = False
'pagmat.Locked = False
valcon.Locked = False
numcon.Locked = False
numdoc.Locked = False
otro.Locked = False
parametro.Enabled = False
modificar.Enabled = False
Actualizar.Enabled = True
nuevo.Enabled = False
eliminar.Enabled = False
busqueda.Enabled = False
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False

Dim modif As Variant
modif = bd.Recordset.EditMode
End If
End Sub

Private Sub nuevo_Click()
nuevop = True
'deshabilitar controles
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False
nuevo.Enabled = False
guardar.Enabled = True
modificar.Enabled = False
busqueda.Enabled = False
eliminar.Enabled = False
parametro.Enabled = False
'bloquear cajas
numdoc.Locked = False
pagmat.Locked = False

valcon.Locked = False
numcon.Locked = False

cajas Pagos
If PagoActual = 0 Then
    pagmat.Text = "Matricula"
    valcon.Enabled = True
    numcon.Enabled = True
    feccon.Enabled = True
ElseIf PagoActual = 1 Then
    pagmat.Text = "Pensión"
    mes.Text = ""
    mes.Enabled = True
    valcon.Enabled = True
    numcon.Enabled = True
    feccon.Enabled = True
ElseIf PagoActual = 2 Then
    pagmat.Text = "Otros"
    otro.Enabled = True
    valcon.Enabled = True
    numcon.Enabled = True
    feccon.Enabled = True
End If

numdoc.SetFocus
End Sub



Private Sub numcon_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
End Sub

Private Sub numdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    If numdoc.Text <> "" Then
        If PagoActual = 0 Then 'matricula
            ConexionBD1 Pagos, "select numdoc from pagos where numdoc='" & numdoc.Text & "'" & " and tipopago='Matricula'"
            If BD1.Recordset.RecordCount > 0 Then
                MsgBox "Pago de matricula ya realizada!", vbInformation, "Pago Matricula"
                Exit Sub
            Else
                valcon.SetFocus
            End If
        End If
    End If
End If
End Sub

Private Sub otro_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    valcon.SetFocus
End If
End Sub



Private Sub pagmat_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    valcon.SetFocus
End If
End Sub

Private Sub parametro_Click()
Para = True
ingresos.Show
End Sub

Private Sub primero_Click()
If bd.Recordset.RecordCount > 0 Then
    bd.Recordset.MoveFirst
    mostrarcampos
End If
numdoc.SetFocus
End Sub
Private Sub salir_Click()
Unload Me
End Sub

Private Sub siguiente_Click()
If bd.Recordset.RecordCount > 0 Then
    bd.Recordset.MoveNext
    If bd.Recordset.EOF Then
        bd.Recordset.MoveLast
    End If
    mostrarcampos
End If
numdoc.SetFocus
End Sub

Private Sub ultimo_Click()
If bd.Recordset.RecordCount > 0 Then
    bd.Recordset.MoveLast
    mostrarcampos
End If
numdoc.SetFocus
End Sub

Private Sub valcon_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    feccon.SetFocus
End If
End Sub

Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "pagos.htm"
End Sub
