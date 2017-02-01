VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Compras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Compras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7095
   Begin MSAdodcLib.Adodc bd1 
      Height          =   330
      Left            =   3720
      Top             =   5040
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
   Begin VB.Frame Frame3 
      Caption         =   "Opciones"
      Height          =   2295
      Left            =   5280
      TabIndex        =   8
      Top             =   120
      Width           =   1695
      Begin JeweledBut.JeweledButton nuevo 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
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
         MICON           =   "Compras.frx":014A
         BC              =   8438015
         FC              =   0
         Picture         =   "Compras.frx":02B8
      End
      Begin JeweledBut.JeweledButton busqueda 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1800
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
         MICON           =   "Compras.frx":2DC2
         BC              =   8438015
         FC              =   0
         Picture         =   "Compras.frx":2F30
      End
      Begin JeweledBut.JeweledButton eliminar 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1320
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
         MICON           =   "Compras.frx":308A
         BC              =   8438015
         FC              =   0
         Picture         =   "Compras.frx":31F8
      End
      Begin JeweledBut.JeweledButton guardar 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   840
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
         MICON           =   "Compras.frx":3792
         BC              =   8438015
         FC              =   0
         Picture         =   "Compras.frx":3900
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Navegación"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   3495
      Begin JeweledBut.JeweledButton primero 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Primer Registro"
         Top             =   240
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
         MICON           =   "Compras.frx":952E
         BC              =   8438015
         FC              =   0
         Picture         =   "Compras.frx":969C
      End
      Begin JeweledBut.JeweledButton siguiente 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         ToolTipText     =   "Siguiente Registro"
         Top             =   720
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
         MICON           =   "Compras.frx":97F6
         BC              =   8438015
         FC              =   0
         Picture         =   "Compras.frx":9964
      End
      Begin JeweledBut.JeweledButton ultimo 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         ToolTipText     =   "Ultimo Registro"
         Top             =   240
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
         MICON           =   "Compras.frx":9ABE
         BC              =   8438015
         FC              =   0
         Picture         =   "Compras.frx":9C2C
      End
      Begin JeweledBut.JeweledButton anterior 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Anterior Registro"
         Top             =   720
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
         MICON           =   "Compras.frx":9D86
         BC              =   8438015
         FC              =   0
         Picture         =   "Compras.frx":9EF4
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Compras de Material"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin Jardin.xpgroupbox xpgroupbox1 
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1296
         Caption         =   ""
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
         Begin VB.OptionButton meo 
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton nmo 
            Height          =   255
            Left            =   2880
            TabIndex        =   36
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Material Nuevo"
            Height          =   195
            Left            =   3240
            TabIndex        =   39
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Material Existente"
            Height          =   195
            Left            =   480
            TabIndex        =   38
            Top             =   240
            Width           =   1515
         End
      End
      Begin MSAdodcLib.Adodc bd 
         Height          =   330
         Left            =   3000
         Top             =   120
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
      Begin MSComCtl2.DTPicker fecfac 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Tag             =   "1"
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   38193
      End
      Begin VB.TextBox numfac 
         Height          =   285
         Left            =   2040
         MaxLength       =   7
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.Frame nmm 
         Height          =   2175
         Left            =   120
         TabIndex        =   26
         Top             =   2520
         Visible         =   0   'False
         Width           =   4815
         Begin VB.TextBox valorn 
            Height          =   285
            Left            =   2040
            MaxLength       =   8
            TabIndex        =   30
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox cann 
            Height          =   285
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   29
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox refn 
            Height          =   285
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   28
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox nommat 
            Height          =   285
            Left            =   2040
            MaxLength       =   20
            TabIndex        =   27
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Valor Compra:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   1800
            Width           =   1275
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Referencia Material:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1725
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   1320
            Width           =   840
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre del material"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   1755
         End
      End
      Begin VB.Frame med 
         Height          =   1575
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox can 
            Height          =   285
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   21
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox refmat 
            Height          =   315
            Left            =   2040
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox valor 
            Height          =   285
            Left            =   2040
            MaxLength       =   8
            TabIndex        =   19
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   840
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Referencia Material:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   1725
         End
         Begin VB.Label desmat 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   3000
            TabIndex        =   23
            Tag             =   "1"
            Top             =   360
            Width           =   60
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Valor Compra:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   1200
            Width           =   1275
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Factura:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número de Factura:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1710
      End
      Begin VB.Label numcom 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2040
         TabIndex        =   2
         Tag             =   "1"
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número de compra:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1725
      End
   End
   Begin JeweledBut.JeweledButton salir 
      Height          =   375
      Left            =   5280
      TabIndex        =   17
      Top             =   5640
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
      MICON           =   "Compras.frx":A04E
      BC              =   8438015
      FC              =   0
      Picture         =   "Compras.frx":A1BC
   End
   Begin Jardin.xphelp xphelp1 
      Height          =   315
      Left            =   6000
      Top             =   2520
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin JeweledBut.JeweledButton cancelar 
      Height          =   375
      Left            =   5280
      TabIndex        =   40
      Top             =   5160
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
      MICON           =   "Compras.frx":A316
      BC              =   8438015
      FC              =   0
      Picture         =   "Compras.frx":A484
   End
End
Attribute VB_Name = "Compras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nuevoc As Boolean
Private Sub anterior_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MovePrevious
    If bd.Recordset.BOF Then
        bd.Recordset.MoveFirst
    End If
    med.Visible = True
    mostrarcampos
End If
numfac.SetFocus
End Sub

Private Sub busqueda_Click()
MB.Formulario = Me.Name
MB.Descripcion = "Compras"
elBuscador.Show
End Sub

Private Sub can_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    valor.SetFocus
End If
End Sub

Private Sub cancelar_Click()
If bd.Recordset.RecordCount > 0 Then
mostrarcampos
End If
If nuevoc = True Then
    nuevo.Enabled = True
    'modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    anterior.Enabled = True
    siguiente.Enabled = True
    ultimo.Enabled = True
    guardar.Enabled = False
    busqueda.Enabled = True
    nuevoc = False
End If
bloquearc
End Sub

Private Sub cann_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    valorn.SetFocus
End If
End Sub

Private Sub eliminar_Click()
If bd.Recordset.RecordCount > 0 Then
If MsgBox("Está seguro de querer eliminar este registro?", vbYesNo + vbQuestion, "Eliminar Registro") = vbYes Then
    bd.Recordset.Delete
    If bd.Recordset.RecordCount > 0 Then
    bd.Recordset.MoveFirst
    bd.Refresh
    mostrarcampos
    Else
        Unload Me
        Compras.Show
   End If
End If
End If
End Sub

Private Sub fecfac_LostFocus()
Dim ANIOs As Integer
ANIOs = Format(Date, "yyyy")

If fecfac.Year > ANIOs Or fecfac.Year < (ANIOs - 1) Then
    MsgBox "Fecha no permitida!", vbExclamation, "Compras"
    fecfac.SetFocus
    Exit Sub
End If

End Sub

Private Sub Form_Activate()
FormularioActivo = True
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height

End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
LlenarCombos
menu.estado.Panels(4).Text = "Control de compras de Material"
On Error Resume Next
ConexionBD Compras, "select * from compras order by numcom"
If bd.Recordset.RecordCount > 0 Then
   med.Visible = True
   mostrarcampos
End If
'bloquear cajas
bloquearc
'Deshabilitarl Compras
End Sub
Sub bloquearc()
numfac.Locked = True
refmat.Locked = True
can.Locked = True
valor.Locked = True
refn.Locked = True
nommat.Locked = True
cann.Locked = True
valorn.Locked = True
End Sub
Private Sub LlenarCombos()
'llenar cargo
ConexionBD1 Compras, "select ref from material"
refmat.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        refmat.AddItem BD1.Recordset!ref
        BD1.Recordset.MoveNext
    Next i
End If
End Sub
Function mostrarcampos()
numcom = bd.Recordset!numcom
numfac.Text = bd.Recordset!numfac
fecfac.Value = bd.Recordset!fecfac
can.Text = bd.Recordset!can
refmat.Text = bd.Recordset!refmat
valor.Text = bd.Recordset!valor
ConexionBD1 Compras, "select * from material where ref=" & refmat.Text
If BD1.Recordset.RecordCount > 0 Then
    desmat = BD1.Recordset!nommat
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
FormularioActivo = False
menu.estado.Panels(4).Text = "Menú Principal"
End Sub

Private Sub guardar_Click()
If MsgBox("Está seguro(a) de guardar la compra?" & vbCrLf & "Recuerde que al guardar los datos no es posible modificarlos!", vbYesNo + vbQuestion, "Compras") = vbYes Then
'validamos cajas de texto y combos
If meo.Value = True Then
    If numfac.Text = "" Or can.Text = "" Or refmat.Text = "" Or valor.Text = "" Then
        MsgBox "Falta datos por ingresar!", vbInformation, "Material Despachado"
        Exit Sub
    End If
End If
If nmo.Value = True Then
    If refn.Text = "" Or nommat.Text = "" Or cann.Text = "" Or numfac.Text = "" Then
        MsgBox "Falta datos por ingresar!", vbInformation, "Material Despachado"
        Exit Sub
    End If
End If

If meo.Value = False And nmo.Value = False Then
    MsgBox "Para guardar seleccione que tipo de compra desea hacer!", vbInformation, "Material Despachado"
    Exit Sub
End If
guardarregistro
nuevoc = False
End If
End Sub
Function guardarregistro()
If meo.Value = True Then 'si el material es existente
    bd.Recordset.AddNew
        bd.Recordset!numcom = numcom
        bd.Recordset!numfac = numfac
        bd.Recordset!fecfac = fecfac.Value
        bd.Recordset!can = can.Text
        bd.Recordset!refmat = refmat.Text
        bd.Recordset!valor = valor.Text
    bd.Recordset.Update
    'modifica el registro en material existente
    ConexionBD1 Compras, "select * from material where ref=" & refmat.Text
    Dim modi
    modi = BD1.Recordset.EditMode
    BD1.Recordset!can = Val(BD1.Recordset!can) + Val(can.Text)
    BD1.Recordset.Update
ElseIf nmo.Value = True Then 'si el material es nuevo
    'pregunta si la referencia del material ya esta
    ConexionBD1 Compras, "select ref from material where ref=" & refn
    If BD1.Recordset.RecordCount = 1 Then
        MsgBox "La referencia: " & refn.Text & " ya existe en el sistema!", vbInformation, "Compras"
        Exit Function
    End If
    bd.Recordset.AddNew
        bd.Recordset!numcom = numcom
        bd.Recordset!numfac = numfac
        bd.Recordset!fecfac = fecfac.Value
        bd.Recordset!can = cann.Text
        bd.Recordset!refmat = refn.Text
        bd.Recordset!valor = valorn.Text
    bd.Recordset.Update
    'guarda en material existente
    ConexionBD1 Compras, "select * from material"
    BD1.Recordset.AddNew
        BD1.Recordset!ref = refn.Text
        BD1.Recordset!nommat = nommat.Text
        BD1.Recordset!can = cann.Text
    BD1.Recordset.Update
End If
'controla los botones
guardar.Enabled = False
nuevo.Enabled = True
eliminar.Enabled = True
busqueda.Enabled = True
primero.Enabled = True
anterior.Enabled = True
siguiente.Enabled = True
ultimo.Enabled = True
bloquearc
End Function

Private Sub meo_Click()
If meo.Value = True Then
    med.Visible = True
    nmm.Visible = False
End If
End Sub

Private Sub nmo_Click()
If nmo.Value = True Then
    med.Visible = False
    nmm.Visible = True
End If
End Sub


Private Sub nommat_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    cann.SetFocus
End If
End Sub

Private Sub nuevo_Click()
'limpiar cajas
Habilitarl Compras
cajas Compras
numcom = ""
desmat = ""
med.Visible = False
'genera el autonumerico para numero de compra
ConexionBD1 Compras, "select * from compras"
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveLast
    numcom = BD1.Recordset!numcom + 1
ElseIf BD1.Recordset.RecordCount = 0 Then
    numcom = 1
End If
numfac.SetFocus
'controla los botones
guardar.Enabled = True
nuevo.Enabled = False
eliminar.Enabled = False
busqueda.Enabled = False
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False
nuevoc = True
End Sub

Private Sub numfac_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    fecfac.SetFocus
End If
End Sub

Private Sub primero_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveFirst
    med.Visible = True
    mostrarcampos
End If
numfac.SetFocus
End Sub

Private Sub refmat_Click()
'cuando selecciona una referencia busca la descripcion del material
On Error Resume Next
ConexionBD1 Compras, "select nommat from material where ref=" & refmat.Text
If BD1.Recordset.RecordCount > 0 Then
    desmat = BD1.Recordset!nommat
End If
End Sub

Private Sub refmat_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    'cuando selecciona una referencia busca la descripcion del material
    If refmat.Text <> "" Then
        ConexionBD1 Compras, "select nommat from material where ref=" & refmat.Text
        If BD1.Recordset.RecordCount > 0 Then
            desmat = BD1.Recordset!nommat
        End If
        can.SetFocus
    End If
End If
End Sub

Private Sub refn_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    ConexionBD1 Compras, "SELECT REF FROM MATERIAL WHERE REF=" & refn
    If BD1.Recordset.RecordCount = 1 Then
        MsgBox "Esta referencia de material ya existe!", vbInformation, "Compras"
        refn.SetFocus
        Exit Sub
    End If
    nommat.SetFocus
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
    med.Visible = True
    mostrarcampos
End If
numfac.SetFocus
End Sub



Private Sub ultimo_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveLast
    med.Visible = True
    mostrarcampos
End If
numfac.SetFocus
End Sub

Private Sub valor_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
End Sub

Private Sub valorn_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
End Sub

Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "compras.htm"

End Sub
