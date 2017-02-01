VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form Perfiles 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Perfiles de usuario"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Perfiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc bd1 
      Height          =   330
      Left            =   3120
      Top             =   3600
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
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5415
      _extentx        =   11456
      _extenty        =   5741
      font            =   "Perfiles.frx":628A
      backcolor       =   -2147483633
      caption         =   "Cuentas"
      Begin Jardin.xphelp xphelp1 
         Height          =   315
         Left            =   5040
         Top             =   120
         Width           =   315
         _extentx        =   556
         _extenty        =   556
      End
      Begin VB.ListBox cuentas 
         Height          =   2595
         ItemData        =   "Perfiles.frx":62B2
         Left            =   240
         List            =   "Perfiles.frx":62B4
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin MSAdodcLib.Adodc bd 
         Height          =   330
         Left            =   2880
         Top             =   0
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
         Caption         =   "Data1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin JeweledBut.JeweledButton nuevo 
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "Crear una cuenta de usuario"
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         TX              =   "Cuenta nueva"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "Perfiles.frx":62B6
         BC              =   8438015
         FC              =   0
         Picture         =   "Perfiles.frx":62D2
      End
      Begin JeweledBut.JeweledButton configurar 
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "Configurar una cuenta existente"
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         TX              =   "Configurar cuenta"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "Perfiles.frx":65EC
         BC              =   8438015
         FC              =   0
         Picture         =   "Perfiles.frx":6608
      End
      Begin JeweledBut.JeweledButton eliminar 
         Height          =   495
         Left            =   3000
         TabIndex        =   5
         Tag             =   "1"
         ToolTipText     =   "Eliminar una cuenta de usuario"
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         TX              =   "Eliminar cuenta"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "Perfiles.frx":6762
         BC              =   8438015
         FC              =   0
         Picture         =   "Perfiles.frx":677E
      End
      Begin JeweledBut.JeweledButton CERRAR 
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Tag             =   "1"
         ToolTipText     =   "Configurar una cuenta existente"
         Top             =   2520
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         TX              =   "Cerrar perfiles"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "Perfiles.frx":6D18
         BC              =   8438015
         FC              =   0
         Picture         =   "Perfiles.frx":6D34
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5160
      Top             =   3600
   End
   Begin Jardin.xpgroupbox frame2 
      Height          =   2295
      Left            =   240
      TabIndex        =   14
      Top             =   4320
      Width           =   5415
      _extentx        =   9551
      _extenty        =   4048
      font            =   "Perfiles.frx":6E8E
      backcolor       =   -2147483633
      caption         =   "Cuenta nueva"
      Begin VB.TextBox cc 
         Height          =   255
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   10
         PasswordChar    =   "•"
         TabIndex        =   30
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton admin 
         Caption         =   "Administrador"
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton normal 
         Caption         =   "Usuario Norrmal"
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox nombre 
         Height          =   270
         Left            =   2040
         MaxLength       =   13
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox pw 
         Height          =   255
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   10
         PasswordChar    =   "•"
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin JeweledBut.JeweledButton crear 
         Height          =   375
         Left            =   4080
         TabIndex        =   17
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         TX              =   "Crear"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "Perfiles.frx":6EB6
         BC              =   8438015
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cancelar 
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         TX              =   "Cancelar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "Perfiles.frx":6ED2
         BC              =   8438015
         FC              =   0
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmar:"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de usuario:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   1065
      End
   End
   Begin Jardin.xpgroupbox frame3 
      Height          =   2295
      Left            =   240
      TabIndex        =   7
      Top             =   4320
      Width           =   5415
      _extentx        =   9551
      _extenty        =   4048
      font            =   "Perfiles.frx":6EEE
      backcolor       =   -2147483633
      caption         =   "Configurar cuenta (cambiar)"
      Begin VB.TextBox ccc 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   10
         PasswordChar    =   "•"
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton admin1 
         Caption         =   "Administrador"
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton normal1 
         Caption         =   "Usuario Normal"
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox cpw 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   10
         PasswordChar    =   "•"
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox cnombre 
         Height          =   270
         Left            =   2040
         MaxLength       =   13
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin JeweledBut.JeweledButton cancelar1 
         Height          =   375
         Left            =   4080
         TabIndex        =   10
         Tag             =   "1"
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         TX              =   "Cancelar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "Perfiles.frx":6F16
         BC              =   8438015
         FC              =   0
      End
      Begin JeweledBut.JeweledButton cambiar 
         Height          =   375
         Left            =   4080
         TabIndex        =   11
         Tag             =   "1"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         TX              =   "Cambiar"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "Perfiles.frx":6F32
         BC              =   8438015
         FC              =   0
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmar:"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil:"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de usuario:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1710
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1065
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   60
   End
End
Attribute VB_Name = "Perfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bcerrar_Click()
Panel.Enabled = True
Unload Me
End Sub

Private Sub cambiar_Click()
'verifica que la contraseña sea igual en juntas cajas
If cpw.Text = "" Or ccc.Text = "" Then
    MsgBox "Por favor ingrese la contraseña y/o la confirmación de esta!", vbInformation, "Configurar Cuenta"
    Exit Sub
End If
If cpw.Text <> ccc.Text Then
    MsgBox "La contraseña no coincide con la confirmación de esta!", vbExclamation, "Configurar Cuenta"
    ccc.SetFocus
    Exit Sub
End If
Me.Height = 4600
frame3.Visible = False
CambiarCuenta
eliminar.Enabled = True
nuevo.Enabled = True
End Sub

Private Sub cancelar_Click()
nuevo.Enabled = True
eliminar.Enabled = True
configurar.Enabled = True
Me.Height = 4600
End Sub

Private Sub cancelar1_Click()
nuevo.Enabled = True
eliminar.Enabled = True
configurar.Enabled = True
Me.Height = 4600
End Sub

Private Sub cerrar_Click()
Unload Me
End Sub

Private Sub configurar_Click()
If usuario = "Administrador" Then
    Label7.Visible = False
    admin1.Visible = False
    normal1.Visible = False
ElseIf cuentas.List(cuentas.ListIndex) = usuario Then
    Label7.Visible = False
    admin1.Visible = False
    normal1.Visible = False
ElseIf cuentas.List(cuentas.ListIndex) <> usuario Then
    MsgBox "La cuenta que quiere modificar no es la suya actualmente!", vbExclamation, "Configuraciòn de Usuarios"
    Exit Sub
End If

ConfigurarCuenta cuentas.List(cuentas.ListIndex)
End Sub



Private Sub crear_Click()
If admin.Value = False And normal.Value = False Then
    MsgBox "Escoja un perfil de usuario!", vbInformation, "Configuración de Cuentas"
    Exit Sub
End If
'verifica que la contraseña sea introducida en ambos lados
If pw.Text = "" Or cc.Text = "" Then
    MsgBox "Por favor ingrese la contraseña y/o la confirmación de esta!", vbInformation, "Crear Cuenta"
    Exit Sub
End If
If pw.Text <> cc.Text Then
    MsgBox "La contraseña no coincide con la confirmación de esta!", vbExclamation, "Crear Cuenta"
    Exit Sub
End If
Me.Height = 4600
frame2.Visible = False
CuentaNueva
nombre.Text = ""
pw.Text = ""
eliminar.Enabled = True
configurar.Enabled = True
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
des = Description
End Sub

Private Sub eliminar_Click()
BuscarUs usuario
EliminarCuenta cuentas.List(cuentas.ListIndex)
End Sub

Private Sub Form_Activate()
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height
FormularioActivo = True
On Error Resume Next
ConexionBD1 Perfiles, "select * from acceso"
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
menu.estado.Panels(4).Text = "Configuración de Usuarios"
'Controles Perfiles
On Error GoTo Error
ConexionBD Perfiles, "select * from acceso"
Label1 = "Usuario actual: " & usuario
CargarCuentas
SalirError:
    Exit Sub
Error:
    If Err.Number Then
        valor = Mid(des, InStr(des, "[Administrador de controladores ODBC]") + 38)
       MsgBox valor, vbCritical, "Error"
    End If
End Sub


Private Sub Form_Resize()
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)
menu.Enabled = True
FormularioActivo = False
menu.estado.Panels(4).Text = "Menú Principal"
bd.Recordset.Close
Set bd.Recordset = Nothing
End Sub

Private Sub nuevo_Click()
If usuario = "Administrador" Then
    eliminar.Enabled = False
    configurar.Enabled = False
    Me.Height = 7155
    frame2.Visible = True
    frame3.Visible = False
    'crear usuario nuevo en la BD
    crear.Enabled = False
    chmHelp.PopUp "El sistema acepta hasta diez caracteres. Para mayor seguridad utilice más de 4 caracteres en su contraseña!"
Else
    MsgBox usuario & " No tiene permisos para realizar esta operación!", vbInformation, "Perfile de Usuario"
End If
End Sub

Private Sub Timer1_Timer()
If nombre.Text <> "" And pw.Text <> "" Then
    crear.Enabled = True
End If
If cnombre.Text <> "" And cpw.Text <> "" Then
    cambiar.Enabled = True
End If
End Sub

Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "usuarios.htm"
End Sub
