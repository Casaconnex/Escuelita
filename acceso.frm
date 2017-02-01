VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form acceso 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inicio de Sesión en SISJACE"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "acceso.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   146
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox pass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "•"
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.ComboBox user 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc bd 
      Height          =   330
      Left            =   840
      Top             =   2280
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer instancia 
      Interval        =   1
      Left            =   2160
      Top             =   2280
   End
   Begin Jardin.xphelp ayuda 
      Height          =   315
      Left            =   1800
      Top             =   1680
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin JeweledBut.JeweledButton cancelar 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Salir de Dedalus"
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "&Cancelar"
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
      MICON           =   "acceso.frx":0442
      BC              =   8438015
      FC              =   0
      Picture         =   "acceso.frx":045E
   End
   Begin JeweledBut.JeweledButton aceptar 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "&Aceptar"
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
      MPTR            =   0
      MICON           =   "acceso.frx":05B8
      BC              =   8438015
      FC              =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JARDÍN ARTÍSTICO COMUNITARIO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   3240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LA ESCUELITA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   360
      Width           =   1365
   End
End
Attribute VB_Name = "acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aceptar_Click()
usuario = ""
If user.Text = "" Or pass.Text = "" Then
    MsgBox "Por favor escoja un usuario y/o digite la contraseña!", vbCritical, "Error al iniciar sesión"
    pass.SetFocus
    Exit Sub
End If
usuario = user.List(user.ListIndex)
BuscarU usuario  'llama la funcion de busqueda del usuario

If pass = passw Then
    ConexionBD acceso, "select * from log"
    On Error Resume Next
    bd.Recordset.AddNew
    bd.Recordset!usuario = usuario
    bd.Recordset!fecha = Format(Date, "dd/mm/yyyy")
    bd.Recordset!hora = Format(Time, "hh:mm:ss ampm")
    bd.Recordset.Update
    bd.Recordset.Close
    bd.Refresh
    If Err.Number <> 0 Then
        MsgBox "La base de datos no responde!", vbCritical, "Error"
    End If
    Unload FONDO
    Unload acceso
    menu.Show
ElseIf pass <> passw Then
    MsgBox "La contraseña es incorrecta!", vbCritical, "Inicio de sesión"
    pass.SetFocus
    Exit Sub
End If
End Sub


Private Sub ayuda_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "acceso.htm"
End Sub

Private Sub cancelar_Click()
End
End Sub

Private Sub cerrar_Click()
End
End Sub

Private Sub Form_Activate()
On Error Resume Next
ConexionBD acceso, "select * from acceso"
CargarUsuarios
user.SetFocus
End Sub
Private Sub Form_Initialize()
InitCommonControls
End Sub

'verifica que el programa no se ejecute mas de una vez
Private Sub instancia_Timer()
If App.PrevInstance Then
    MsgBox "El programa ya se está ejecutando", vbInformation, "Jardín"
    End
    End
End If
End Sub


Private Sub pass_GotFocus()
pass.SelStart = 0
pass.SelLength = Len(pass.Text)
End Sub

Private Sub pass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    aceptar_Click
End If
End Sub

Private Sub user_Click()
pass.SetFocus
End Sub

