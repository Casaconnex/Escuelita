VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form ingresos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingresos"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5550
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin Jardin.xpgroupbox xpgroupbox1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6588
      Caption         =   "Parametros"
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
      Begin MSAdodcLib.Adodc bd2 
         Height          =   330
         Left            =   3720
         Top             =   1200
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
      Begin Jardin.xphelp ayuda 
         Height          =   315
         Left            =   4920
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin MSAdodcLib.Adodc bd1 
         Height          =   330
         Left            =   2160
         Top             =   1200
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
      Begin VB.TextBox dato 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1680
         Width           =   3735
      End
      Begin VB.ListBox Lista 
         Height          =   1035
         ItemData        =   "ingresos.frx":0000
         Left            =   240
         List            =   "ingresos.frx":0002
         TabIndex        =   9
         Top             =   2400
         Width           =   3495
      End
      Begin VB.ComboBox parametro 
         Height          =   315
         ItemData        =   "ingresos.frx":0004
         Left            =   240
         List            =   "ingresos.frx":0065
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   4815
      End
      Begin Jardin.xpgroupbox xpgroupbox2 
         Height          =   1215
         Left            =   240
         TabIndex        =   3
         Top             =   3720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2143
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
         Begin MSAdodcLib.Adodc BD 
            Height          =   330
            Left            =   3480
            Top             =   720
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
         Begin JeweledBut.JeweledButton nuevo 
            Height          =   375
            Left            =   120
            TabIndex        =   4
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
            MICON           =   "ingresos.frx":0259
            BC              =   8438015
            FC              =   0
            Picture         =   "ingresos.frx":03C7
         End
         Begin JeweledBut.JeweledButton eliminar 
            Height          =   375
            Left            =   3240
            TabIndex        =   5
            Top             =   240
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
            MICON           =   "ingresos.frx":2ED1
            BC              =   8438015
            FC              =   0
            Picture         =   "ingresos.frx":303F
         End
         Begin JeweledBut.JeweledButton guardar 
            Height          =   375
            Left            =   1680
            TabIndex        =   6
            Top             =   240
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
            MICON           =   "ingresos.frx":35D9
            BC              =   8438015
            FC              =   0
            Picture         =   "ingresos.frx":3747
         End
         Begin JeweledBut.JeweledButton salir 
            Height          =   375
            Left            =   1680
            TabIndex        =   8
            Top             =   720
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
            MICON           =   "ingresos.frx":9375
            BC              =   8438015
            FC              =   0
            Picture         =   "ingresos.frx":94E3
         End
      End
      Begin VB.Label num 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1320
         TabIndex        =   13
         Top             =   2160
         Width           =   60
      End
      Begin VB.Label des 
         AutoSize        =   -1  'True
         Caption         =   "Elemento(s) en la lista."
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   2160
         Width           =   1980
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contenido:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese el dato:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Escoja el parametro que desee personalizar:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3855
      End
   End
End
Attribute VB_Name = "ingresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ayuda_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "ingresos.htm"
End Sub

Private Sub bd_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub bd1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub dato_GotFocus()
dato.SelStart = 0
dato.SelLength = Len(dato.Text)
End Sub

Private Sub eliminar_Click()
If parametro.ListIndex = 0 Then
    MsgBox "Escoja un parametro para que el sistema pueda" & vbCrLf & "ejecutar esta acción!", vbInformation, "Ingresos de Parametros"
    Exit Sub
Else
    Dim tipo As Integer
    tipo = parametro.ListIndex
    If lista.ListIndex = -1 Then
        MsgBox "Escoja un dato de la lista para que el sistema pueda" & vbCrLf & "esta acción!", vbInformation, "Eliminar Datos"
        Exit Sub
    Else
        On Error Resume Next
        ConexionBD1 ingresos, "delete from parametrizacion where tippar=" & tipo & " and dato='" & lista.List(lista.ListIndex) & "'"
        parametro_Click
    End If
End If

End Sub

Private Sub Form_Activate()
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height
parametro.ListIndex = 0
ConexionBD ingresos, "select * from parametrizacion"
menu.estado.Panels(4).Text = "Ingresos de Parametros"
FormularioActivo = True
End Sub

Private Sub Form_Initialize()
InitCommonControls
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
If dato.Text = "" Then
    MsgBox "Falta por ingresar el dato!", vbExclamation, "Ingresos de parámetros"
    dato.SetFocus
    Exit Sub
End If
'verifica si el item ya esta en la lista
ConexionBD2 ingresos, "select * from parametrizacion where tippar=" & parametro.ListIndex & "and dato='" & Trim$(dato.Text) & "'"
If bd2.Recordset.RecordCount = 0 Then
    bd.Recordset.AddNew
    bd.Recordset!tippar = parametro.ListIndex
    bd.Recordset!dato = Trim$(dato.Text)
    bd.Recordset.Update
    guardar.Enabled = False
    nuevo.Enabled = True
    dato.Text = ""
    dato.Enabled = False
    bd.Refresh
    parametro_Click
Else
    MsgBox "No es posible realizar esta acción. El dato: '" & dato.Text & "' en el parametro: '" & parametro.List(parametro.ListIndex) & "' ya existe!", vbExclamation, "Ingresos"
    dato.SetFocus
End If
End Sub

Private Sub nuevo_Click()
If parametro.ListIndex = 0 Then
    MsgBox "Escoja un parametro para que el sistema pueda" & vbCrLf & "ejecutar esta acción!", vbInformation, "Ingresos de Parametros"
    Exit Sub
Else
    dato.Enabled = True
    dato.Text = ""
    dato.SetFocus
    guardar.Enabled = True
    nuevo.Enabled = False
End If
End Sub

Private Sub parametro_Click()
Dim tipo As Integer
tipo = parametro.ListIndex
Select Case parametro.ListIndex
    Case 0: num.Visible = False
            des.Visible = False
    Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30:
        ConexionBD1 ingresos, "select * from parametrizacion where tippar=" & tipo & " order by dato;"
        lista.Clear
        If BD1.Recordset.RecordCount > 0 Then
            num.Caption = BD1.Recordset.RecordCount
            num.Visible = True
            des.Visible = True
            BD1.Recordset.MoveFirst
            For i = 1 To BD1.Recordset.RecordCount
                lista.AddItem BD1.Recordset!dato
                BD1.Recordset.MoveNext
            Next i
        Else
            num.Visible = False
            des.Visible = False
        End If
End Select
End Sub

Private Sub salir_Click()
Unload Me
End Sub

