VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form listespera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado De Espera"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   10455
   Begin MSAdodcLib.Adodc bd1 
      Height          =   330
      Left            =   7200
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
         Size            =   9.75
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
      Left            =   9960
      Top             =   4200
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin Jardin.xpgroupbox frame 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8916
      Caption         =   "Datos Personales"
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
      Begin VB.TextBox eda 
         Height          =   270
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   54
         Top             =   2040
         Width           =   735
      End
      Begin VB.ComboBox edaano 
         Height          =   315
         ItemData        =   "Listado.frx":0442
         Left            =   2040
         List            =   "Listado.frx":0458
         Locked          =   -1  'True
         TabIndex        =   53
         Tag             =   "1"
         Top             =   2400
         Width           =   735
      End
      Begin VB.ComboBox edames 
         Height          =   315
         ItemData        =   "Listado.frx":046E
         Left            =   2880
         List            =   "Listado.frx":0496
         Locked          =   -1  'True
         TabIndex        =   52
         Tag             =   "1"
         Top             =   2400
         Width           =   735
      End
      Begin VB.ComboBox bar 
         Height          =   315
         ItemData        =   "Listado.frx":04C1
         Left            =   6480
         List            =   "Listado.frx":0504
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2400
         Width           =   1815
      End
      Begin VB.ComboBox sec 
         Height          =   315
         ItemData        =   "Listado.frx":05FB
         Left            =   2040
         List            =   "Listado.frx":0605
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   3600
         Width           =   855
      End
      Begin VB.ComboBox nivel 
         Height          =   315
         ItemData        =   "Listado.frx":0611
         Left            =   6480
         List            =   "Listado.frx":0621
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   3480
         Width           =   855
      End
      Begin Jardin.xpgroupbox xpgroupbox1 
         Height          =   855
         Left            =   120
         TabIndex        =   32
         Top             =   4080
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1508
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
            TabIndex        =   33
            ToolTipText     =   "Primer Registro"
            Top             =   360
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
            MICON           =   "Listado.frx":0631
            BC              =   8438015
            FC              =   0
            Picture         =   "Listado.frx":079F
         End
         Begin JeweledBut.JeweledButton siguiente 
            Height          =   375
            Left            =   3480
            TabIndex        =   34
            ToolTipText     =   "Siguiente Registro"
            Top             =   360
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
            MICON           =   "Listado.frx":08F9
            BC              =   8438015
            FC              =   0
            Picture         =   "Listado.frx":0A67
         End
         Begin JeweledBut.JeweledButton ultimo 
            Height          =   375
            Left            =   5160
            TabIndex        =   35
            ToolTipText     =   "Ultimo Registro"
            Top             =   360
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
            MICON           =   "Listado.frx":0BC1
            BC              =   8438015
            FC              =   0
            Picture         =   "Listado.frx":0D2F
         End
         Begin JeweledBut.JeweledButton anterior 
            Height          =   375
            Left            =   1800
            TabIndex        =   37
            ToolTipText     =   "Anterior Registro"
            Top             =   360
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
            MICON           =   "Listado.frx":0E89
            BC              =   8438015
            FC              =   0
            Picture         =   "Listado.frx":0FF7
         End
      End
      Begin MSComCtl2.DTPicker fecnac 
         Height          =   375
         Left            =   6480
         TabIndex        =   31
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   20185089
         UpDown          =   -1  'True
         CurrentDate     =   38146.4711574074
         MaxDate         =   401768
         MinDate         =   2
      End
      Begin VB.ComboBox parfam 
         Height          =   315
         ItemData        =   "Listado.frx":1151
         Left            =   6480
         List            =   "Listado.frx":116A
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox numdoc 
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox priape 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox segape 
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox prinom 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox segnom 
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox dir 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   7
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox tel 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3240
         Width           =   1335
      End
      Begin VB.ComboBox sal 
         Height          =   315
         ItemData        =   "Listado.frx":11C4
         Left            =   6480
         List            =   "Listado.frx":11D4
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ComboBox sex 
         Height          =   315
         ItemData        =   "Listado.frx":11F2
         Left            =   2040
         List            =   "Listado.frx":11FC
         Locked          =   -1  'True
         TabIndex        =   4
         Tag             =   "1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox tipdoc 
         Height          =   315
         ItemData        =   "Listado.frx":1206
         Left            =   2040
         List            =   "Listado.frx":1219
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox niv 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   6240
         Width           =   1815
      End
      Begin VB.ComboBox secsal 
         Height          =   315
         ItemData        =   "Listado.frx":124B
         Left            =   6480
         List            =   "Listado.frx":1255
         TabIndex        =   1
         Top             =   6240
         Width           =   1815
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   3000
         TabIndex        =   60
         Top             =   1680
         Width           =   60
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2880
         TabIndex        =   58
         Top             =   2040
         Width           =   60
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Edad Aprox. (años/meses)"
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   120
         TabIndex        =   51
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Afiliado Sec. Salud"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   3720
         Width           =   1605
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel de Seguridad Social"
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   4440
         TabIndex        =   45
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   1350
         Left            =   7200
         Picture         =   "Listado.frx":1261
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   1125
      End
      Begin VB.Label Label71 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "No Documento"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Segundo apellido"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4440
         TabIndex        =   28
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Segundo Nombre"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4440
         TabIndex        =   27
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label Label76 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Edad"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label78 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label81 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Seguridad Social en Salud"
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   4440
         TabIndex        =   23
         Top             =   2880
         Width           =   2010
      End
      Begin VB.Label Label85 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Parentesco Familiar"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4440
         TabIndex        =   21
         Top             =   2040
         Width           =   1680
      End
      Begin VB.Label Label86 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Nacimiento"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4440
         TabIndex        =   20
         Top             =   1560
         Width           =   1770
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Barrio"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4440
         TabIndex        =   19
         Top             =   2520
         Width           =   525
      End
      Begin VB.Label Label74 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Primer Nombre"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label72 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Primer Apellido"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label70 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Doc"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   6240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Secretaria de Salud"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Top             =   6240
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc bd 
      Height          =   330
      Left            =   8880
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
         Size            =   9.75
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
      Left            =   8760
      TabIndex        =   36
      Top             =   5520
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
      MICON           =   "Listado.frx":3BF6
      BC              =   8438015
      FC              =   0
      Picture         =   "Listado.frx":3D64
   End
   Begin Jardin.xpgroupbox xpgroupbox2 
      Height          =   3615
      Left            =   8640
      TabIndex        =   38
      Top             =   480
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
         TabIndex        =   39
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
         MICON           =   "Listado.frx":3EBE
         BC              =   8438015
         FC              =   0
         Picture         =   "Listado.frx":402C
      End
      Begin JeweledBut.JeweledButton busqueda 
         Height          =   375
         Left            =   120
         TabIndex        =   40
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
         MICON           =   "Listado.frx":6B36
         BC              =   8438015
         FC              =   0
         Picture         =   "Listado.frx":6CA4
      End
      Begin JeweledBut.JeweledButton eliminar 
         Height          =   375
         Left            =   120
         TabIndex        =   41
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
         MICON           =   "Listado.frx":6DFE
         BC              =   8438015
         FC              =   0
         Picture         =   "Listado.frx":6F6C
      End
      Begin JeweledBut.JeweledButton guardar 
         Height          =   375
         Left            =   120
         TabIndex        =   42
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
         MICON           =   "Listado.frx":7506
         BC              =   8438015
         FC              =   0
         Picture         =   "Listado.frx":7674
      End
      Begin JeweledBut.JeweledButton Actualizar 
         Height          =   375
         Left            =   120
         TabIndex        =   43
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
         MICON           =   "Listado.frx":D2A2
         BC              =   8438015
         FC              =   0
         Picture         =   "Listado.frx":D410
      End
      Begin JeweledBut.JeweledButton modificar 
         Height          =   375
         Left            =   120
         TabIndex        =   44
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
         MICON           =   "Listado.frx":D9AA
         BC              =   8438015
         FC              =   0
         Picture         =   "Listado.frx":DB18
      End
      Begin JeweledBut.JeweledButton parametro 
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   3120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         TX              =   "Parámetros"
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
         MICON           =   "Listado.frx":DC72
         BC              =   8438015
         FC              =   0
      End
   End
   Begin JeweledBut.JeweledButton cancelar 
      Height          =   375
      Left            =   8760
      TabIndex        =   59
      Top             =   5040
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
      MICON           =   "Listado.frx":DDE0
      BC              =   8438015
      FC              =   0
      Picture         =   "Listado.frx":DF4E
   End
   Begin VB.Label fecins 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2040
      TabIndex        =   56
      Top             =   120
      Width           =   60
   End
   Begin VB.Label numreg 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8760
      TabIndex        =   55
      Top             =   4680
      Width           =   60
   End
   Begin VB.Label total 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   8760
      TabIndex        =   48
      Top             =   4320
      Width           =   60
   End
   Begin VB.Label Label69 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Inscripción"
      ForeColor       =   &H80000006&
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "listespera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Actualizar_Click()
On Error Resume Next
nuevo.Enabled = True
modificar.Enabled = True
parametro.Enabled = True
eliminar.Enabled = True
primero.Enabled = True
ultimo.Enabled = True
siguiente.Enabled = True
anterior.Enabled = True
Actualizar.Enabled = False
busqueda.Enabled = True
tipdoc.Locked = True
numdoc.Locked = True
priape.Locked = True
segape.Locked = True
prinom.Locked = True
segnom.Locked = True
sex.Locked = True
eda.Locked = True
parfam.Locked = True
dir.Locked = True
bar.Locked = True
tel.Locked = True
sal.Locked = True
sec.Locked = True
nivel.Locked = True
If ModificadoL = True Then
    MODIFICARR
    ModificadoL = False
End If
End Sub
Sub MODIFICARR()
On Error Resume Next
'modifica el REGISTRO EN INSCRIPCIONES
'ConexionBD listespera, "select * from listadodeespera where numdoc='" & numdoc.Text & "'"
Dim modii
modii = bd.Recordset.EditMode
bd.Recordset!fecins = fecins.Caption
bd.Recordset!tipdoc = tipdoc.Text
bd.Recordset!numdoc = numdoc.Text
bd.Recordset!priape = priape.Text
bd.Recordset!segape = segape.Text
bd.Recordset!prinom = prinom.Text
bd.Recordset!segnom = segnom.Text
bd.Recordset!sex = sex.Text
bd.Recordset!fecnac = fecnac.Value
bd.Recordset!eda = eda.Text
bd.Recordset!edaano = edaano.Text
bd.Recordset!edames = edames.Text
'calcula los meses de edad del niño

If Not IsNull(fecnac.Value) Then 'si metio la fecha
    bd.Recordset!meses = (Val(Format(Date, "yyyy") - fecnac.Year) * 12) + (Val(Format(Date, "mm") - fecnac.Month))
Else  ' si no metio la fecha
    bd.Recordset!meses = (Val(edaano.Text) * 12) + Val(edames.Text)
End If

bd.Recordset!parfam = parfam.Text
bd.Recordset!dir = dir.Text
bd.Recordset!tel = tel.Text
bd.Recordset!bar = bar.Text
bd.Recordset!sal = sal.Text
bd.Recordset!niv = nivel.Text
bd.Recordset!secrsal = sec.Text
bd.Recordset.Update


End Sub
Private Sub anterior_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MovePrevious
    If bd.Recordset.BOF Then
        bd.Recordset.MoveFirst
    End If
    mostrarcampos
End If

End Sub

Private Sub bar_GotFocus()
ObjetoActual = bar.Name
End Sub

Private Sub bar_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    dir.SetFocus
End If
End Sub

Private Sub bd1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub busqueda_Click()
'motor de busqueda
MB.Formulario = Me.Name
MB.Descripcion = "Listado de Espera"
elBuscador.Show
End Sub

Private Sub cancelar_Click()
If bd.Recordset.RecordCount > 0 Then
    mostrarcampos
End If

If NuevoRegL = True Then
    nuevo.Enabled = True
    modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    anterior.Enabled = True
    siguiente.Enabled = True
    ultimo.Enabled = True
    guardar.Enabled = False
    busqueda.Enabled = True
    NuevoRegL = False
ElseIf ModificadoL = True Then
    nuevo.Enabled = True
    modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    ultimo.Enabled = True
    siguiente.Enabled = True
    anterior.Enabled = True
    Actualizar.Enabled = False
    busqueda.Enabled = True
    ModificadoL = False
End If
parametro.Enabled = True
End Sub

Private Sub dir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        sal.SetFocus
End If
End Sub




Private Sub eda_GotFocus()
If IsNull(fecnac.Value) Then
    eda.Enabled = False
    eda.Text = ""
    edaano.Enabled = True
    edames.Enabled = True
End If
End Sub

Private Sub eda_KeyPress(KeyAscii As Integer)

KeyAscii = Validar_numero(KeyAscii)
End Sub

Private Sub edaano_GotFocus()
If Not IsNull(fecnac.Value) Then
    eda.Enabled = True
    edaano.Enabled = False
    edames.Enabled = False
    edaano.Text = ""
    edames.Text = ""
End If
End Sub

Private Sub edaano_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    edames.SetFocus
End If
End Sub

Private Sub edames_GotFocus()
If Not IsNull(fecnac.Value) Then
    eda.Enabled = True
    edaano.Enabled = False
    edames.Enabled = False
    edaano.Text = ""
    edames.Text = ""
End If
End Sub

Private Sub edames_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    parfam.SetFocus
End If
End Sub

Private Sub eliminar_Click()
If bd.Recordset.RecordCount > 0 Then
If MsgBox("Está seguro de querer eliminar este registro?", vbYesNo + vbQuestion, "Eliminar Registro") = vbYes Then
    bd.Recordset.Delete
    If bd.Recordset.RecordCount > 0 Then
    bd.Recordset.MoveFirst
    bd.Refresh
    total = bd.Recordset.RecordCount & " en espera."
    mostrarcampos
    Else
        Unload Me
        listespera.Show
   End If
End If
End If
End Sub

Private Sub fecnac_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Calculo
End If
End Sub
Sub Calculo()
If fecnac.Value <> "" Then
        Dim meses, ANIOs
        Dim ed As Integer
        'meses = Val(Format(Date, "yyyy") - fecnac.Month)
        If Val(Format(Date, "yyyy")) > fecnac.Year Then
            ANIOs = Val(Format(Date, "yyyy") - fecnac.Year)
            meses = ANIOs * 12
            ed = 1
        Else
            meses = Val(Format(Date, "mm")) - fecnac.Month
            ed = 2
        End If
        
        If meses >= 3 And meses <= 60 Then
            If ed = 1 Then
                eda.Text = ANIOs
                Label6 = "años"
            ElseIf ed = 2 Then
                eda.Text = meses
                Label6 = "meses"
            End If
            
        Else
            MsgBox "Edad no permitida para ingresar al listado de Espera", vbInformation, "Listado de Espera"
            Exit Sub
        End If
    
        parfam.SetFocus
    Else
        edaano.SetFocus
    End If
End Sub
Private Sub LlenarCombos()
menu.estado.Panels(4).Text = "Cargando..."
'llenar tipo documento del niño
ConexionBD1 listespera, "select * from parametrizacion where tippar=6" & " order by dato;"
tipdoc.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        tipdoc.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar barrio
ConexionBD1 listespera, "select * from parametrizacion where tippar=16" & " order by dato;"
bar.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        bar.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
BD1.Recordset.Close
Set BD1.Recordset = Nothing
End Sub



Private Sub fecnac_LostFocus()
Calculo
End Sub

Private Sub Form_Activate()
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height
total = bd.Recordset.RecordCount & " en espera."
If ConsultaF = False Then
    FormularioActivo = True
End If
If Me.Tag <> "" Then
    bd.Recordset.MoveFirst
    For i = 1 To bd.Recordset.RecordCount
        If bd.Recordset!numdoc = Me.Tag Then
            mostrarcampos
            Exit For
        End If
        bd.Recordset.MoveNext
    Next i
End If
fecins = Format(Date, "dd/mm/yyyy")
tipdoc.SetFocus
If Para = True Then
    LlenarCombos
    Para = False
    If bd.Recordset.RecordCount > 0 Then
        mostrarcampos
    End If
End If

End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
LlenarCombos
menu.estado.Panels(4).Text = "Listado de Espera de los niños"
On Error Resume Next
ConexionBD listespera, "select * from listadodeespera where inscrito=0"
If bd.Recordset.RecordCount > 0 Then
    mostrarcampos
End If

End Sub
Function mostrarcampos()
numreg = bd.Recordset.AbsolutePosition & " registro."
fecins = bd.Recordset!fecins
tipdoc = bd.Recordset!tipdoc
numdoc = bd.Recordset!numdoc
priape = bd.Recordset!priape
segape = bd.Recordset!segape
prinom = bd.Recordset!prinom
segnom = bd.Recordset!segnom
sex = bd.Recordset!sex
If IsNull(bd.Recordset!fecnac) = False Then
    fecnac = bd.Recordset!fecnac
    fecnac.Enabled = True
Else
    fecnac.Enabled = False
End If
eda = bd.Recordset!eda
If IsNull(bd.Recordset!edaano) = False Then
    edaano.Text = bd.Recordset!edaano
End If
If IsNull(bd.Recordset!edames) = False Then
    edames.Text = bd.Recordset!edames
End If
parfam = bd.Recordset!parfam
dir = bd.Recordset!dir
bar = bd.Recordset!bar
tel = bd.Recordset!tel
sal = bd.Recordset!sal
sec = bd.Recordset!secrsal
nivel = bd.Recordset!niv
fecnac.Enabled = False
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If NuevoRegL = True Then
    If MsgBox("Esta agregando un nuevo registro" & vbCrLf & "Desea continuar?", vbYesNo + vbQuestion, "Listado de Espera") = vbYes Then
        Cancel = True
    Else
        NuevoRegL = False
    End If
End If
End Sub

Private Sub Form_Resize()
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)
FormularioActivo = False
menu.estado.Panels(4).Text = "Menú Principal"
End Sub

Private Sub guardar_Click()

'validar campos
If tipdoc.Text = "" Or numdoc.Text = "" Or priape.Text = "" Or _
prinom.Text = "" Or sex.Text = "" Or parfam.Text = "" Or dir.Text = "" Or bar.Text = "" Or tel.Text = "" Or sal.Text = "" Then
    MsgBox "Hace falta datos por ingresar!", vbInformation, "Lista de Espera"
    Exit Sub
End If
If fecnac.Value = "" Then
    If edaano.Text = "" Or edames.Text = "" Then
        MsgBox "Ha seleccionado que no sabe la fecha de nacimiento del niño." & vbCrLf & "Por favor, ingrese la edad aproximada!", vbInformation, "Listado de Espera"
        edaano.SetFocus
        Exit Sub
    End If
ElseIf fecnac.Value <> "" Then
    If eda.Text = "" Then
        MsgBox "Ha seleccionado que sabe la fecha de nacimiento del niño." & vbCrLf & "Por favor, ingrese la edad de este!", vbInformation, "Listado de Espera"
        eda.SetFocus
        Exit Sub
    End If
End If

Deshabilitarl listespera


nuevo.Enabled = True
modificar.Enabled = True
eliminar.Enabled = True
primero.Enabled = True
anterior.Enabled = True
siguiente.Enabled = True
parametro.Enabled = True
ultimo.Enabled = True
guardar.Enabled = False

guardarregistro
busqueda.Enabled = True
NuevoRegL = False
bd.Refresh
total = bd.Recordset.RecordCount & " en espera."
End Sub

Private Sub modificar_Click()
If bd.Recordset.RecordCount > 0 Then
ModificadoL = True
modificar.Enabled = False
Actualizar.Enabled = True
nuevo.Enabled = False
eliminar.Enabled = False
fecnac.Enabled = True
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False
parametro.Enabled = False
busqueda.Enabled = False
tipdoc.Locked = False
numdoc.Locked = False
priape.Locked = False
segape.Locked = False
prinom.Locked = False
segnom.Locked = False
sex.Locked = False
eda.Locked = False
parfam.Locked = False
dir.Locked = False
bar.Locked = False
tel.Locked = False
sal.Locked = False
sec.Locked = False
nivel.Locked = False
fecnac.Enabled = True

'bd.Recordset.AddNew
End If
End Sub

Private Sub nivel_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    sec.SetFocus
End If
End Sub

Private Sub nuevo_Click()
NuevoRegL = True
cajasl listespera 'habilita cajas para meter datos
nuevo.Enabled = False
guardar.Enabled = True
eliminar.Enabled = False
modificar.Enabled = False
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False
parametro.Enabled = False
fecins = Format(Date, "dd/mm/yyyy")
fecnac.Enabled = True
tipdoc.SetFocus
sex.Text = ""
busqueda.Enabled = False
'crear nuevo registro
bd.Refresh

End Sub
Private Sub numdoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'CONSULTA PARA VER SI YA ESTA UN NUMDOC INSCRITO
    ConexionBD1 listespera, "select * from listadodeespera where numdoc='" & numdoc.Text & "'"
    If BD1.Recordset.RecordCount > 0 Then
        MsgBox "El niño con documento: " & numdoc.Text & " ya está en listado de espera!", vbInformation, "Incripciones"
        Exit Sub
    End If
    priape.SetFocus
End If
End Sub

Private Sub parametro_Click()
'guarda el objeto actual y muestra el fomrulariomde paramaterizacióm
Para = True
ingresos.Tag = ObjetoActual
ingresos.Show
End Sub

Private Sub parfam_GotFocus()
ObjetoActual = parfam.Name
End Sub

Private Sub parfam_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    If fecnac.Value = Null Then
        edaano.SetFocus
    Else
        bar.SetFocus
    End If
End If
End Sub

Private Sub priape_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    segape.SetFocus
End If
End Sub

Private Sub primero_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveFirst
    mostrarcampos
End If
End Sub

Private Sub prinom_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    segnom.SetFocus
End If
End Sub

Private Sub sal_GotFocus()
ObjetoActual = sal.Name
End Sub

Private Sub sal_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    tel.SetFocus
End If
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Private Sub sec_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)

End Sub

Private Sub segape_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    prinom.SetFocus
End If
End Sub

Private Sub segnom_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    sex.SetFocus
End If
End Sub

Private Sub sex_Click()
If sex.Text = "M" Then
    Label7 = "Masculino"
ElseIf sex.Text = "F" Then
    Label7 = "Femenino"
End If
End Sub

Private Sub sex_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    fecnac.SetFocus
End If
End Sub
Private Sub siguiente_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveNext
    If bd.Recordset.EOF Then
        bd.Recordset.MoveLast
    End If
    mostrarcampos
End If

End Sub
Private Sub tel_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    nivel.SetFocus
End If
End Sub

Private Sub tipdoc_GotFocus()
ObjetoActual = tipdoc.Name
End Sub

Private Sub tipdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    numdoc.SetFocus
End If
End Sub
Private Sub ultimo_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveLast
    mostrarcampos
End If
End Sub
Function guardarregistro()
On Error Resume Next
bd.Recordset.AddNew
bd.Recordset!fecins = fecins.Caption
bd.Recordset!tipdoc = tipdoc.Text
bd.Recordset!numdoc = numdoc.Text
bd.Recordset!priape = priape.Text
bd.Recordset!segape = segape.Text
bd.Recordset!prinom = prinom.Text
bd.Recordset!segnom = segnom.Text
bd.Recordset!sex = sex.Text
bd.Recordset!fecnac = fecnac.Value
bd.Recordset!eda = eda.Text
bd.Recordset!edaano = edaano.Text
bd.Recordset!edames = edames.Text
'calcula los meses de edad del niño

If Not IsNull(fecnac.Value) Then 'si metio la fecha
    bd.Recordset!meses = (Val(Format(Date, "yyyy") - fecnac.Year) * 12) + (Val(Format(Date, "mm") - fecnac.Month))
Else  ' si no metio la fecha
    bd.Recordset!meses = (Val(edaano.Text) * 12) + Val(edames.Text)
End If

bd.Recordset!parfam = parfam.Text
bd.Recordset!dir = dir.Text
bd.Recordset!tel = tel.Text
bd.Recordset!bar = bar.Text
bd.Recordset!sal = sal.Text
bd.Recordset!niv = nivel.Text
bd.Recordset!secrsal = sec.Text
bd.Recordset.Update
bd.Refresh

End Function

Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "listadoespera.htm"
End Sub
