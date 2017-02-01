VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form reportes 
   Caption         =   "Generador de Reportes"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   -6315
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
   Icon            =   "Reportes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin Jardin.xpgroupbox xpgroupbox1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13996
      Caption         =   "Reportes personalizados de SISJACE"
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
      Begin VB.Frame Frame1 
         Caption         =   "Compras"
         Height          =   2175
         Left            =   7680
         TabIndex        =   38
         Top             =   2640
         Width           =   3495
         Begin MSComCtl2.DTPicker fecc 
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50200577
            CurrentDate     =   38181
         End
         Begin VB.ComboBox lcompras 
            Height          =   315
            ItemData        =   "Reportes.frx":014A
            Left            =   240
            List            =   "Reportes.frx":015D
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   360
            Width           =   3015
         End
         Begin JeweledBut.JeweledButton bcom 
            Height          =   495
            Left            =   240
            TabIndex        =   42
            Top             =   1560
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   873
            TX              =   "Ver Reporte"
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
            MICON           =   "Reportes.frx":01AA
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":0318
         End
         Begin VB.TextBox datoc 
            Height          =   285
            Left            =   240
            TabIndex        =   41
            Top             =   960
            Visible         =   0   'False
            Width           =   2655
         End
      End
      Begin Jardin.xpgroupbox grade 
         Height          =   1815
         Left            =   3960
         TabIndex        =   33
         Top             =   4920
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3201
         Caption         =   "Listado de niños por grado y nivel"
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
         Begin JeweledBut.JeweledButton ver 
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
            TX              =   ">>"
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
            MICON           =   "Reportes.frx":0472
            BC              =   8438015
            FC              =   0
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "De click en botón para escoger el grado y nivel."
            Height          =   435
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   2445
         End
      End
      Begin Jardin.xpgroupbox xpgroupbox8 
         Height          =   2055
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   3625
         Caption         =   "Listado de Espera"
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
         Begin VB.TextBox numdocle 
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   53
            Top             =   1080
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.ComboBox filtrole 
            Height          =   315
            ItemData        =   "Reportes.frx":05E0
            Left            =   120
            List            =   "Reportes.frx":05ED
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   360
            Width           =   3375
         End
         Begin MSAdodcLib.Adodc BD 
            Height          =   330
            Left            =   3000
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
         Begin MSComCtl2.DTPicker fecins 
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50200577
            CurrentDate     =   38177
         End
         Begin JeweledBut.JeweledButton listado 
            Height          =   495
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Visible         =   0   'False
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   873
            TX              =   "Imprimir Listado de Espera"
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
            MICON           =   "Reportes.frx":0631
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":079F
         End
         Begin JeweledBut.JeweledButton vplis 
            Height          =   495
            Left            =   3000
            TabIndex        =   45
            ToolTipText     =   "Ver listado de niños inscritos"
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            TX              =   ""
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
            MICON           =   "Reportes.frx":0D39
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":0EA7
         End
         Begin JeweledBut.JeweledButton listadoe 
            Height          =   495
            Left            =   120
            TabIndex        =   46
            ToolTipText     =   "Ver listado de niños en espera"
            Top             =   1440
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   873
            TX              =   "Ver Reporte"
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
            MICON           =   "Reportes.frx":1001
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":101D
         End
      End
      Begin Jardin.xphelp ayuda 
         Height          =   315
         Left            =   9600
         Top             =   5520
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin Jardin.xpgroupbox xpgroupbox3 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   3625
         Caption         =   "Inscripciones"
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
         Begin VB.TextBox fecf 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   55
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox filtroins 
            Height          =   315
            ItemData        =   "Reportes.frx":1177
            Left            =   120
            List            =   "Reportes.frx":1184
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   360
            Width           =   3375
         End
         Begin MSAdodcLib.Adodc BD1 
            Height          =   330
            Left            =   3360
            Top             =   0
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
         Begin VB.ComboBox numdocins 
            Height          =   315
            ItemData        =   "Reportes.frx":11CF
            Left            =   120
            List            =   "Reportes.frx":11D1
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   960
            Visible         =   0   'False
            Width           =   2295
         End
         Begin JeweledBut.JeweledButton reportinscripciones 
            Height          =   495
            Left            =   120
            TabIndex        =   5
            Top             =   1440
            Visible         =   0   'False
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   873
            TX              =   "Imprimir Hoja Inscripción"
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
            MICON           =   "Reportes.frx":11D3
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":1341
         End
         Begin JeweledBut.JeweledButton listain 
            Height          =   495
            Left            =   3000
            TabIndex        =   43
            ToolTipText     =   "Ver listado de niños inscritos"
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            TX              =   ""
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
            MICON           =   "Reportes.frx":18DB
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":1A49
         End
         Begin JeweledBut.JeweledButton listadoi 
            Height          =   495
            Left            =   120
            TabIndex        =   47
            ToolTipText     =   "Ver listado de niños inscritos"
            Top             =   1440
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   873
            TX              =   "Ver Reporte"
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
            MICON           =   "Reportes.frx":1BA3
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":1BBF
         End
      End
      Begin Jardin.xpgroupbox xpgroupbox4 
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   4680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2990
         Caption         =   "Matriculas"
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
         Begin VB.ComboBox ANIO 
            Height          =   315
            ItemData        =   "Reportes.frx":1D19
            Left            =   120
            List            =   "Reportes.frx":1D3E
            TabIndex        =   49
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin MSAdodcLib.Adodc bd3 
            Height          =   330
            Left            =   1080
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
         Begin MSAdodcLib.Adodc bd2 
            Height          =   330
            Left            =   1080
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
         Begin VB.ComboBox listamat 
            Height          =   315
            ItemData        =   "Reportes.frx":1D84
            Left            =   120
            List            =   "Reportes.frx":1D91
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   480
            Width           =   2175
         End
         Begin JeweledBut.JeweledButton listama 
            Height          =   495
            Left            =   3000
            TabIndex        =   44
            ToolTipText     =   "Ver listado de niños matriculados"
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            TX              =   ""
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
            MICON           =   "Reportes.frx":1DCC
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":1F3A
         End
         Begin JeweledBut.JeweledButton listadom 
            Height          =   495
            Left            =   2520
            TabIndex        =   48
            ToolTipText     =   "Ver lista de niños matriculados"
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            TX              =   ""
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
            MICON           =   "Reportes.frx":2094
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":20B0
         End
         Begin JeweledBut.JeweledButton reportanio 
            Height          =   375
            Left            =   1320
            TabIndex        =   50
            ToolTipText     =   "Ver listado de niños matriculados"
            Top             =   960
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            TX              =   "Ver Reporte"
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
            MICON           =   "Reportes.frx":220A
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":2378
         End
         Begin JeweledBut.JeweledButton reportmatriculas 
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   873
            TX              =   "Imprimir Registro de Matricula"
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
            MICON           =   "Reportes.frx":24D2
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":2640
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Número de Documento:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   2055
         End
      End
      Begin Jardin.xpgroupbox xpgroupbox5 
         Height          =   2175
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3836
         Caption         =   "Empleados"
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
         Begin MSComCtl2.DTPicker fecvin 
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50200577
            CurrentDate     =   38181
         End
         Begin VB.TextBox docuemp 
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.ComboBox listemp 
            Height          =   315
            ItemData        =   "Reportes.frx":2BDA
            Left            =   120
            List            =   "Reportes.frx":2BF0
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   3255
         End
         Begin JeweledBut.JeweledButton reportempleados 
            Height          =   495
            Left            =   360
            TabIndex        =   7
            Top             =   1560
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   873
            TX              =   "Ver Reporte"
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
            MICON           =   "Reportes.frx":2C7F
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":2DED
         End
      End
      Begin JeweledBut.JeweledButton salir 
         Height          =   375
         Left            =   9960
         TabIndex        =   4
         Top             =   7080
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
         MICON           =   "Reportes.frx":2F47
         BC              =   8438015
         FC              =   0
         Picture         =   "Reportes.frx":30B5
      End
      Begin Jardin.xpgroupbox xpgroupbox6 
         Height          =   2175
         Left            =   3960
         TabIndex        =   12
         Top             =   2640
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   3836
         Caption         =   "Pagos"
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
         Begin VB.TextBox pnum 
            Height          =   285
            Left            =   240
            TabIndex        =   37
            Top             =   960
            Visible         =   0   'False
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker fecha4 
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50200577
            CurrentDate     =   38154
         End
         Begin VB.ComboBox listpag 
            Height          =   315
            ItemData        =   "Reportes.frx":320F
            Left            =   240
            List            =   "Reportes.frx":3225
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   3135
         End
         Begin JeweledBut.JeweledButton reportpag 
            Height          =   495
            Left            =   480
            TabIndex        =   14
            Top             =   1560
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   873
            TX              =   "Ver Reporte"
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
            MICON           =   "Reportes.frx":328E
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":33FC
         End
         Begin MSComCtl2.DTPicker fechapagos 
            Height          =   375
            Left            =   1920
            TabIndex        =   21
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50200577
            CurrentDate     =   38168
         End
         Begin JeweledBut.JeweledButton reportp 
            Height          =   495
            Left            =   2640
            TabIndex        =   51
            Top             =   840
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            TX              =   ""
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
            MICON           =   "Reportes.frx":3556
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":3572
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "De:"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   1920
            TabIndex        =   22
            Top             =   720
            Visible         =   0   'False
            Width           =   555
         End
      End
      Begin Jardin.xpgroupbox xpgroupbox7 
         Height          =   2175
         Left            =   7680
         TabIndex        =   16
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3836
         Caption         =   "Material Didactico"
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
         Begin VB.TextBox datos 
            Height          =   285
            Left            =   240
            TabIndex        =   20
            Top             =   960
            Visible         =   0   'False
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker fechades 
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50200577
            CurrentDate     =   38154
         End
         Begin VB.ComboBox listamaterial 
            Height          =   315
            ItemData        =   "Reportes.frx":36CC
            Left            =   240
            List            =   "Reportes.frx":36E2
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   360
            Width           =   3135
         End
         Begin JeweledBut.JeweledButton reportmaterial 
            Height          =   495
            Left            =   360
            TabIndex        =   18
            Top             =   1560
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   873
            TX              =   "Ver Reporte"
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
            MICON           =   "Reportes.frx":377F
            BC              =   8438015
            FC              =   0
            Picture         =   "Reportes.frx":38ED
         End
         Begin MSComCtl2.DTPicker fechamat 
            Height          =   375
            Left            =   1920
            TabIndex        =   24
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50200577
            CurrentDate     =   38168
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   1920
            TabIndex        =   26
            Top             =   720
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "De:"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   720
            Visible         =   0   'False
            Width           =   315
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Formato Fecha: dia/mes/año"
         Height          =   195
         Left            =   7440
         TabIndex        =   27
         Top             =   7200
         Width           =   2475
      End
      Begin VB.Image Image1 
         Height          =   1965
         Left            =   9960
         Picture         =   "Reportes.frx":3A47
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   1425
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C5FAFE&
         Caption         =   "El Generador de Reportes le permite escoger el filtro de información según el item e imprimir los resultados."
         Height          =   1215
         Left            =   7800
         TabIndex        =   36
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C5FAFE&
         BackStyle       =   1  'Opaque
         Height          =   1455
         Left            =   7680
         Top             =   5640
         Width           =   2175
      End
   End
   Begin VB.Menu grades 
      Caption         =   "Grados"
      Visible         =   0   'False
      Begin VB.Menu sc 
         Caption         =   "Sala Cuna (3 meses a 1 año)"
         Begin VB.Menu sa 
            Caption         =   "Nivel A"
         End
      End
      Begin VB.Menu cm 
         Caption         =   "Caminadores (1 año 1 mes a 2 años)"
         Begin VB.Menu ca 
            Caption         =   "Nivel A "
         End
         Begin VB.Menu cb 
            Caption         =   "Nivel B"
         End
      End
      Begin VB.Menu pv 
         Caption         =   "Párvulos (2 años 1 mes a 2½ años)"
         Begin VB.Menu pa 
            Caption         =   "Nivel A"
         End
      End
      Begin VB.Menu pk 
         Caption         =   "PreKinder"
         Begin VB.Menu p1 
            Caption         =   "Prekinder (2½ años a 3 años)"
            Begin VB.Menu p1a 
               Caption         =   "Nivel A"
            End
            Begin VB.Menu p1b 
               Caption         =   "Nivel B"
            End
            Begin VB.Menu p1c 
               Caption         =   "Nivel C"
            End
            Begin VB.Menu p1d 
               Caption         =   "Nivel D"
            End
         End
         Begin VB.Menu p2 
            Caption         =   "PreKinder (3 años 1 mes a 4 años)"
            Begin VB.Menu p2a 
               Caption         =   "Nivel A"
            End
            Begin VB.Menu p2b 
               Caption         =   "Nivel B"
            End
            Begin VB.Menu p2c 
               Caption         =   "Nivel C"
            End
            Begin VB.Menu p2d 
               Caption         =   "Nivel D"
            End
         End
      End
      Begin VB.Menu k 
         Caption         =   "Kinder (4 años 1 mes a 5 años)"
         Begin VB.Menu ka 
            Caption         =   "Nivel A"
         End
         Begin VB.Menu kb 
            Caption         =   "Nivel B"
         End
         Begin VB.Menu kc 
            Caption         =   "Nivel C"
         End
      End
   End
   Begin VB.Menu mat 
      Caption         =   "matricula"
      Visible         =   0   'False
      Begin VB.Menu lt 
         Caption         =   "Listado Total"
      End
      Begin VB.Menu la 
         Caption         =   "Listado por años"
      End
   End
   Begin VB.Menu numpagos 
      Caption         =   "numdocpagos"
      Visible         =   0   'False
      Begin VB.Menu pensiones 
         Caption         =   "Pensiones"
      End
      Begin VB.Menu matriculas 
         Caption         =   "Matricula"
      End
      Begin VB.Menu otrosc 
         Caption         =   "Otros Conceptos"
      End
      Begin VB.Menu mt 
         Caption         =   "Mostrar Todos"
      End
   End
End
Attribute VB_Name = "reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'registro de las variables para utilizar con word
Private ListEsp As Word.Application
Private Incripcion As Word.Application
Private matricula As Word.Application

Private Sub ANIO_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
End Sub

Private Sub ayuda_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "reportes.htm"
End Sub

Private Sub bcom_Click()
On Error Resume Next
Select Case lcompras.ListIndex
    Case 0:
        If datoc.Text = "" Then
            MsgBox "No hay ingresado el número de compra para realizar el reporte!", vbInformation, "Reportes"
            datoc.SetFocus
            Exit Sub
        End If
        ConexJardin.rsCompras.Filter = ""
        ConexJardin.rsCompras.Filter = "[numcom]=" & datoc.Text
        LComprasc.Show
    Case 1:
        ConexJardin.rsCompras.Filter = ""
        ConexJardin.rsCompras.Filter = "[fecfac]=" & fecc.Value
        LComprasc.Show
    Case 2:
        If datoc.Text = "" Then
            MsgBox "No hay ingresado el número de compra para realizar el reporte!", vbInformation, "Reportes"
            datoc.SetFocus
            Exit Sub
        End If
        ConexJardin.rsCompras.Filter = ""
        ConexJardin.rsCompras.Filter = "[refmat]=" & datoc.Text
        LComprasc.Show
    Case 3:
        If datoc.Text = "" Then
            MsgBox "No hay ingresado el número de compra para realizar el reporte!", vbInformation, "Reportes"
            datoc.SetFocus
            Exit Sub
        End If
        ConexJardin.rsCompras.Filter = ""
        ConexJardin.rsCompras.Filter = "[can]=" & datoc.Text
        LComprasc.Show
    Case 4:
        ConexJardin.rsCompras.Filter = ""
        LComprasc.Show
    Case Else: MsgBox "Escoja un tipo de filtro para visualizar el reporte!", vbInformation, "Reportes"
End Select
End Sub

Private Sub ca_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Caminadores' and [nivel]='A'"
LGrados.Caption = "Listado de niños para el grado de Caminadores - Nivel A"
LGrados.Show
End Sub

Private Sub cb_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Caminadores' and [nivel]='B'"
LGrados.Caption = "Listado de niños para el grado de Caminadores - Nivel B"
LGrados.Show
End Sub



Private Sub datoc_GotFocus()
datoc.SelStart = 0
datoc.SelLength = Len(datoc.Text)
End Sub

Private Sub datoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    bcom_Click
End If
End Sub

Private Sub datos_KeyPress(KeyAscii As Integer)
Select Case listamaterial.ListIndex
    Case 3:
        KeyAscii = Validar_letra(KeyAscii)
    Case 4:
        KeyAscii = Validar_numero(KeyAscii)
End Select
End Sub

Private Sub docuemp_GotFocus()
docuemp.SelStart = 0
docuemp.SelLength = Len(docuemp.Text)
End Sub

Private Sub docuemp_KeyPress(KeyAscii As Integer)
If EmpleadoVal = True Then
    KeyAscii = Validar_numero(KeyAscii)
ElseIf EmpleadoVal = False Then
    KeyAscii = Validar_letra(KeyAscii)
End If
If KeyAscii = 13 Then
    reportempleados_Click
End If
End Sub

Private Sub fecha4_LostFocus()
Dim ANIOs As Integer
ANIOs = Format(Date, "yyyy")

If fecha4.Year > ANIOs Or fecha4.Year < ANIOs Then
    MsgBox "Fecha no permitida!", vbExclamation, "Reportes"
    fecha4.SetFocus
    Exit Sub
End If
End Sub

Private Sub fechapagos_LostFocus()
Dim ANIOs As Integer
ANIOs = Format(Date, "yyyy")

If fechapagos.Year > ANIOs Or fechapagos.Year < ANIOs Then
    MsgBox "Fecha no permitida!", vbExclamation, "Reportes"
    fechapagos.SetFocus
    Exit Sub
End If
End Sub

Private Sub filtroins_Click()
Select Case filtroins.ListIndex
    Case 0:
        fecf.Visible = False
        numdocins.Visible = True
        numdocins.SetFocus
        listadoi.Visible = False
        reportinscripciones.Visible = True
        listain.Visible = True
    Case 1:
        fecf.Visible = True
        fecf.SetFocus
        numdocins.Visible = False
        listadoi.Visible = True
        reportinscripciones.Visible = False
        listain.Visible = False
    Case 2:
        fecf.Visible = False
        numdocins.Visible = False
        listadoi.Visible = True
        reportinscripciones.Visible = False
        listain.Visible = False
End Select
End Sub

Private Sub filtrole_Click()
Select Case filtrole.ListIndex
     Case 0:
        numdocle.Visible = True
        listadoe.Visible = True
        fecins.Visible = False
        listado.Visible = False
        vplis.Visible = False
     Case 1:
        numdocle.Visible = False
        listadoe.Visible = False
        fecins.Visible = True
        listado.Visible = True
        vplis.Visible = True
     Case 2:
        listadoe.Visible = True
        numdocle.Visible = False
        fecins.Visible = False
        listado.Visible = False
        vplis.Visible = False
End Select
End Sub

Private Sub Form_Activate()
FormularioActivo = True
menu.estado.Panels(4).Text = "Generador de Reportes"
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
menu.estado.Panels(4).Text = "Cargando..."
'llenar tipo documento en inscripciones
ConexionBD reportes, "select numdoc from inscripciones"
numdocins.Clear
If bd.Recordset.RecordCount > 0 Then
    bd.Recordset.MoveFirst
    For i = 1 To bd.Recordset.RecordCount
        numdocins.AddItem bd.Recordset!numdoc
        bd.Recordset.MoveNext
    Next i
End If
'llenar tipo documento en matriculas
ConexionBD reportes, "select numdoc from matricula"
listamat.Clear
If bd.Recordset.RecordCount > 0 Then
    bd.Recordset.MoveFirst
    For i = 1 To bd.Recordset.RecordCount
        listamat.AddItem bd.Recordset!numdoc
        bd.Recordset.MoveNext
    Next i
End If
'generar, llenar todo lo de grados y niveles de los niños

'verificar que el niño este matriculado
'para salacuna
Dim NombreC As String
On Error Resume Next
ConexionBD1 reportes, "select * from listadodeespera where meses >=3 and meses <=12"
If BD1.Recordset.RecordCount > 0 Then
    For i = 1 To BD1.Recordset.RecordCount
        'nivel A
        If i <= CantMax.SalaCuna Then
            'mira si el niño esta matriculado
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "SalaCuna", "A"
            End If
            BD1.Recordset.MoveNext
        End If
    Next i
End If
'para caminadores
On Error Resume Next
ConexionBD1 reportes, "select * from listadodeespera where meses >=13 and meses <=24"
If BD1.Recordset.RecordCount > 0 Then
    For i = 1 To BD1.Recordset.RecordCount
        'nivel A
        If i <= CantMax.Caminadores Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "Caminadores", "A"
            End If
            BD1.Recordset.MoveNext
        End If
        'nivel B
        If i > CantMax.Caminadores And i <= CantMax.Caminadores * 2 Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "Caminadores", "B"
            End If
            BD1.Recordset.MoveNext
        End If
    Next i
End If
'para parvulos
On Error Resume Next
ConexionBD1 reportes, "select * from listadodeespera where meses >=25 and meses <=30"
If BD1.Recordset.RecordCount > 0 Then
    For i = 1 To BD1.Recordset.RecordCount
        'nivel A
        If i <= CantMax.Parvulos Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "Párvulos", "A"
            End If
            BD1.Recordset.MoveNext
        End If
    Next i
End If
'para pre-kinder (2 1/2 años a 3 años)
On Error Resume Next
ConexionBD1 reportes, "select * from listadodeespera where meses >=31 and meses <=36"
If BD1.Recordset.RecordCount > 0 Then
    For i = 1 To BD1.Recordset.RecordCount
        'nivel A
        If i <= CantMax.Prekinder1 Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "PreKinder1", "A"
            End If
            BD1.Recordset.MoveNext
        End If
        'nivel B
        If i > CantMax.Prekinder1 And i <= CantMax.Prekinder1 * 2 Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "PreKinder1", "B"
            End If
            BD1.Recordset.MoveNext
        End If
        'nivel C
        If i > CantMax.Prekinder1 * 2 And i <= CantMax.Prekinder1 * 3 Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "PreKinder1", "C"
            End If
            BD1.Recordset.MoveNext
        End If
        'nivel D
        If i > CantMax.Prekinder1 * 3 And i <= CantMax.Prekinder1 * 4 Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "PreKinder1", "D"
            End If
            BD1.Recordset.MoveNext
        End If
    Next i
End If
'para pre-kinder (3 años a 4 años)
On Error Resume Next
ConexionBD1 reportes, "select * from listadodeespera where meses >=37 and meses <=48"
If BD1.Recordset.RecordCount > 0 Then
    For i = 1 To BD1.Recordset.RecordCount
        'nivel A
        If i <= CantMax.Prekinder2 Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "PreKinder2", "A"
            End If
            BD1.Recordset.MoveNext
        End If
        'nivel B
        If i > CantMax.Prekinder2 And i <= CantMax.Prekinder2 * 2 Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "PreKinder1", "B"
            End If
            BD1.Recordset.MoveNext
        End If
        'nivel C
        If i > CantMax.Prekinder2 * 2 And i <= CantMax.Prekinder2 * 3 Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "PreKinder1", "C"
            End If
            BD1.Recordset.MoveNext
        End If
        'nivel D
        If i > CantMax.Prekinder2 * 3 And i <= CantMax.Prekinder2 * 4 Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "PreKinder1", "D"
            End If
            BD1.Recordset.MoveNext
        End If
    Next i
End If
'para kinder
On Error Resume Next
ConexionBD1 reportes, "select * from listadodeespera where meses >=49 and meses <=60"
If BD1.Recordset.RecordCount > 0 Then
    For i = 1 To BD1.Recordset.RecordCount
        'nivel A
        If i <= CantMax.Kinder Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "Kinder", "A"
            End If
            BD1.Recordset.MoveNext
        End If
        'nivel B
        If i > CantMax.Kinder And i <= CantMax.Kinder * 2 Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "Kinder", "B"
            End If
            BD1.Recordset.MoveNext
        End If
        'nivel C
        If i > CantMax.Kinder * 2 And i <= CantMax.Kinder * 3 Then
            ConexionBD2 reportes, "select numdoc from matricula where numdoc='" & BD1.Recordset!numdoc & "'"
            If bd2.Recordset.RecordCount > 0 Then
                NombreC = BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
                GuardarGrado BD1.Recordset!numdoc, NombreC, "Kinder", "C"
            End If
            BD1.Recordset.MoveNext
        End If
    Next i
End If
End Sub
Private Sub GuardarGrado(Doc As String, nombre As String, Grado As String, nivel As String)
'agrega un nuevo registro a la tabla grado
ConexionBD2 reportes, "select * from grado"
bd2.Recordset.AddNew
    bd2.Recordset!numdoc = Doc
    bd2.Recordset!nombre = nombre
    bd2.Recordset!Grado = Grado
    bd2.Recordset!nivel = nivel
bd2.Recordset.Update
End Sub

Private Sub Form_Resize()
Me.WindowState = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
'borra los registros de grados
ConexionBD1 reportes, "select * from grado"
If BD1.Recordset.RecordCount > 0 Then
        BD1.Recordset.MoveFirst
        For i = 1 To BD1.Recordset.RecordCount
            On Error Resume Next
            BD1.Recordset.Delete
            BD1.Recordset.MoveNext
        Next i
        BD1.Refresh
End If
FormularioActivo = False
menu.estado.Panels(4).Text = "Menú Principal"

End Sub

Private Sub ka_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Kinder' and [nivel]='A'"
LGrados.Caption = "Listado de niños para el grado de Kinder - Nivel A"
LGrados.Show
End Sub

Private Sub kb_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Kinder' and [nivel]='B'"
LGrados.Caption = "Listado de niños para el grado de Kinder - Nivel B"
LGrados.Show
End Sub

Private Sub kc_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Kinder' and [nivel]='C'"
LGrados.Caption = "Listado de niños para el grado de Kinder - Nivel C"
LGrados.Show
End Sub

Private Sub la_Click()
reportmatriculas.Visible = False
ANIO.Visible = True
reportanio.Visible = True
End Sub

Private Sub lcompras_Click()
Select Case lcompras.ListIndex
    Case 0:
        datoc.Visible = True
        fecc.Visible = False
        datoc.Text = ""
        datoc.SetFocus
    Case 1:
        datoc.Visible = False
        fecc.Visible = True
        fecc.SetFocus
    Case 2:
        datoc.Visible = True
        fecc.Visible = False
        datoc.Text = ""
        datoc.SetFocus
    Case 3:
        datoc.Visible = True
        fecc.Visible = False
        datoc.Text = ""
        datoc.SetFocus
    Case 4:
        datoc.Visible = False
        fecc.Visible = False
End Select
End Sub

Private Sub listado_Click()

ConexionBD reportes, "select * from listadodeespera where fecins=#" & fecins.Month & "/" & fecins.Day & "/" & fecins.Year & "#"
If bd.Recordset.RecordCount = 0 Then
    MsgBox "No se puede imprimir el Listado de Espera por que la fecha no coincide con ningún registro!", vbCritical, "Reportes"
    Exit Sub
End If
If bd.Recordset.RecordCount > 0 Then
    'llena la tabla en word y la deja lista para imprimir
    On Error Resume Next
    Set ListEsp = New Word.Application
    ListEsp.Documents.Open App.Path + "\listado.doc"
    'llena el tipo del documento , numero de documento, sexo,edad,parentesco familiar,barrio, telefono,tipo salud,nivel,sec salud
    bd.Recordset.MoveFirst
    ListEsp.ActiveDocument.Tables(1).Cell(Row:=1, Column:=4).Tables(1).Cell(Row:=2, Column:=2).Range.InsertAfter fecins.Day
    ListEsp.ActiveDocument.Tables(1).Cell(Row:=1, Column:=4).Tables(1).Cell(Row:=2, Column:=3).Range.InsertAfter fecins.Month
    ListEsp.ActiveDocument.Tables(1).Cell(Row:=1, Column:=4).Tables(1).Cell(Row:=2, Column:=4).Range.InsertAfter fecins.Year
    
    For i = 3 To bd.Recordset.RecordCount + 5 Step 2
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=2).Range.InsertAfter bd.Recordset!tipdoc
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=3).Range.InsertAfter bd.Recordset!numdoc
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=6).Range.InsertAfter bd.Recordset!sex
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=7).Range.InsertAfter Left(bd.Recordset!fecnac, 2)
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=8).Range.InsertAfter Mid(bd.Recordset!fecnac, 4, 2)
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=9).Range.InsertAfter Right(bd.Recordset!fecnac, 4)
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=10).Range.InsertAfter bd.Recordset!eda
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=11).Range.InsertAfter bd.Recordset!parfam
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=12).Range.InsertAfter bd.Recordset!dir
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=13).Range.InsertAfter bd.Recordset!bar
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=14).Range.InsertAfter bd.Recordset!tel
        If bd.Recordset!sal = "EPS" Then
            ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=15).Range.InsertAfter "X"
        ElseIf bd.Recordset!sal = "ARS" Then
            ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=16).Range.InsertAfter "X"
        End If
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=17).Range.InsertAfter bd.Recordset!niv
        If bd.Recordset!secrsal = "SI" Then
            ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=18).Range.InsertAfter "X"
        End If
        
        bd.Recordset.MoveNext
    Next i
    'llena los apellidos, nombres
    bd.Recordset.MoveFirst
    For i = 4 To bd.Recordset.RecordCount + 6
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i, Column:=4).Range.InsertAfter bd.Recordset!priape
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i, Column:=5).Range.InsertAfter bd.Recordset!prinom
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=4).Range.InsertAfter bd.Recordset!segape
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=5).Range.InsertAfter bd.Recordset!segnom
        i = i + 1
        bd.Recordset.MoveNext
    Next i

    'muestra word con todo
    'ListEsp.Application.Visible = True
    ListEsp.ActiveDocument.PrintOut
    ListEsp.ActiveDocument.Close wdDoNotSaveChanges
    ListEsp.Application.Quit
Else
    MsgBox "No hay niños en listado con esta fecha de inscripción!", vbInformation, "Reportes"
End If
End Sub

Private Sub listains_Click()
Select Case listains.ListIndex
    Case 0:
        docuins.Visible = False
    Case 1:
        docuins.Visible = True
        docuins.SetFocus
End Select
End Sub

Private Sub listadoe_Click()
On Error Resume Next
Select Case filtrole.ListIndex
     Case 0:
        If numdocle.Text <> "" Then
            ConexJardin.rsListadoe.Filter = ""
            ConexJardin.rsListadoe.Filter = "[numdoc]=" & "'" & numdocle.Text & "'"
            Llm.Caption = "Listado de niños en espera de cupo"
            Llm.Show
        Else
            MsgBox "Por favor ingrese el número de documento del niño o niña", vbInformation, "Reportes"
            Exit Sub
        End If
     Case 2:
        ConexJardin.rsListadoe.Filter = ""
        Llm.Caption = "Listado de niños en espera de cupo"
        Llm.Show
End Select

End Sub

Private Sub listadoe_MouseOver()
menu.estado.Panels(4).Text = "Listado de todos los niños y niñas en espera de cupo"
End Sub

Private Sub listadoi_Click()
On Error Resume Next
Select Case filtroins.ListIndex
    Case 1:
        If fecf.Text <> "" Then
            ConexJardin.rsListadoi.Filter = ""
            ConexJardin.rsListadoi.Filter = "[numins]='" & fecf.Text & "'"
            Llmi.Caption = "Listado de niños en inscripciones"
            Llmi.Show
        Else
            MsgBox "Por favor ingrese el número de Inscripción", vbInformation, "Reportes"
        End If
    Case 2:
        ConexJardin.rsListadoi.Filter = ""
        Llmi.Caption = "Listado de niños en inscripciones"
        Llmi.Show
End Select


End Sub

Private Sub listadoi_MouseOver()
menu.estado.Panels(4).Text = "Listado de todos los niños y niñas en inscripciones"
End Sub

Private Sub listadom_Click()
On Error Resume Next
Me.PopupMenu mat, 2
End Sub

Private Sub listadom_MouseOver()
menu.estado.Panels(4).Text = "Listado de todos los niños y niñas en matriculas"
End Sub

Private Sub listain_Click()
ConexionBD reportes, "select * from inscripciones where numdoc='" & numdocins.Text & "'"
ConexionBD1 reportes, "select * from listadodeespera where numdoc='" & numdocins.Text & "'"
If bd.Recordset.RecordCount > 0 Then
    'llena la tabla en word y la deja lista para imprimir
    On Error Resume Next
    Set Incripcion = New Word.Application
    Incripcion.Documents.Open App.Path + "\inscripcion.doc"
    'coloca el numero de inscripcion
        Incripcion.ActiveDocument.Tables(1).Cell(Row:=1, Column:=2).Range.InsertAfter bd.Recordset!numins
    'llena los datos del niño
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=2, Column:=2).Range.InsertAfter BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=3, Column:=2).Range.InsertAfter BD1.Recordset!sex
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=3, Column:=4).Range.InsertAfter bd.Recordset!lugnac
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=4, Column:=2).Range.InsertAfter BD1.Recordset!fecnac
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=5, Column:=2).Range.InsertAfter bd.Recordset!numdoc
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=6, Column:=2).Range.InsertAfter BD1.Recordset!dir
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=7, Column:=2).Range.InsertAfter BD1.Recordset!bar
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=8, Column:=2).Range.InsertAfter bd.Recordset!prealgenf
    If BD1.Recordset!sal = "EPS" Then
        Incripcion.ActiveDocument.Tables(2).Cell(Row:=9, Column:=3).Range.InsertAfter "X"
    ElseIf BD1.Recordset!sal = "ARS" Then
        Incripcion.ActiveDocument.Tables(2).Cell(Row:=9, Column:=5).Range.InsertAfter "X"
    ElseIf BD1.Recordset!sal = "SISBEN" Then
        Incripcion.ActiveDocument.Tables(2).Cell(Row:=9, Column:=7).Range.InsertAfter "X"
    End If
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=10, Column:=2).Range.InsertAfter bd.Recordset!numher
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=10, Column:=4).Range.InsertAfter bd.Recordset!lugocufam
    'llena los datos de la familia
    'papa
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=2, Column:=2).Range.InsertAfter bd.Recordset!ninviv
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=3, Column:=2).Range.InsertAfter bd.Recordset!nompad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=4, Column:=2).Range.InsertAfter bd.Recordset!ocupad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=4, Column:=4).Range.InsertAfter bd.Recordset!ingmenpad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=4, Column:=6).Range.InsertAfter bd.Recordset!edapad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=5, Column:=2).Range.InsertAfter bd.Recordset!nomemppad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=5, Column:=4).Range.InsertAfter bd.Recordset!telemppad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=6, Column:=2).Range.InsertAfter bd.Recordset!nivacapad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=7, Column:=2).Range.InsertAfter bd.Recordset!otringpad
    'mama
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=9, Column:=2).Range.InsertAfter bd.Recordset!nommad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=10, Column:=2).Range.InsertAfter bd.Recordset!ocumad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=10, Column:=4).Range.InsertAfter bd.Recordset!ingmenmad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=10, Column:=6).Range.InsertAfter bd.Recordset!edamad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=11, Column:=2).Range.InsertAfter bd.Recordset!nomempmad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=11, Column:=4).Range.InsertAfter bd.Recordset!telempmad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=12, Column:=2).Range.InsertAfter bd.Recordset!nivacamad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=13, Column:=2).Range.InsertAfter bd.Recordset!otringmad
    
    'llenar datos de la condiciones de la vivienda
    Incripcion.ActiveDocument.Tables(4).Cell(Row:=2, Column:=2).Range.InsertAfter bd.Recordset!tenviv
    Incripcion.ActiveDocument.Tables(4).Cell(Row:=3, Column:=2).Range.InsertAfter bd.Recordset!tipviv
    Incripcion.ActiveDocument.Tables(4).Cell(Row:=5, Column:=2).Range.InsertAfter bd.Recordset!conviv
    Incripcion.ActiveDocument.Tables(4).Cell(Row:=6, Column:=2).Range.InsertAfter bd.Recordset!estviv
    Incripcion.ActiveDocument.Tables(4).Cell(Row:=8, Column:=1).Range.InsertAfter bd.Recordset!serpub
    'muestra word con todo
    'Incripcion.Application.Visible = True
    'Incripcion.ActiveDocument.PrintOut
    'Incripcion.ActiveDocument.Close wdDoNotSaveChanges
    'Incripcion.Application.Quit
Else
    MsgBox "El niño con documento: " & numdocins.Text & " no existe o no está inscrito!", vbCritical, "Reportes"
End If
End Sub

Private Sub listain_MouseOver()
menu.estado.Panels(4).Text = "vista Previa de la hoja de inscripción."
End Sub

Private Sub listama_Click()
ConexionBD reportes, "select * from matricula where numdoc='" & listamat.Text & "'"
ConexionBD1 reportes, "select * from listadodeespera where numdoc='" & listamat.Text & "'"
If BD1.Recordset.RecordCount = 0 Then
    MsgBox "No se puede imprimir este registro de matricula porque el documento: " & listamat.Text & vbCrLf & "no se encuentra en Listado de espera!", vbCritical, "Reportes"
    Exit Sub
End If
ConexionBD2 reportes, "select * from inscripciones where numdoc='" & listamat.Text & "'"
If bd2.Recordset.RecordCount = 0 Then
    MsgBox "No se puede imprimir este registro de matricula porque el documento: " & listamat.Text & vbCrLf & "no se encuentra en Inscripciones!", vbCritical, "Reportes"
    Exit Sub
End If
ConexionBD3 reportes, "select * from hermanos where numdoc='" & listamat.Text & "'"

If bd.Recordset.RecordCount > 0 Then
    'llena la tabla en word y la deja lista para imprimir
    On Error Resume Next
    Set matricula = New Word.Application
    matricula.Documents.Open App.Path + "\matricula.doc"
    'PRIMERA HOJA
    'llena los datos principales: numero formulario, fecha matricula, col, etc
    matricula.ActiveDocument.Tables(1).Cell(Row:=2, Column:=1).Range.InsertAfter bd.Recordset!col
    matricula.ActiveDocument.Tables(1).Cell(Row:=2, Column:=2).Range.InsertAfter bd.Recordset!uniope
    matricula.ActiveDocument.Tables(1).Cell(Row:=2, Column:=3).Range.InsertAfter bd.Recordset!modal
    matricula.ActiveDocument.Tables(1).Cell(Row:=2, Column:=4).Range.InsertAfter bd.Recordset!submod
    matricula.ActiveDocument.Tables(1).Cell(Row:=2, Column:=5).Range.InsertAfter bd.Recordset!fecmat
    matricula.ActiveDocument.Tables(1).Cell(Row:=2, Column:=6).Range.InsertAfter bd.Recordset!numfor
    'info del proyecto
    matricula.ActiveDocument.Tables(2).Cell(Row:=2, Column:=1).Range.InsertAfter bd.Recordset!persolser
    matricula.ActiveDocument.Tables(2).Cell(Row:=2, Column:=2).Range.InsertAfter bd.Recordset!rempor
    matricula.ActiveDocument.Tables(2).Cell(Row:=2, Column:=3).Range.InsertAfter bd.Recordset!prorem
    matricula.ActiveDocument.Tables(2).Cell(Row:=2, Column:=4).Range.InsertAfter bd.Recordset!entrem
    'llena los datos identificacion del beneficiario
    matricula.ActiveDocument.Tables(3).Cell(Row:=2, Column:=1).Range.InsertAfter bd.Recordset!numdoc
    matricula.ActiveDocument.Tables(3).Cell(Row:=2, Column:=2).Range.InsertAfter BD1.Recordset!tipdoc
    matricula.ActiveDocument.Tables(3).Cell(Row:=3, Column:=3).Range.InsertAfter BD1.Recordset!priape
    matricula.ActiveDocument.Tables(3).Cell(Row:=3, Column:=4).Range.InsertAfter BD1.Recordset!segape
    matricula.ActiveDocument.Tables(3).Cell(Row:=3, Column:=5).Range.InsertAfter BD1.Recordset!prinom & " " & BD1.Recordset!segnom
    'llenar datos basicos
    matricula.ActiveDocument.Tables(4).Cell(Row:=2, Column:=1).Range.InsertAfter BD1.Recordset!sex
    matricula.ActiveDocument.Tables(4).Cell(Row:=2, Column:=2).Range.InsertAfter BD1.Recordset!fecnac
    matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=3).Range.InsertAfter BD1.Recordset!edaano
    matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=4).Range.InsertAfter BD1.Recordset!edames
    matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=5).Range.InsertAfter bd.Recordset!depnac
    matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=6).Range.InsertAfter bd.Recordset!munnac
    matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=7).Range.InsertAfter bd.Recordset!Painac
    If bd.Recordset!mandis = "SI" Then
        matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=8).Range.InsertAfter "X"
    ElseIf bd.Recordset!mandis = "NO" Then
        matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=9).Range.InsertAfter "X"
    End If
    matricula.ActiveDocument.Tables(4).Cell(Row:=2, Column:=10).Range.InsertAfter bd.Recordset!tipdisest
    matricula.ActiveDocument.Tables(4).Cell(Row:=5, Column:=1).Range.InsertAfter BD1.Recordset!parfam
    matricula.ActiveDocument.Tables(4).Cell(Row:=6, Column:=2).Range.InsertAfter bd.Recordset!niveduben
    matricula.ActiveDocument.Tables(4).Cell(Row:=6, Column:=3).Range.InsertAfter bd.Recordset!asiactcenedu
    matricula.ActiveDocument.Tables(4).Cell(Row:=8, Column:=1).Range.InsertAfter bd.Recordset!proaso
    'llena los datos de seguridad social en salud
    matricula.ActiveDocument.Tables(5).Cell(Row:=2, Column:=1).Range.InsertAfter bd.Recordset!afisegsocben
    matricula.ActiveDocument.Tables(5).Cell(Row:=2, Column:=2).Range.InsertAfter bd.Recordset!regsegsocben
    matricula.ActiveDocument.Tables(5).Cell(Row:=2, Column:=3).Range.InsertAfter bd.Recordset!calben
    matricula.ActiveDocument.Tables(5).Cell(Row:=2, Column:=4).Range.InsertAfter bd.Recordset!vinsecsalben
    matricula.ActiveDocument.Tables(5).Cell(Row:=3, Column:=5).Range.InsertAfter bd.Recordset!numficsis
    matricula.ActiveDocument.Tables(5).Cell(Row:=3, Column:=6).Range.InsertAfter bd.Recordset!punsis
    'llena datos de la ubicacion del nucleo familiar
    matricula.ActiveDocument.Tables(6).Cell(Row:=2, Column:=1).Range.InsertAfter BD1.Recordset!dir
    matricula.ActiveDocument.Tables(6).Cell(Row:=2, Column:=2).Range.InsertAfter bd.Recordset!loc
    If IsNull(bd.Recordset!est) = False Then
        matricula.ActiveDocument.Tables(6).Cell(Row:=3, Column:=3).Range.InsertAfter bd.Recordset!est
    End If
    matricula.ActiveDocument.Tables(6).Cell(Row:=3, Column:=4).Range.InsertAfter bd2.Recordset!tipviv
    matricula.ActiveDocument.Tables(6).Cell(Row:=3, Column:=5).Range.InsertAfter bd2.Recordset!conviv
    matricula.ActiveDocument.Tables(6).Cell(Row:=3, Column:=6).Range.InsertAfter bd2.Recordset!tenviv
    matricula.ActiveDocument.Tables(6).Cell(Row:=3, Column:=7).Range.InsertAfter bd.Recordset!forpagviv
    matricula.ActiveDocument.Tables(6).Cell(Row:=6, Column:=2).Range.InsertAfter bd.Recordset!dptoprofam
    matricula.ActiveDocument.Tables(6).Cell(Row:=6, Column:=3).Range.InsertAfter bd.Recordset!munprofam
    matricula.ActiveDocument.Tables(6).Cell(Row:=6, Column:=4).Range.InsertAfter bd.Recordset!paiprofam
    matricula.ActiveDocument.Tables(6).Cell(Row:=6, Column:=6).Range.InsertAfter Mid(bd.Recordset!fecllebogfam, 4, 2)
    matricula.ActiveDocument.Tables(6).Cell(Row:=6, Column:=7).Range.InsertAfter Mid(bd.Recordset!fecllebogfam, 7, 4)
    'llena la informacion especifica para el proyecto
    matricula.ActiveDocument.Tables(7).Cell(Row:=2, Column:=1).Range.InsertAfter bd.Recordset!ninvivpapmam
    matricula.ActiveDocument.Tables(7).Cell(Row:=2, Column:=1).Range.InsertAfter bd.Recordset!ninvivperpadmadotr
    matricula.ActiveDocument.Tables(7).Cell(Row:=2, Column:=3).Range.InsertAfter bd.Recordset!vivperpadmad
    matricula.ActiveDocument.Tables(7).Cell(Row:=2, Column:=4).Range.InsertAfter bd.Recordset!edaninvivpapmad
    matricula.ActiveDocument.Tables(7).Cell(Row:=2, Column:=5).Range.InsertAfter bd.Recordset!cuinindurdia
    matricula.ActiveDocument.Tables(7).Cell(Row:=2, Column:=6).Range.InsertAfter bd.Recordset!graasp
    'SEGUNDA HOJA
    If bd3.Recordset.RecordCount > 0 Then
        bd3.Recordset.MoveFirst
        For i = 5 To bd3.Recordset.RecordCount + 4
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=2).Range.InsertAfter bd3.Recordset!tipdocher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=3).Range.InsertAfter bd3.Recordset!numdocher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=4).Range.InsertAfter bd3.Recordset!apeher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=5).Range.InsertAfter bd3.Recordset!nomher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=6).Range.InsertAfter bd3.Recordset!sexher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=7).Range.InsertAfter Mid(bd3.Recordset!fecnacher, 1, 2)
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=8).Range.InsertAfter Mid(bd3.Recordset!fecnacher, 4, 2)
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=9).Range.InsertAfter Mid(bd3.Recordset!fecnacher, 7, 4)
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=10).Range.InsertAfter bd3.Recordset!edaproher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=11).Range.InsertAfter bd3.Recordset!estcivher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=12).Range.InsertAfter bd3.Recordset!tipdisher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=13).Range.InsertAfter bd3.Recordset!parher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=14).Range.InsertAfter bd3.Recordset!esther
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=15).Range.InsertAfter bd3.Recordset!asiacteduher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=16).Range.InsertAfter bd3.Recordset!actocuher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=17).Range.InsertAfter bd3.Recordset!posocuher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=18).Range.InsertAfter bd3.Recordset!ingmesher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=19).Range.InsertAfter bd3.Recordset!forperingher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=20).Range.InsertAfter bd3.Recordset!afisalher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=21).Range.InsertAfter bd3.Recordset!regsegsalher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=22).Range.InsertAfter bd3.Recordset!calbenher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=23).Range.InsertAfter bd3.Recordset!vinsecsalher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=24).Range.InsertAfter bd3.Recordset!agrviointher
            bd3.Recordset.MoveNext
        Next i
    End If
    'llena los datos finales
    matricula.ActiveDocument.Tables(10).Cell(Row:=3, Column:=2).Range.InsertAfter bd.Recordset!nomdilform
    matricula.ActiveDocument.Tables(10).Cell(Row:=3, Column:=3).Range.InsertAfter bd.Recordset!fecdighojsir
    matricula.ActiveDocument.Tables(10).Cell(Row:=3, Column:=4).Range.InsertAfter bd.Recordset!nomfundighojsir
    matricula.ActiveDocument.Tables(10).Cell(Row:=3, Column:=5).Range.InsertAfter bd.Recordset!obs
    
    matricula.Application.Visible = True
    'matricula.ActiveDocument.PrintOut
    'matricula.ActiveDocument.Close wdDoNotSaveChanges
    'matricula.Application.Quit
Else
    MsgBox "El niño con documento: " & listamat.Text & " no existe o no está matriculado!", vbInformation, "Reportes"
End If
End Sub

Private Sub listama_MouseOver()
menu.estado.Panels(4).Text = "Vista previa de la hoja de matricula."
End Sub



Private Sub listamat_GotFocus()
reportmatriculas.Visible = True
ANIO.Visible = False
reportanio.Visible = False

End Sub

Private Sub listamaterial_Click()
Select Case listamaterial.ListIndex
    Case 0:
        datos.Visible = False
        fechades.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        fechamat.Visible = False
    Case 1:
        datos.Visible = False
        fechades.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        fechamat.Visible = False
    Case 2:
        datos.Visible = False
        fechades.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        fechamat.Visible = True
        fechades.SetFocus
    Case 3:
        datos.Visible = True
        fechades.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        fechamat.Visible = False
        datos.Text = ""
        datos.SetFocus
    Case 4:
        datos.Visible = True
        fechades.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        fechamat.Visible = False
        datos.Text = ""
        datos.SetFocus
    Case 5:
        datos.Visible = True
        fechades.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        fechamat.Visible = False
        datos.Text = ""
        datos.SetFocus
End Select
End Sub

Private Sub listemp_Click()
Select Case listemp.ListIndex
    Case 0:
        docuemp.Visible = False
        fecvin.Visible = False
    Case 1:
        docuemp.Visible = True
        fecvin.Visible = False
        EmpleadoVal = True
        docuemp.Text = ""
        docuemp.SetFocus
    Case 2:
        docuemp.Visible = True
        fecvin.Visible = False
        EmpleadoVal = False
        docuemp.Text = ""
        docuemp.SetFocus
    Case 3:
        docuemp.Visible = False
        fecvin.Visible = True
        fecvin.SetFocus
    Case 4:
        docuemp.Visible = True
        fecvin.Visible = False
        EmpleadoVal = False
        docuemp.Text = ""
        docuemp.SetFocus
    Case 5:
        docuemp.Visible = False
        fecvin.Visible = True
        fecvin.SetFocus
End Select
End Sub

Private Sub listpag_Click()
Select Case listpag.ListIndex
    Case 0, 1, 2, 3:
        fecha4.Visible = False
        Label5.Visible = False
        Label4.Visible = False
        fechapagos.Visible = False
        reportpag.Visible = True
        reportp.Visible = False
        pnum.Visible = False
    Case 4:
        fecha4.Visible = True
        Label5.Visible = True
        Label4.Visible = True
        pnum.Visible = False
        reportpag.Visible = True
        fechapagos.Visible = True
        reportp.Visible = False
        fecha4.SetFocus
    Case 5:
        pnum.Visible = True
        fecha4.Visible = False
        Label5.Visible = False
        Label4.Visible = False
        fechapagos.Visible = False
        reportp.Visible = True
        pnum.SetFocus
        reportpag.Visible = False
End Select
End Sub

Private Sub lt_Click()
ConexJardin.rsListadom.Filter = ""
Llmim.Caption = "Listado de niños en Matriculas"
Llmim.Show
End Sub

Private Sub matriculas_Click()
If pnum.Text <> "" Then
    ConexJardin.rsPagosJ.Filter = ""
    ConexJardin.rsPagosJ.Filter = "[numdoc]='" & pnum.Text & "' and [tipopago]='Matricula'"
    Matriculap.Show
Else
    MsgBox "No ha ingresado el número de documento!", vbInformation, "Reportes"
End If
End Sub

Private Sub mt_Click()
If pnum.Text <> "" Then
    ConexJardin.rsPagosJ.Filter = ""
    ConexJardin.rsPagosJ.Filter = "[numdoc]='" & pnum.Text & "'"
    LPagos.Show
Else
    MsgBox "No ha ingresado el número de documento!", vbInformation, "Reportes"
End If
End Sub

Private Sub otrosc_Click()
If pnum.Text <> "" Then
    ConexJardin.rsPagosJ.Filter = ""
    ConexJardin.rsPagosJ.Filter = "[numdoc]='" & pnum.Text & "' and [tipopago]='Otros'"
    Otrosp.Show
Else
    MsgBox "No ha ingresado el número de documento!", vbInformation, "Reportes"
End If
End Sub

Private Sub p1a_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Prekinder1' and [nivel]='A'"
LGrados.Caption = "Listado de niños para el grado de Pre-Kinder - Nivel A"
LGrados.Show
End Sub

Private Sub p1b_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Prekinder1' and [nivel]='B'"
LGrados.Caption = "Listado de niños para el grado de Pre-Kinder - Nivel B"
LGrados.Show
End Sub

Private Sub p1c_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Prekinder1' and [nivel]='C'"
LGrados.Caption = "Listado de niños para el grado de Pre-Kinder - Nivel C"
LGrados.Show
End Sub

Private Sub p1d_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Prekinder1' and [nivel]='D'"
LGrados.Caption = "Listado de niños para el grado de Pre-Kinder - Nivel D"
LGrados.Show
End Sub

Private Sub p2a_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Prekinder2' and [nivel]='A'"
LGrados.Caption = "Listado de niños para el grado de Pre-Kinder - Nivel A"
LGrados.Show
End Sub

Private Sub p2b_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Prekinder2' and [nivel]='B'"
LGrados.Caption = "Listado de niños para el grado de Pre-Kinder - Nivel B"
LGrados.Show
End Sub

Private Sub p2c_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Prekinder2' and [nivel]='C'"
LGrados.Caption = "Listado de niños para el grado de Pre-Kinder - Nivel C"
LGrados.Show
End Sub

Private Sub p2d_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Prekinder2' and [nivel]='D'"
LGrados.Caption = "Listado de niños para el grado de Pre-Kinder - Nivel D"
LGrados.Show
End Sub

Private Sub pa_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='Párvulos'"
LGrados.Caption = "Listado de niños para el grado de Párvulos - Nivel A"
LGrados.Show
End Sub

Private Sub pensiones_Click()
If pnum.Text <> "" Then
    ConexJardin.rsPagosJ.Filter = ""
    ConexJardin.rsPagosJ.Filter = "[numdoc]='" & pnum.Text & "' and [tipopago]='Pensión'"
    Pension.Show
Else
    MsgBox "No ha ingresado el número de documento!", vbInformation, "Reportes"
End If
End Sub

Private Sub reportanio_Click()
Dim annior
anior = Val(ANIO.Text)
ConexJardin.rsListadom.Filter = ""
ConexJardin.rsListadom.Filter = "[fecmat]>=" & "#01/01/" & anior & "#" & "and [fecmat]<=" & "#12/31/" & anior & "#"
Llmim.Show
End Sub

Private Sub reportempleados_Click()
On Error Resume Next
Select Case listemp.ListIndex
    Case 0:
        ConexJardin.rsEmpleados.Filter = ""
        LEmpleados.Show
    Case 1:
        If docuemp.Text = "" Then
            MsgBox "No hay ingresado el documento para realizar el filtro del reporte!", vbInformation, "Reportes"
            docuemp.SetFocus
            Exit Sub
        End If
        ConexJardin.rsEmpleados.Filter = ""
        ConexJardin.rsEmpleados.Filter = "[numdoc]=" & "'" & docuemp.Text & "'"
        LEmpleados.Show
    Case 2:
        ConexJardin.rsEmpleados.Filter = ""
        ConexJardin.rsEmpleados.Filter = "[car]=" & "'" & docuemp.Text & "'"
        LEmpleados.Show
    Case 3:
        ConexJardin.rsEmpleados.Filter = ""
        ConexJardin.rsEmpleados.Filter = "[fecvin]=" & fecvin.Value
        LEmpleados.Show
    Case 4:
        ConexJardin.rsEmpleados.Filter = ""
        ConexJardin.rsEmpleados.Filter = "[pro]=" & "'" & docuemp.Text & "'"
        LEmpleados.Show
    Case 5:
        ConexJardin.rsEmpleados.Filter = ""
        ConexJardin.rsEmpleados.Filter = "[fecnac]=" & fecvin.Value
        LEmpleados.Show
    Case Else: MsgBox "Escoja un tipo de filtro para visualizar el reporte!", vbInformation, "Reportes"
End Select
End Sub



Private Sub reportinscripciones_Click()
ConexionBD reportes, "select * from inscripciones where numdoc='" & numdocins.Text & "'"
ConexionBD1 reportes, "select * from listadodeespera where numdoc='" & numdocins.Text & "'"
If bd.Recordset.RecordCount > 0 Then
    'llena la tabla en word y la deja lista para imprimir
    On Error Resume Next
    Set Incripcion = New Word.Application
    Incripcion.Documents.Open App.Path + "\inscripcion.doc"
    'coloca el numero de inscripcion
        Incripcion.ActiveDocument.Tables(1).Cell(Row:=1, Column:=2).Range.InsertAfter bd.Recordset!numins
    'llena los datos del niño
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=2, Column:=2).Range.InsertAfter BD1.Recordset!prinom & " " & BD1.Recordset!segnom & " " & BD1.Recordset!priape & " " & BD1.Recordset!segape
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=3, Column:=2).Range.InsertAfter BD1.Recordset!sex
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=3, Column:=4).Range.InsertAfter bd.Recordset!lugnac
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=4, Column:=2).Range.InsertAfter BD1.Recordset!fecnac
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=5, Column:=2).Range.InsertAfter bd.Recordset!numdoc
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=6, Column:=2).Range.InsertAfter BD1.Recordset!dir
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=7, Column:=2).Range.InsertAfter BD1.Recordset!bar
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=8, Column:=2).Range.InsertAfter bd.Recordset!prealgenf
    If BD1.Recordset!sal = "EPS" Then
        Incripcion.ActiveDocument.Tables(2).Cell(Row:=9, Column:=3).Range.InsertAfter "X"
    ElseIf BD1.Recordset!sal = "ARS" Then
        Incripcion.ActiveDocument.Tables(2).Cell(Row:=9, Column:=5).Range.InsertAfter "X"
    ElseIf BD1.Recordset!sal = "SISBEN" Then
        Incripcion.ActiveDocument.Tables(2).Cell(Row:=9, Column:=7).Range.InsertAfter "X"
    End If
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=10, Column:=2).Range.InsertAfter bd.Recordset!numher
    Incripcion.ActiveDocument.Tables(2).Cell(Row:=10, Column:=4).Range.InsertAfter bd.Recordset!lugocufam
    'llena los datos de la familia
    'papa
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=2, Column:=2).Range.InsertAfter bd.Recordset!ninviv
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=3, Column:=2).Range.InsertAfter bd.Recordset!nompad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=4, Column:=2).Range.InsertAfter bd.Recordset!ocupad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=4, Column:=4).Range.InsertAfter bd.Recordset!ingmenpad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=4, Column:=6).Range.InsertAfter bd.Recordset!edapad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=5, Column:=2).Range.InsertAfter bd.Recordset!nomemppad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=5, Column:=4).Range.InsertAfter bd.Recordset!telemppad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=6, Column:=2).Range.InsertAfter bd.Recordset!nivacapad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=7, Column:=2).Range.InsertAfter bd.Recordset!otringpad
    'mama
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=9, Column:=2).Range.InsertAfter bd.Recordset!nommad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=10, Column:=2).Range.InsertAfter bd.Recordset!ocumad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=10, Column:=4).Range.InsertAfter bd.Recordset!ingmenmad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=10, Column:=6).Range.InsertAfter bd.Recordset!edamad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=11, Column:=2).Range.InsertAfter bd.Recordset!nomempmad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=11, Column:=4).Range.InsertAfter bd.Recordset!telempmad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=12, Column:=2).Range.InsertAfter bd.Recordset!nivacamad
    Incripcion.ActiveDocument.Tables(3).Cell(Row:=13, Column:=2).Range.InsertAfter bd.Recordset!otringmad
    
    'llenar datos de la condiciones de la vivienda
    Incripcion.ActiveDocument.Tables(4).Cell(Row:=2, Column:=2).Range.InsertAfter bd.Recordset!tenviv
    Incripcion.ActiveDocument.Tables(4).Cell(Row:=3, Column:=2).Range.InsertAfter bd.Recordset!tipviv
    Incripcion.ActiveDocument.Tables(4).Cell(Row:=5, Column:=2).Range.InsertAfter bd.Recordset!conviv
    Incripcion.ActiveDocument.Tables(4).Cell(Row:=6, Column:=2).Range.InsertAfter bd.Recordset!estviv
    Incripcion.ActiveDocument.Tables(4).Cell(Row:=8, Column:=1).Range.InsertAfter bd.Recordset!serpub
    'muestra word con todo
    'Incripcion.Application.Visible = True
    Incripcion.ActiveDocument.PrintOut
    Incripcion.ActiveDocument.Close wdDoNotSaveChanges
    Incripcion.Application.Quit
Else
    MsgBox "El niño con documento: " & numdocins.Text & " no existe o no está inscrito!", vbCritical, "Reportes"
End If
End Sub

Private Sub reportmaterial_Click()
On Error Resume Next
Select Case listamaterial.ListIndex
    Case 0: 'material existente
        ConexJardin.rsMateriales.Filter = ""
        LMaterial.Show
    Case 1:
        ConexJardin.rsMeMd.Filter = ""
        LMatDespachado.Show
    Case 2:
        ConexJardin.rsMeMd.Filter = ""
        ConexJardin.rsMeMd.Filter = "[fecdes]>=" & fechades.Value & " and [fecdes]<=" & fechamat.Value
        LMatDespachado.Show
    Case 3:
        If datos.Text = "" Then
            MsgBox "No hay ingresado el nombre de la sede para realizar el filtro del reporte!", vbInformation, "Reportes"
            datos.SetFocus
            Exit Sub
        End If
        ConexJardin.rsMeMd.Filter = ""
        ConexJardin.rsMeMd.Filter = "[sede]=" & "'" & datos.Text & "'"
        LMatDespachado.Show
    Case 4:
        If datos.Text = "" Then
            MsgBox "No hay ingresado la referencia del material para realizar el filtro del reporte!", vbInformation, "Reportes"
            datos.SetFocus
            Exit Sub
        End If
        ConexJardin.rsMeMd.Filter = ""
        ConexJardin.rsMeMd.Filter = "[ref]=" & "'" & Val(datos.Text) & "'"
        LMatDespachado.Show
    Case 5:
        If datos.Text = "" Then
            MsgBox "No hay ingresado el número de codumento del empleado para realizar el reporte!", vbInformation, "Reportes"
            datos.SetFocus
            Exit Sub
        End If
        ConexJardin.rsMeMd.Filter = ""
        ConexJardin.rsMeMd.Filter = "[numdoc]=" & "'" & Val(datos.Text) & "'"
        LMatDespachado.Show
    Case Else: MsgBox "Escoja un tipo de filtro para visualizar el reporte!", vbInformation, "Reportes"
End Select
End Sub

Private Sub reportmatriculas_Click()
ConexionBD reportes, "select * from matricula where numdoc='" & listamat.Text & "'"
ConexionBD1 reportes, "select * from listadodeespera where numdoc='" & listamat.Text & "'"
If BD1.Recordset.RecordCount = 0 Then
    MsgBox "No se puede imprimir este registro de matricula porque el documento: " & listamat.Text & vbCrLf & "no se encuentra en Listado de espera!", vbCritical, "Reportes"
    Exit Sub
End If
ConexionBD2 reportes, "select * from inscripciones where numdoc='" & listamat.Text & "'"
If bd2.Recordset.RecordCount = 0 Then
    MsgBox "No se puede imprimir este registro de matricula porque el documento: " & listamat.Text & vbCrLf & "no se encuentra en Inscripciones!", vbCritical, "Reportes"
    Exit Sub
End If
ConexionBD3 reportes, "select * from hermanos where numdoc='" & listamat.Text & "'"

If bd.Recordset.RecordCount > 0 Then
    'llena la tabla en word y la deja lista para imprimir
    On Error Resume Next
    Set matricula = New Word.Application
    matricula.Documents.Open App.Path + "\matricula.doc"
    'PRIMERA HOJA
    'llena los datos principales: numero formulario, fecha matricula, col, etc
    matricula.ActiveDocument.Tables(1).Cell(Row:=2, Column:=1).Range.InsertAfter bd.Recordset!col
    matricula.ActiveDocument.Tables(1).Cell(Row:=2, Column:=2).Range.InsertAfter bd.Recordset!uniope
    matricula.ActiveDocument.Tables(1).Cell(Row:=2, Column:=3).Range.InsertAfter bd.Recordset!modal
    matricula.ActiveDocument.Tables(1).Cell(Row:=2, Column:=4).Range.InsertAfter bd.Recordset!submod
    matricula.ActiveDocument.Tables(1).Cell(Row:=2, Column:=5).Range.InsertAfter bd.Recordset!fecmat
    matricula.ActiveDocument.Tables(1).Cell(Row:=2, Column:=6).Range.InsertAfter bd.Recordset!numfor
    'info del proyecto
    matricula.ActiveDocument.Tables(2).Cell(Row:=2, Column:=1).Range.InsertAfter bd.Recordset!persolser
    matricula.ActiveDocument.Tables(2).Cell(Row:=2, Column:=2).Range.InsertAfter bd.Recordset!rempor
    matricula.ActiveDocument.Tables(2).Cell(Row:=2, Column:=3).Range.InsertAfter bd.Recordset!prorem
    matricula.ActiveDocument.Tables(2).Cell(Row:=2, Column:=4).Range.InsertAfter bd.Recordset!entrem
    'llena los datos identificacion del beneficiario
    matricula.ActiveDocument.Tables(3).Cell(Row:=2, Column:=1).Range.InsertAfter bd.Recordset!numdoc
    matricula.ActiveDocument.Tables(3).Cell(Row:=2, Column:=2).Range.InsertAfter BD1.Recordset!tipdoc
    matricula.ActiveDocument.Tables(3).Cell(Row:=3, Column:=3).Range.InsertAfter BD1.Recordset!priape
    matricula.ActiveDocument.Tables(3).Cell(Row:=3, Column:=4).Range.InsertAfter BD1.Recordset!segape
    matricula.ActiveDocument.Tables(3).Cell(Row:=3, Column:=5).Range.InsertAfter BD1.Recordset!prinom & " " & BD1.Recordset!segnom
    'llenar datos basicos
    matricula.ActiveDocument.Tables(4).Cell(Row:=2, Column:=1).Range.InsertAfter BD1.Recordset!sex
    matricula.ActiveDocument.Tables(4).Cell(Row:=2, Column:=2).Range.InsertAfter BD1.Recordset!fecnac
    matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=3).Range.InsertAfter BD1.Recordset!edaano
    matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=4).Range.InsertAfter BD1.Recordset!edames
    matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=5).Range.InsertAfter bd.Recordset!depnac
    matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=6).Range.InsertAfter bd.Recordset!munnac
    matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=7).Range.InsertAfter bd.Recordset!Painac
    If bd.Recordset!mandis = "SI" Then
        matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=8).Range.InsertAfter "X"
    ElseIf bd.Recordset!mandis = "NO" Then
        matricula.ActiveDocument.Tables(4).Cell(Row:=3, Column:=9).Range.InsertAfter "X"
    End If
    matricula.ActiveDocument.Tables(4).Cell(Row:=2, Column:=10).Range.InsertAfter bd.Recordset!tipdisest
    matricula.ActiveDocument.Tables(4).Cell(Row:=5, Column:=1).Range.InsertAfter BD1.Recordset!parfam
    matricula.ActiveDocument.Tables(4).Cell(Row:=6, Column:=2).Range.InsertAfter bd.Recordset!niveduben
    matricula.ActiveDocument.Tables(4).Cell(Row:=6, Column:=3).Range.InsertAfter bd.Recordset!asiactcenedu
    matricula.ActiveDocument.Tables(4).Cell(Row:=8, Column:=1).Range.InsertAfter bd.Recordset!proaso
    'llena los datos de seguridad social en salud
    matricula.ActiveDocument.Tables(5).Cell(Row:=2, Column:=1).Range.InsertAfter bd.Recordset!afisegsocben
    matricula.ActiveDocument.Tables(5).Cell(Row:=2, Column:=2).Range.InsertAfter bd.Recordset!regsegsocben
    matricula.ActiveDocument.Tables(5).Cell(Row:=2, Column:=3).Range.InsertAfter bd.Recordset!calben
    matricula.ActiveDocument.Tables(5).Cell(Row:=2, Column:=4).Range.InsertAfter bd.Recordset!vinsecsalben
    matricula.ActiveDocument.Tables(5).Cell(Row:=3, Column:=5).Range.InsertAfter bd.Recordset!numficsis
    matricula.ActiveDocument.Tables(5).Cell(Row:=3, Column:=6).Range.InsertAfter bd.Recordset!punsis
    'llena datos de la ubicacion del nucleo familiar
    matricula.ActiveDocument.Tables(6).Cell(Row:=2, Column:=1).Range.InsertAfter BD1.Recordset!dir
    matricula.ActiveDocument.Tables(6).Cell(Row:=2, Column:=2).Range.InsertAfter bd.Recordset!loc
    If IsNull(bd.Recordset!est) = False Then
        matricula.ActiveDocument.Tables(6).Cell(Row:=3, Column:=3).Range.InsertAfter bd.Recordset!est
    End If
    matricula.ActiveDocument.Tables(6).Cell(Row:=3, Column:=4).Range.InsertAfter bd2.Recordset!tipviv
    matricula.ActiveDocument.Tables(6).Cell(Row:=3, Column:=5).Range.InsertAfter bd2.Recordset!conviv
    matricula.ActiveDocument.Tables(6).Cell(Row:=3, Column:=6).Range.InsertAfter bd2.Recordset!tenviv
    matricula.ActiveDocument.Tables(6).Cell(Row:=3, Column:=7).Range.InsertAfter bd.Recordset!forpagviv
    matricula.ActiveDocument.Tables(6).Cell(Row:=6, Column:=2).Range.InsertAfter bd.Recordset!dptoprofam
    matricula.ActiveDocument.Tables(6).Cell(Row:=6, Column:=3).Range.InsertAfter bd.Recordset!munprofam
    matricula.ActiveDocument.Tables(6).Cell(Row:=6, Column:=4).Range.InsertAfter bd.Recordset!paiprofam
    matricula.ActiveDocument.Tables(6).Cell(Row:=6, Column:=6).Range.InsertAfter Mid(bd.Recordset!fecllebogfam, 4, 2)
    matricula.ActiveDocument.Tables(6).Cell(Row:=6, Column:=7).Range.InsertAfter Mid(bd.Recordset!fecllebogfam, 7, 4)
    'llena la informacion especifica para el proyecto
    matricula.ActiveDocument.Tables(7).Cell(Row:=2, Column:=1).Range.InsertAfter bd.Recordset!ninvivpapmam
    matricula.ActiveDocument.Tables(7).Cell(Row:=2, Column:=1).Range.InsertAfter bd.Recordset!ninvivperpadmadotr
    matricula.ActiveDocument.Tables(7).Cell(Row:=2, Column:=3).Range.InsertAfter bd.Recordset!vivperpadmad
    matricula.ActiveDocument.Tables(7).Cell(Row:=2, Column:=4).Range.InsertAfter bd.Recordset!edaninvivpapmad
    matricula.ActiveDocument.Tables(7).Cell(Row:=2, Column:=5).Range.InsertAfter bd.Recordset!cuinindurdia
    matricula.ActiveDocument.Tables(7).Cell(Row:=2, Column:=6).Range.InsertAfter bd.Recordset!graasp
    'SEGUNDA HOJA
    If bd3.Recordset.RecordCount > 0 Then
        bd3.Recordset.MoveFirst
        For i = 5 To bd3.Recordset.RecordCount + 4
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=2).Range.InsertAfter bd3.Recordset!tipdocher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=3).Range.InsertAfter bd3.Recordset!numdocher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=4).Range.InsertAfter bd3.Recordset!apeher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=5).Range.InsertAfter bd3.Recordset!nomher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=6).Range.InsertAfter bd3.Recordset!sexher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=7).Range.InsertAfter Mid(bd3.Recordset!fecnacher, 1, 2)
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=8).Range.InsertAfter Mid(bd3.Recordset!fecnacher, 4, 2)
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=9).Range.InsertAfter Mid(bd3.Recordset!fecnacher, 7, 4)
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=10).Range.InsertAfter bd3.Recordset!edaproher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=11).Range.InsertAfter bd3.Recordset!estcivher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=12).Range.InsertAfter bd3.Recordset!tipdisher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=13).Range.InsertAfter bd3.Recordset!parher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=14).Range.InsertAfter bd3.Recordset!esther
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=15).Range.InsertAfter bd3.Recordset!asiacteduher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=16).Range.InsertAfter bd3.Recordset!actocuher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=17).Range.InsertAfter bd3.Recordset!posocuher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=18).Range.InsertAfter bd3.Recordset!ingmesher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=19).Range.InsertAfter bd3.Recordset!forperingher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=20).Range.InsertAfter bd3.Recordset!afisalher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=21).Range.InsertAfter bd3.Recordset!regsegsalher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=22).Range.InsertAfter bd3.Recordset!calbenher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=23).Range.InsertAfter bd3.Recordset!vinsecsalher
            matricula.ActiveDocument.Tables(9).Cell(Row:=i, Column:=24).Range.InsertAfter bd3.Recordset!agrviointher
            bd3.Recordset.MoveNext
        Next i
    End If
    'llena los datos finales
    matricula.ActiveDocument.Tables(10).Cell(Row:=3, Column:=2).Range.InsertAfter bd.Recordset!nomdilform
    matricula.ActiveDocument.Tables(10).Cell(Row:=3, Column:=3).Range.InsertAfter bd.Recordset!fecdighojsir
    matricula.ActiveDocument.Tables(10).Cell(Row:=3, Column:=4).Range.InsertAfter bd.Recordset!nomfundighojsir
    matricula.ActiveDocument.Tables(10).Cell(Row:=3, Column:=5).Range.InsertAfter bd.Recordset!obs
    
    'matricula.Application.Visible = True
    matricula.ActiveDocument.PrintOut
    matricula.ActiveDocument.Close wdDoNotSaveChanges
    matricula.Application.Quit
Else
    MsgBox "El niño con documento: " & listamat.Text & " no existe o no está matriculado!", vbInformation, "Reportes"
End If
End Sub

Private Sub reportp_Click()
Me.PopupMenu numpagos, 2
End Sub

Private Sub reportpag_Click()
On Error Resume Next
Select Case listpag.ListIndex
    Case 0:
        ConexJardin.rsPagosJ.Filter = ""
        LPagos.Show
    Case 1:
        ConexJardin.rsPagosJ.Filter = ""
        ConexJardin.rsPagosJ.Filter = "[tipopago]= 'Pensión'"
        LPagos.Show
    Case 2:
        ConexJardin.rsPagosJ.Filter = ""
        ConexJardin.rsPagosJ.Filter = "[tipopago]= 'Matricula'"
        LPagos.Show
    Case 3:
        ConexJardin.rsPagosJ.Filter = ""
        ConexJardin.rsPagosJ.Filter = "[tipopago]= 'Otros'"
        LPagos.Show
    Case 4:
        If DateValue(fecha4.Value) > DateValue(fechapagos.Value) Then
            MsgBox "La fecha inicial debe ser menor a la final!", vbInformation, "Reportes"
            Exit Sub
        End If
        ConexJardin.rsPagosJ.Filter = ""
        ConexJardin.rsPagosJ.Filter = "[feccon]>=" & fecha4.Value & " And [feccon]<=" & fechapagos.Value
        LPagos.Show
        
    Case Else: MsgBox "Escoja un tipo de filtro para visualizar el reporte!", vbInformation, "Reportes"
End Select
End Sub

Private Sub sa_Click()
ConexJardin.rsGrados.Filter = ""
ConexJardin.rsGrados.Filter = "[grado]='SalaCuna'"
LGrados.Caption = "Listado de niños para el grado de Sala Cuna - Nivel A"
LGrados.Show
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Private Sub tipo_Click()
Select Case tipo.ListIndex
    Case 0:
        fechalista.Visible = False
        fecha2.Visible = False
        Doc.Visible = False
        Label2.Visible = False
        Label3.Visible = False
    Case 1:
        fechalista.Visible = True
        fecha2.Visible = True
        Label2.Visible = True
        Label3.Visible = True
        Doc.Visible = False
        fechalista.SetFocus
    Case 2:
        Doc.Visible = True
        fechalista.Visible = False
        fecha2.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Doc.SetFocus
End Select
End Sub

Private Sub ver_Click()
Me.PopupMenu grades, 2
End Sub

Private Sub vplis_Click()

ConexionBD reportes, "select * from listadodeespera where fecins=#" & fecins.Month & "/" & fecins.Day & "/" & fecins.Year & "#"
If bd.Recordset.RecordCount = 0 Then
    MsgBox "No se puede imprimir el Listado de Espera por que la fecha no coincide con ningún registro!", vbCritical, "Reportes"
    Exit Sub
End If
If bd.Recordset.RecordCount > 0 Then
    'llena la tabla en word y la deja lista para imprimir
    On Error Resume Next
    Set ListEsp = New Word.Application
    ListEsp.Documents.Open App.Path + "\listado.doc"
    'llena el tipo del documento , numero de documento, sexo,edad,parentesco familiar,barrio, telefono,tipo salud,nivel,sec salud
    bd.Recordset.MoveFirst
    ListEsp.ActiveDocument.Tables(1).Cell(Row:=1, Column:=4).Tables(1).Cell(Row:=2, Column:=2).Range.InsertAfter fecins.Day
    ListEsp.ActiveDocument.Tables(1).Cell(Row:=1, Column:=4).Tables(1).Cell(Row:=2, Column:=3).Range.InsertAfter fecins.Month
    ListEsp.ActiveDocument.Tables(1).Cell(Row:=1, Column:=4).Tables(1).Cell(Row:=2, Column:=4).Range.InsertAfter fecins.Year
    
    For i = 3 To bd.Recordset.RecordCount + 5 Step 2
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=2).Range.InsertAfter bd.Recordset!tipdoc
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=3).Range.InsertAfter bd.Recordset!numdoc
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=6).Range.InsertAfter bd.Recordset!sex
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=7).Range.InsertAfter Left(bd.Recordset!fecnac, 2)
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=8).Range.InsertAfter Mid(bd.Recordset!fecnac, 4, 2)
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=9).Range.InsertAfter Right(bd.Recordset!fecnac, 4)
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=10).Range.InsertAfter bd.Recordset!eda
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=11).Range.InsertAfter bd.Recordset!parfam
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=12).Range.InsertAfter bd.Recordset!dir
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=13).Range.InsertAfter bd.Recordset!bar
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=14).Range.InsertAfter bd.Recordset!tel
        If bd.Recordset!sal = "EPS" Then
            ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=15).Range.InsertAfter "X"
        ElseIf bd.Recordset!sal = "ARS" Then
            ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=16).Range.InsertAfter "X"
        End If
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=17).Range.InsertAfter bd.Recordset!niv
        If bd.Recordset!secrsal = "SI" Then
            ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=18).Range.InsertAfter "X"
        End If
        
        bd.Recordset.MoveNext
    Next i
    'llena los apellidos, nombres
    bd.Recordset.MoveFirst
    For i = 4 To bd.Recordset.RecordCount + 6
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i, Column:=4).Range.InsertAfter bd.Recordset!priape
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i, Column:=5).Range.InsertAfter bd.Recordset!prinom
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=4).Range.InsertAfter bd.Recordset!segape
        ListEsp.ActiveDocument.Tables(1).Cell(Row:=i + 1, Column:=5).Range.InsertAfter bd.Recordset!segnom
        i = i + 1
        bd.Recordset.MoveNext
    Next i

    'muestra word con todo
    ListEsp.Application.Visible = True
    'ListEsp.ActiveDocument.PrintOut
    'ListEsp.ActiveDocument.Close wdDoNotSaveChanges
    'ListEsp.Application.Quit
Else
    MsgBox "No hay niños en listado con esta fecha de inscripción!", vbInformation, "Reportes"
End If
End Sub

Private Sub vplis_MouseOver()
menu.estado.Panels(4).Text = "Vista previa del listado de Espera."
End Sub
