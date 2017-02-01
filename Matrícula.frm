VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form matricula 
   Caption         =   "Matrículas"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11565
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
   Icon            =   "Matrícula.frx":0000
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc docu 
      Height          =   330
      Left            =   7560
      Top             =   7320
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
      Left            =   11160
      Top             =   4560
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Entidad Remitente/Id Beneficiario"
      TabPicture(0)   =   "Matrícula.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "numfor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fecmat"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "continuar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "bd"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "xpgroupbox1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "bd1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Datos Básicos Estudiante"
      TabPicture(1)   =   "Matrícula.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "continuar1"
      Tab(1).Control(1)=   "xpgroupbox2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Seguridad Social en Salud"
      TabPicture(2)   =   "Matrícula.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "continuar2"
      Tab(2).Control(1)=   "xpgroupbox3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Ubicación Nucleo Familiar"
      TabPicture(3)   =   "Matrícula.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "xpgroupbox4"
      Tab(3).Control(1)=   "continuar3"
      Tab(3).Control(2)=   "Label34"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Información Proyecto"
      TabPicture(4)   =   "Matrícula.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "xpgroupbox5"
      Tab(4).Control(1)=   "continuar4"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Datos Personas Nucleo Familiar"
      TabPicture(5)   =   "Matrícula.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "xpgroupbox6"
      Tab(5).ControlCount=   1
      Begin MSAdodcLib.Adodc bd1 
         Height          =   330
         Left            =   8280
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
      Begin Jardin.xpgroupbox xpgroupbox6 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   42
         Top             =   840
         Width           =   9015
         _extentx        =   15901
         _extenty        =   10186
         font            =   "Matrícula.frx":04EA
         backcolor       =   -2147483633
         caption         =   "Nucleo Familiar"
         Begin MSComCtl2.DTPicker datof 
            Height          =   375
            Left            =   3960
            TabIndex        =   142
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   49938433
            CurrentDate     =   38171
         End
         Begin VB.TextBox datot 
            Height          =   285
            Left            =   2280
            TabIndex        =   141
            Tag             =   "2"
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSAdodcLib.Adodc bd2 
            Height          =   330
            Left            =   5400
            Top             =   360
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
         Begin VB.ComboBox datoc 
            Height          =   315
            Left            =   6600
            TabIndex        =   140
            Tag             =   "2"
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid malla 
            Height          =   2655
            Left            =   120
            TabIndex        =   139
            Top             =   960
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   4683
            _Version        =   393216
            Rows            =   9
            Cols            =   21
            FixedCols       =   0
            ScrollBars      =   1
         End
         Begin MSComCtl2.DTPicker fecdighojsir 
            Height          =   375
            Left            =   6960
            TabIndex        =   50
            Tag             =   "1"
            Top             =   4440
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   49938433
            CurrentDate     =   38117
         End
         Begin VB.TextBox nomdilfor 
            Height          =   285
            Left            =   2520
            MaxLength       =   25
            TabIndex        =   45
            Top             =   4440
            Width           =   1935
         End
         Begin VB.TextBox nomfundighojsir 
            Height          =   285
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   44
            Top             =   5040
            Width           =   1935
         End
         Begin VB.TextBox obs 
            Height          =   285
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   43
            Top             =   5040
            Width           =   1935
         End
         Begin JeweledBut.JeweledButton mostrar 
            Height          =   375
            Left            =   120
            TabIndex        =   143
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            TX              =   "Mostrar Datos..."
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
            MICON           =   "Matrícula.frx":0512
            BC              =   8438015
            FC              =   0
         End
         Begin VB.Label Label42 
            Caption         =   $"Matrícula.frx":0680
            Height          =   495
            Left            =   120
            TabIndex        =   146
            Top             =   3720
            Width           =   8655
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre quien Diligencio Formulario"
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   4440
            Width           =   2055
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Digito Hoja Sirbe"
            Height          =   195
            Left            =   4680
            TabIndex        =   48
            Top             =   4560
            Width           =   2010
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre Funcionario digito Hoja Sirbe"
            Height          =   495
            Left            =   120
            TabIndex        =   47
            Top             =   5040
            Width           =   2295
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Observaciones"
            Height          =   255
            Left            =   4680
            TabIndex        =   46
            Top             =   5040
            Width           =   1335
         End
      End
      Begin Jardin.xpgroupbox xpgroupbox5 
         Height          =   3735
         Left            =   -73560
         TabIndex        =   30
         Top             =   1080
         Width           =   6855
         _extentx        =   12091
         _extenty        =   6588
         font            =   "Matrícula.frx":0713
         backcolor       =   -2147483633
         caption         =   "Información"
         Begin VB.TextBox graasp 
            Height          =   285
            Left            =   4080
            TabIndex        =   151
            Top             =   3120
            Width           =   2055
         End
         Begin VB.TextBox edaninvivpapmad 
            Height          =   285
            Left            =   4080
            MaxLength       =   1
            TabIndex        =   35
            Top             =   1920
            Width           =   855
         End
         Begin VB.ComboBox ninvivperpadmadotr 
            Height          =   315
            ItemData        =   "Matrícula.frx":073B
            Left            =   4080
            List            =   "Matrícula.frx":074B
            TabIndex        =   34
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox cuinindurdia 
            Height          =   315
            ItemData        =   "Matrícula.frx":076A
            Left            =   4080
            List            =   "Matrícula.frx":0786
            TabIndex        =   33
            Top             =   2520
            Width           =   2415
         End
         Begin VB.ComboBox vivpermpadmad 
            Height          =   315
            ItemData        =   "Matrícula.frx":07EA
            Left            =   4080
            List            =   "Matrícula.frx":07FA
            TabIndex        =   32
            Top             =   1320
            Width           =   1695
         End
         Begin VB.ComboBox ninvivpapmam 
            Height          =   315
            ItemData        =   "Matrícula.frx":0824
            Left            =   4080
            List            =   "Matrícula.frx":082E
            TabIndex        =   31
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "El Niño Vive permanentemente con Papá y Mamá"
            Height          =   435
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "El niño vive permanentemente con "
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   960
            Width           =   3045
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuidado del Niño durante el día"
            Height          =   195
            Left            =   240
            TabIndex        =   39
            Top             =   2640
            Width           =   2700
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grado al que aspira"
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   3240
            Width           =   1695
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Edad del Niño Cuando Dejó de Vivir con Papá y Mamá"
            Height          =   495
            Left            =   240
            TabIndex        =   37
            Top             =   1920
            Width           =   3135
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "El niño(a) permanece con "
            Height          =   195
            Left            =   240
            TabIndex        =   36
            Top             =   1440
            Width           =   2265
         End
      End
      Begin Jardin.xpgroupbox xpgroupbox4 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   29
         Top             =   960
         Width           =   9015
         _extentx        =   15901
         _extenty        =   9340
         font            =   "Matrícula.frx":083A
         backcolor       =   -2147483633
         caption         =   "Datos Generales"
         Begin Jardin.xpgroupbox xpgroupbox11 
            Height          =   2175
            Left            =   4440
            TabIndex        =   124
            Top             =   240
            Width           =   4095
            _extentx        =   7223
            _extenty        =   3836
            font            =   "Matrícula.frx":0862
            backcolor       =   -2147483633
            caption         =   "Procedencia Nucleo Familiar"
            Begin VB.ComboBox mun 
               Height          =   315
               Left            =   1680
               TabIndex        =   133
               Top             =   720
               Width           =   2175
            End
            Begin VB.ComboBox dep 
               Height          =   315
               Left            =   1680
               TabIndex        =   132
               Top             =   240
               Width           =   2175
            End
            Begin MSComCtl2.DTPicker feclle 
               Height          =   375
               Left            =   1680
               TabIndex        =   130
               Tag             =   "1"
               Top             =   1680
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "mm/yyyy"
               Format          =   49938435
               CurrentDate     =   38170
            End
            Begin VB.TextBox pais 
               Height          =   285
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   128
               Tag             =   "1"
               Text            =   "Colombia"
               Top             =   1200
               Width           =   2175
            End
            Begin VB.Label Label72 
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Llegada a Bogotá"
               Height          =   435
               Left            =   120
               TabIndex        =   129
               Top             =   1680
               Width           =   1560
            End
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "País "
               Height          =   195
               Left            =   120
               TabIndex        =   127
               Top             =   1320
               Width           =   405
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Municipio "
               Height          =   195
               Left            =   120
               TabIndex        =   126
               Top             =   840
               Width           =   840
            End
            Begin VB.Label Label63 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Departamento "
               Height          =   195
               Left            =   120
               TabIndex        =   125
               Top             =   360
               Width           =   1290
            End
         End
         Begin Jardin.xpgroupbox xpgroupbox10 
            Height          =   1935
            Left            =   120
            TabIndex        =   116
            Top             =   3120
            Width           =   4935
            _extentx        =   8705
            _extenty        =   3413
            font            =   "Matrícula.frx":088A
            backcolor       =   -2147483633
            caption         =   "Caracteristicas de la Vivienda"
            Begin VB.ComboBox forpagviv 
               Height          =   315
               Left            =   2520
               TabIndex        =   131
               Tag             =   "1"
               Top             =   1320
               Width           =   2175
            End
            Begin VB.TextBox tenviv 
               Height          =   285
               Left            =   2520
               TabIndex        =   122
               Top             =   960
               Width           =   2175
            End
            Begin VB.TextBox conviv 
               Height          =   285
               Left            =   2520
               TabIndex        =   121
               Top             =   600
               Width           =   2175
            End
            Begin VB.TextBox tipviv 
               Height          =   285
               Left            =   2520
               TabIndex        =   120
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Forma de pago"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   123
               Top             =   1440
               Width           =   1290
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de la Vivienda"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   119
               Top             =   360
               Width           =   1635
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Condiciones de la Vivenda"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   118
               Top             =   720
               Width           =   2265
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tenencia de la Vivienda"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   117
               Top             =   1080
               Width           =   2040
            End
         End
         Begin VB.ComboBox est 
            Height          =   315
            ItemData        =   "Matrícula.frx":08B2
            Left            =   1920
            List            =   "Matrícula.frx":08BF
            TabIndex        =   114
            Tag             =   "1"
            Top             =   2520
            Width           =   855
         End
         Begin VB.TextBox tel 
            Height          =   285
            Left            =   1920
            TabIndex        =   113
            Top             =   2040
            Width           =   2175
         End
         Begin VB.TextBox bar 
            Height          =   285
            Left            =   1920
            TabIndex        =   111
            Top             =   1440
            Width           =   2175
         End
         Begin VB.ComboBox loc 
            Height          =   315
            Left            =   1920
            TabIndex        =   108
            Tag             =   "1"
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox dir 
            Height          =   285
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   106
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estrato"
            Height          =   195
            Left            =   120
            TabIndex        =   115
            Top             =   2640
            Width           =   600
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefono"
            Height          =   195
            Left            =   120
            TabIndex        =   112
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Barrio"
            Height          =   195
            Left            =   120
            TabIndex        =   110
            Top             =   1560
            Width           =   525
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Localidad"
            Height          =   195
            Left            =   120
            TabIndex        =   109
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección"
            Height          =   195
            Left            =   120
            TabIndex        =   107
            Top             =   480
            Width           =   795
         End
      End
      Begin Jardin.xpgroupbox xpgroupbox3 
         Height          =   4695
         Left            =   -72360
         TabIndex        =   17
         Top             =   1080
         Width           =   5175
         _extentx        =   9128
         _extenty        =   8281
         font            =   "Matrícula.frx":08CC
         backcolor       =   -2147483633
         caption         =   "Seguridad Social"
         Begin VB.ComboBox punsis 
            Height          =   315
            ItemData        =   "Matrícula.frx":08F4
            Left            =   2880
            List            =   "Matrícula.frx":0901
            TabIndex        =   71
            Tag             =   "1"
            Top             =   4080
            Width           =   1935
         End
         Begin VB.ComboBox afisegsocfam 
            Height          =   315
            ItemData        =   "Matrícula.frx":0923
            Left            =   2880
            List            =   "Matrícula.frx":092D
            TabIndex        =   22
            Top             =   600
            Width           =   1935
         End
         Begin VB.ComboBox regsegsocfam 
            Height          =   315
            ItemData        =   "Matrícula.frx":0948
            Left            =   2880
            List            =   "Matrícula.frx":0952
            TabIndex        =   21
            Tag             =   "1"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.ComboBox calbenfam 
            Height          =   315
            ItemData        =   "Matrícula.frx":097C
            Left            =   2880
            List            =   "Matrícula.frx":0986
            TabIndex        =   20
            Tag             =   "1"
            Top             =   2160
            Width           =   1935
         End
         Begin VB.ComboBox vinsecsalfam 
            Height          =   315
            ItemData        =   "Matrícula.frx":09B9
            Left            =   2880
            List            =   "Matrícula.frx":09C3
            Locked          =   -1  'True
            TabIndex        =   19
            Tag             =   "1"
            Top             =   3000
            Width           =   855
         End
         Begin VB.TextBox numficsis 
            Height          =   285
            Left            =   2880
            MaxLength       =   3
            TabIndex        =   18
            Tag             =   "1"
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Afiliado Seguridad Social"
            Height          =   495
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Regimen de Seguridad Social en salud"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Calidad del Beneficiario familiar"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Vinculado a la Secretaría de Salud "
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   3000
            Width           =   2535
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Puntaje Sisben"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   4200
            Width           =   1575
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Número de ficha Sisben"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   3600
            Width           =   1695
         End
      End
      Begin Jardin.xpgroupbox xpgroupbox2 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   16
         Top             =   960
         Width           =   9135
         _extentx        =   16113
         _extenty        =   9340
         font            =   "Matrícula.frx":09CF
         backcolor       =   -2147483633
         caption         =   "Datos Básicos niño(a)"
         Begin VB.ComboBox proaso 
            Height          =   315
            ItemData        =   "Matrícula.frx":09F7
            Left            =   5040
            List            =   "Matrícula.frx":0A40
            TabIndex        =   104
            Tag             =   "1"
            Top             =   1560
            Width           =   3855
         End
         Begin VB.ComboBox asiactcenedu 
            Height          =   315
            ItemData        =   "Matrícula.frx":0C89
            Left            =   6960
            List            =   "Matrícula.frx":0C93
            TabIndex        =   102
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox nivestalc 
            Height          =   315
            ItemData        =   "Matrícula.frx":0C9F
            Left            =   6960
            List            =   "Matrícula.frx":0CAF
            TabIndex        =   100
            Top             =   240
            Width           =   2055
         End
         Begin VB.ComboBox parjeffam 
            Height          =   315
            ItemData        =   "Matrícula.frx":0CE6
            Left            =   2400
            List            =   "Matrícula.frx":0D11
            TabIndex        =   98
            Top             =   4440
            Width           =   1935
         End
         Begin VB.ComboBox tipdisest 
            Height          =   315
            ItemData        =   "Matrícula.frx":0DB9
            Left            =   2520
            List            =   "Matrícula.frx":0DC9
            TabIndex        =   96
            Top             =   3840
            Width           =   2535
         End
         Begin VB.ComboBox depnac 
            Height          =   315
            ItemData        =   "Matrícula.frx":0DF1
            Left            =   2520
            List            =   "Matrícula.frx":0E4F
            Sorted          =   -1  'True
            TabIndex        =   94
            Tag             =   "1"
            Top             =   1920
            Width           =   2415
         End
         Begin VB.ComboBox mandis 
            Height          =   315
            ItemData        =   "Matrícula.frx":0F71
            Left            =   2520
            List            =   "Matrícula.frx":0F7B
            TabIndex        =   90
            Tag             =   "1"
            Top             =   3360
            Width           =   855
         End
         Begin VB.TextBox Painac 
            Height          =   285
            Left            =   2520
            TabIndex        =   89
            Tag             =   "1"
            Text            =   "Colombia"
            Top             =   2880
            Width           =   2415
         End
         Begin VB.ComboBox munnac 
            Height          =   315
            Left            =   2520
            TabIndex        =   88
            Tag             =   "1"
            Top             =   2400
            Width           =   2415
         End
         Begin VB.TextBox edad 
            Height          =   285
            Left            =   1800
            MaxLength       =   1
            TabIndex        =   87
            Top             =   1440
            Width           =   735
         End
         Begin MSComCtl2.DTPicker fecnac 
            Height          =   375
            Left            =   1800
            TabIndex        =   85
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   49938433
            CurrentDate     =   38170
         End
         Begin VB.TextBox sexo 
            Height          =   285
            Left            =   1800
            TabIndex        =   83
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Problemas asociados"
            Height          =   195
            Left            =   5040
            TabIndex        =   105
            Top             =   1320
            Width           =   1800
         End
         Begin VB.Label Label69 
            BackStyle       =   0  'Transparent
            Caption         =   "Asiste Actualmente al Centro Educativo"
            Height          =   435
            Left            =   5040
            TabIndex        =   103
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estudios alcanzados"
            Height          =   195
            Left            =   5040
            TabIndex        =   101
            Top             =   360
            Width           =   1710
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Parentesco con el Jefe Familiar"
            Height          =   375
            Left            =   120
            TabIndex        =   99
            Top             =   4440
            Width           =   2055
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Discapacidad"
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Top             =   3960
            Width           =   1815
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento Nacimiento"
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   2040
            Width           =   2235
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Manifiesta discapacidad"
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   3480
            Width           =   2010
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "País Nacimiento"
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   3000
            Width           =   1350
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Municipio Nacimiento"
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   2520
            Width           =   1785
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edad aproximada"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   1560
            Width           =   1500
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nacimiento"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   960
            Width           =   1500
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo:"
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   480
            Width           =   510
         End
      End
      Begin Jardin.xpgroupbox xpgroupbox1 
         Height          =   4095
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   8655
         _extentx        =   15266
         _extenty        =   7223
         font            =   "Matrícula.frx":0F87
         backcolor       =   -2147483633
         caption         =   "Información General"
         Begin VB.ComboBox col 
            Height          =   315
            Left            =   2040
            TabIndex        =   149
            Tag             =   "1"
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox prorem 
            Height          =   285
            Left            =   6120
            MaxLength       =   20
            TabIndex        =   138
            Tag             =   "1"
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox rempor 
            Height          =   315
            Left            =   2040
            TabIndex        =   137
            Tag             =   "1"
            Top             =   2880
            Width           =   1815
         End
         Begin VB.ComboBox persolser 
            Height          =   315
            Left            =   2040
            TabIndex        =   136
            Tag             =   "1"
            Top             =   2280
            Width           =   1815
         End
         Begin VB.ComboBox uniope 
            Height          =   315
            Left            =   2040
            TabIndex        =   135
            Tag             =   "1"
            Top             =   720
            Width           =   1815
         End
         Begin Jardin.xpgroupbox xpgroupbox9 
            Height          =   2535
            Left            =   4200
            TabIndex        =   74
            Top             =   1200
            Width           =   4215
            _extentx        =   7435
            _extenty        =   4471
            font            =   "Matrícula.frx":0FAF
            backcolor       =   -2147483633
            caption         =   "Beneficiario niño(a)"
            Begin VB.TextBox tipdocnin 
               Height          =   285
               Left            =   1920
               TabIndex        =   134
               Top             =   840
               Width           =   2055
            End
            Begin VB.TextBox apenin 
               Height          =   285
               Left            =   1200
               TabIndex        =   81
               Top             =   1800
               Width           =   2775
            End
            Begin VB.TextBox nomnin 
               Height          =   285
               Left            =   1200
               TabIndex        =   79
               Top             =   1320
               Width           =   2775
            End
            Begin VB.ComboBox numdoc 
               Height          =   315
               Left            =   1920
               TabIndex        =   75
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label Label67 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Apellidos"
               Height          =   195
               Left            =   120
               TabIndex        =   80
               Top             =   1920
               Width           =   765
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nombres"
               Height          =   195
               Left            =   120
               TabIndex        =   78
               Top             =   1440
               Width           =   765
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Documento"
               Height          =   195
               Left            =   120
               TabIndex        =   77
               Top             =   960
               Width           =   1395
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Documento "
               Height          =   195
               Left            =   120
               TabIndex        =   76
               Top             =   480
               Width           =   1035
            End
         End
         Begin VB.ComboBox modal 
            Height          =   315
            ItemData        =   "Matrícula.frx":0FD7
            Left            =   2040
            List            =   "Matrícula.frx":0FDE
            TabIndex        =   70
            Tag             =   "1"
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox submod 
            Height          =   285
            Left            =   2040
            MaxLength       =   12
            TabIndex        =   7
            Tag             =   "1"
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox entrem 
            Height          =   285
            Left            =   6240
            MaxLength       =   30
            TabIndex        =   6
            Tag             =   "1"
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uni Operativa"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modalidad"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub modalidad"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1800
            Width           =   1275
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Quien solicita el servicio"
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remitido por"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proyecto que remite"
            Height          =   195
            Left            =   4200
            TabIndex        =   10
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entidad que remite"
            Height          =   195
            Left            =   4200
            TabIndex        =   9
            Top             =   840
            Width           =   1620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Col"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   285
         End
      End
      Begin MSAdodcLib.Adodc bd 
         Height          =   330
         Left            =   6720
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
      Begin VB.ComboBox nivestfam 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Matrícula.frx":0FEB
         Left            =   -2160
         List            =   "Matrícula.frx":1001
         TabIndex        =   1
         Top             =   9840
         Width           =   1335
      End
      Begin JeweledBut.JeweledButton continuar 
         Height          =   375
         Left            =   7800
         TabIndex        =   65
         Top             =   6120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         TX              =   "&Continuar"
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
         MPTR            =   0
         MICON           =   "Matrícula.frx":1069
         BC              =   8438015
         FC              =   0
      End
      Begin JeweledBut.JeweledButton continuar1 
         Height          =   375
         Left            =   -66960
         TabIndex        =   66
         Top             =   6360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         TX              =   "&Continuar"
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
         MPTR            =   0
         MICON           =   "Matrícula.frx":1085
         BC              =   8438015
         FC              =   0
      End
      Begin JeweledBut.JeweledButton continuar2 
         Height          =   375
         Left            =   -66960
         TabIndex        =   67
         Top             =   6360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         TX              =   "&Continuar"
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
         MPTR            =   0
         MICON           =   "Matrícula.frx":10A1
         BC              =   8438015
         FC              =   0
      End
      Begin JeweledBut.JeweledButton continuar3 
         Height          =   375
         Left            =   -66960
         TabIndex        =   68
         Top             =   6360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         TX              =   "&Continuar"
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
         MPTR            =   0
         MICON           =   "Matrícula.frx":10BD
         BC              =   8438015
         FC              =   0
      End
      Begin JeweledBut.JeweledButton continuar4 
         Height          =   375
         Left            =   -67800
         TabIndex        =   69
         Top             =   4920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         TX              =   "&Continuar"
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
         MPTR            =   0
         MICON           =   "Matrícula.frx":10D9
         BC              =   8438015
         FC              =   0
      End
      Begin VB.Label fecmat 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   7080
         TabIndex        =   148
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Matrícula"
         Height          =   195
         Left            =   5520
         TabIndex        =   147
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label numfor 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   145
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Matricula"
         Height          =   195
         Left            =   240
         TabIndex        =   73
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label34 
         Caption         =   "Label34"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   -73920
         TabIndex        =   4
         Top             =   4440
         Width           =   1575
      End
   End
   Begin Jardin.xpgroupbox xpgroupbox7 
      Height          =   3615
      Left            =   9720
      TabIndex        =   51
      Top             =   360
      Width           =   1695
      _extentx        =   2990
      _extenty        =   6376
      font            =   "Matrícula.frx":10F5
      backcolor       =   -2147483633
      caption         =   "Opciones"
      Begin JeweledBut.JeweledButton nuevo 
         Height          =   375
         Left            =   120
         TabIndex        =   52
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
         MICON           =   "Matrícula.frx":111D
         BC              =   8438015
         FC              =   0
         Picture         =   "Matrícula.frx":128B
      End
      Begin JeweledBut.JeweledButton busqueda 
         Height          =   375
         Left            =   120
         TabIndex        =   53
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
         MICON           =   "Matrícula.frx":3D95
         BC              =   8438015
         FC              =   0
         Picture         =   "Matrícula.frx":3F03
      End
      Begin JeweledBut.JeweledButton eliminar 
         Height          =   375
         Left            =   120
         TabIndex        =   54
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
         MICON           =   "Matrícula.frx":405D
         BC              =   8438015
         FC              =   0
         Picture         =   "Matrícula.frx":41CB
      End
      Begin JeweledBut.JeweledButton guardar 
         Height          =   375
         Left            =   120
         TabIndex        =   55
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
         MICON           =   "Matrícula.frx":4765
         BC              =   8438015
         FC              =   0
         Picture         =   "Matrícula.frx":48D3
      End
      Begin JeweledBut.JeweledButton Actualizar 
         Height          =   375
         Left            =   120
         TabIndex        =   56
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
         MICON           =   "Matrícula.frx":A501
         BC              =   8438015
         FC              =   0
         Picture         =   "Matrícula.frx":A66F
      End
      Begin JeweledBut.JeweledButton modificar 
         Height          =   375
         Left            =   120
         TabIndex        =   57
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
         MICON           =   "Matrícula.frx":AC09
         BC              =   8438015
         FC              =   0
         Picture         =   "Matrícula.frx":AD77
      End
      Begin JeweledBut.JeweledButton parametro 
         Height          =   375
         Left            =   120
         TabIndex        =   150
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
         MICON           =   "Matrícula.frx":AED1
         BC              =   8438015
         FC              =   0
      End
   End
   Begin Jardin.xpgroupbox xpgroupbox8 
      Height          =   855
      Left            =   120
      TabIndex        =   58
      Top             =   7080
      Width           =   6855
      _extentx        =   12091
      _extenty        =   1508
      font            =   "Matrícula.frx":B03F
      backcolor       =   -2147483633
      caption         =   "Navegación"
      Begin JeweledBut.JeweledButton primero 
         Height          =   375
         Left            =   120
         TabIndex        =   59
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
         MICON           =   "Matrícula.frx":B067
         BC              =   8438015
         FC              =   0
         Picture         =   "Matrícula.frx":B1D5
      End
      Begin JeweledBut.JeweledButton siguiente 
         Height          =   375
         Left            =   3480
         TabIndex        =   60
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
         MICON           =   "Matrícula.frx":B32F
         BC              =   8438015
         FC              =   0
         Picture         =   "Matrícula.frx":B49D
      End
      Begin JeweledBut.JeweledButton ultimo 
         Height          =   375
         Left            =   5160
         TabIndex        =   61
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
         MICON           =   "Matrícula.frx":B5F7
         BC              =   8438015
         FC              =   0
         Picture         =   "Matrícula.frx":B765
      End
      Begin JeweledBut.JeweledButton anterior 
         Height          =   375
         Left            =   1800
         TabIndex        =   62
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
         MICON           =   "Matrícula.frx":B8BF
         BC              =   8438015
         FC              =   0
         Picture         =   "Matrícula.frx":BA2D
      End
   End
   Begin JeweledBut.JeweledButton salir 
      Height          =   375
      Left            =   9840
      TabIndex        =   63
      Top             =   7440
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
      MICON           =   "Matrícula.frx":BB87
      BC              =   8438015
      FC              =   0
      Picture         =   "Matrícula.frx":BCF5
   End
   Begin JeweledBut.JeweledButton cancelar 
      Height          =   375
      Left            =   9840
      TabIndex        =   152
      Top             =   6960
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
      MICON           =   "Matrícula.frx":BE4F
      BC              =   8438015
      FC              =   0
      Picture         =   "Matrícula.frx":BFBD
   End
   Begin VB.Label numreg 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   9840
      TabIndex        =   144
      Top             =   5640
      Width           =   60
   End
   Begin VB.Label Label65 
      BackColor       =   &H80000018&
      Caption         =   "Formato Fecha:   dia/mes/año"
      Height          =   435
      Left            =   9720
      TabIndex        =   72
      Top             =   4560
      Width           =   1395
   End
   Begin VB.Label ma 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   9840
      TabIndex        =   64
      Top             =   5280
      Width           =   60
   End
   Begin VB.Label Label33 
      Caption         =   "País donde nació"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label21 
      Caption         =   "Tipo de Discapacidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
End
Attribute VB_Name = "matricula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Integer
Function mostrarcampos()
'muestra todos los campos de matricula
    numreg = bd.Recordset.AbsolutePosition & " registro."
    numfor = bd.Recordset!numfor
    col.Text = bd.Recordset!col
    uniope.Text = bd.Recordset!uniope
    modal.Text = bd.Recordset!modal
    submod.Text = bd.Recordset!submod
    fecmat = bd.Recordset!fecmat
    persolser.Text = bd.Recordset!persolser
    rempor.Text = bd.Recordset!rempor
    prorem.Text = bd.Recordset!prorem
    entrem.Text = bd.Recordset!entrem
    numdoc.Text = bd.Recordset!numdoc
    depnac.Text = bd.Recordset!depnac
    munnac.Text = bd.Recordset!munnac
    Painac.Text = bd.Recordset!Painac
    tipdisest.Text = bd.Recordset!tipdisest
    nivestalc.Text = bd.Recordset!niveduben
    asiactcenedu.Text = bd.Recordset!asiactcenedu
    proaso.Text = bd.Recordset!proaso
    afisegsocfam.Text = bd.Recordset!afisegsocben
    regsegsocfam.Text = bd.Recordset!regsegsocben
    calbenfam.Text = bd.Recordset!calben
    vinsecsalfam.Text = bd.Recordset!vinsecsalben
    numficsis.Text = bd.Recordset!numficsis
    punsis.Text = bd.Recordset!punsis
    loc.Text = bd.Recordset!loc
    forpagviv.Text = bd.Recordset!forpagviv
    dep.Text = bd.Recordset!dptoprofam
    mun.Text = bd.Recordset!munprofam
    pais.Text = bd.Recordset!paiprofam
    feclle.Value = bd.Recordset!fecllebogfam
    ninvivpapmam.Text = bd.Recordset!ninvivpapmam
    ninvivperpadmadotr.Text = bd.Recordset!ninvivperpadmadotr
    vivpermpadmad.Text = bd.Recordset!vivperpadmad
    edaninvivpapmad.Text = bd.Recordset!edaninvivpapmad
    cuinindurdia.Text = bd.Recordset!cuinindurdia
    graasp.Text = bd.Recordset!graasp
    If IsNull(bd.Recordset!nomdilform) = False Then
        nomdilfor.Text = bd.Recordset!nomdilform
    End If
    nomfundighojsir.Text = bd.Recordset!nomfundighojsir
    fecdighojsir.Value = bd.Recordset!fecdighojsir
    obs.Text = bd.Recordset!obs
    SSTab1.Tab = 0
    'muestra los campos de listado de espera
    ConexionBD1 matricula, "select * from listadodeespera where numdoc='" & numdoc.Text & "'"
    If BD1.Recordset.RecordCount > 0 Then
        tipdocnin.Text = BD1.Recordset!tipdoc
        nomnin.Text = BD1.Recordset!prinom & " " & BD1.Recordset!segnom
        apenin.Text = BD1.Recordset!priape & " " & BD1.Recordset!segape
        sexo.Text = BD1.Recordset!sex
        If IsNull(BD1.Recordset!fecnac) = False Then
            fecnac = BD1.Recordset!fecnac
        End If
        edad.Text = BD1.Recordset!eda
        parjeffam.Text = BD1.Recordset!parfam
        dir.Text = BD1.Recordset!dir
        bar.Text = BD1.Recordset!bar
        tel.Text = BD1.Recordset!tel
    End If
        'conectamos bd para cargar datos de inscripciones
        ConexionBD1 matricula, "select * from inscripciones where numdoc='" & numdoc.Text & "'"
        If BD1.Recordset.RecordCount > 0 Then
            tipviv.Text = BD1.Recordset!tipviv
            conviv.Text = BD1.Recordset!conviv
            tenviv.Text = BD1.Recordset!tenviv
        End If
End Function
Private Sub act_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    posocu.SetFocus
End If
End Sub



Private Sub Actualizar_Click()
nuevo.Enabled = True
modificar.Enabled = True
eliminar.Enabled = True
primero.Enabled = True
ultimo.Enabled = True
siguiente.Enabled = True
anterior.Enabled = True
parametro.Enabled = True
Actualizar.Enabled = False
busqueda.Enabled = True
'modifica el registro
If ModificadoM = True Then
    MODIFICARR
    'MODIFICARR
    ModificadoM = False
End If
malla.Clear
bd.Refresh
SSTab1.Tab = 0
End Sub
Sub MODIFICARR()
'ConexionBD1 matricula, "select * from matricula where numdoc='" & numdoc.Text & "'"
Dim modii
modii = bd.Recordset.EditMode
bd.Recordset!numfor = numfor
    bd.Recordset!col = col.Text
    bd.Recordset!uniope = uniope.Text
    bd.Recordset!modal = modal.Text
    bd.Recordset!submod = submod.Text
    bd.Recordset!fecmat = fecmat
    bd.Recordset!persolser = persolser.Text
    bd.Recordset!rempor = rempor.Text
    bd.Recordset!prorem = prorem.Text
    bd.Recordset!entrem = entrem.Text
    bd.Recordset!numdoc = numdoc.Text
    bd.Recordset!depnac = depnac.Text
    bd.Recordset!munnac = munnac.Text
    bd.Recordset!Painac = Painac.Text
    bd.Recordset!tipdisest = tipdisest.Text
    bd.Recordset!niveduben = nivestalc.Text
    bd.Recordset!asiactcenedu = asiactcenedu.Text
    bd.Recordset!proaso = proaso.Text
    bd.Recordset!afisegsocben = afisegsocfam.Text
    bd.Recordset!regsegsocben = regsegsocfam.Text
    bd.Recordset!calben = calbenfam.Text
    bd.Recordset!vinsecsalben = vinsecsalfam.Text
    bd.Recordset!numficsis = numficsis.Text
    bd.Recordset!punsis = punsis.Text
    bd.Recordset!loc = loc.Text
    bd.Recordset!forpagviv = forpagviv.Text
    bd.Recordset!dptoprofam = dep.Text
    bd.Recordset!munprofam = mun.Text
    bd.Recordset!paiprofam = pais.Text
    bd.Recordset!fecllebogfam = feclle.Value
    bd.Recordset!ninvivpapmam = ninvivpapmam.Text
    bd.Recordset!ninvivperpadmadotr = ninvivperpadmadotr.Text
    bd.Recordset!vivperpadmad = vivpermpadmad.Text
    bd.Recordset!edaninvivpapmad = edaninvivpapmad.Text
    bd.Recordset!cuinindurdia = cuinindurdia.Text
    bd.Recordset!graasp = graasp.Text
    bd.Recordset!nomdilform = nomdilfor.Text
    bd.Recordset!nomfundighojsir = nomfundighojsir.Text
    bd.Recordset!fecdighojsir = fecdighojsir.Value
    bd.Recordset!obs = obs.Text
    bd.Recordset!mat = 1
bd.Recordset.Update
'borran de la tabla hermanos todos los registros que correspondan al numdoc
'ConexionBD2 matricula, "SELECT * FROM HERMANOS"
ConexionBD1 matricula, "delete from hermanos where numdoc='" & numdoc.Text & "'"
'graba los registros de nuevo que han sido cambiados
'guardar en la tabla hermanos
ConexionBD2 matricula, "select * from hermanos"
For i = 1 To 8 'filas
    If malla.TextMatrix(i, 0) <> "" Then
        bd2.Recordset.AddNew
            bd2.Recordset!numdoc = numdoc.Text
            bd2.Recordset!tipdocher = malla.TextMatrix(i, 0)
            bd2.Recordset!numdocher = malla.TextMatrix(i, 1)
            bd2.Recordset!apeher = malla.TextMatrix(i, 2)
            bd2.Recordset!nomher = malla.TextMatrix(i, 3)
            bd2.Recordset!sexher = malla.TextMatrix(i, 4)
            bd2.Recordset!fecnacher = malla.TextMatrix(i, 5)
            bd2.Recordset!edaproher = Val(malla.TextMatrix(i, 6))
            bd2.Recordset!estcivher = malla.TextMatrix(i, 7)
            bd2.Recordset!tipdisher = malla.TextMatrix(i, 8)
            bd2.Recordset!parher = malla.TextMatrix(i, 9)
            bd2.Recordset!esther = malla.TextMatrix(i, 10)
            bd2.Recordset!asiacteduher = malla.TextMatrix(i, 11)
            bd2.Recordset!actocuher = malla.TextMatrix(i, 12)
            bd2.Recordset!posocuher = malla.TextMatrix(i, 13)
            bd2.Recordset!ingmesher = malla.TextMatrix(i, 14)
            bd2.Recordset!forperingher = malla.TextMatrix(i, 15)
            bd2.Recordset!afisalher = malla.TextMatrix(i, 16)
            bd2.Recordset!regsegsalher = malla.TextMatrix(i, 17)
            bd2.Recordset!calbenher = malla.TextMatrix(i, 18)
            bd2.Recordset!vinsecsalher = malla.TextMatrix(i, 19)
            bd2.Recordset!agrviointher = malla.TextMatrix(i, 20)
        bd2.Recordset.Update
    End If
Next i


End Sub

Private Sub bd_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
'fCancelDisplay = True
End Sub

Private Sub bd1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub busqueda_Click()
MB.Formulario = Me.Name
MB.Descripcion = "Matriculas"
elBuscador.Show
End Sub



Private Sub calben_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    numficsis.SetFocus
End If
End Sub

Private Sub afisegsocben_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    regsegsociben.SetFocus
End If
End Sub

Private Sub afisegsocfam_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    regsegsocfam.SetFocus
End If
End Sub

Private Sub agrviointfam_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    afisegsocben.SetFocus
End If
End Sub

Private Sub anterior_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MovePrevious
    If bd.Recordset.BOF Then
        bd.Recordset.MoveFirst
    End If
    mostrarcampos
    nd = numdoc.Text
End If
numdoc.SetFocus
End Sub

Private Sub apredufam_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    edafam.SetFocus
End If
End Sub

Private Sub asiactcenedu_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    proaso.SetFocus
End If
End Sub

Private Sub asiactcenedufam_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    apredufam.SetFocus
End If
End Sub

Private Sub calbenfam_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    vinsecsalfam.SetFocus
End If
End Sub

Private Sub cancelar_Click()
If bd.Recordset.RecordCount > 0 Then
    mostrarcampos
    malla.Clear
End If
If NuevoRegM = True Then
    nuevo.Enabled = True
    modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    anterior.Enabled = True
    siguiente.Enabled = True
    ultimo.Enabled = True
    guardar.Enabled = False
    busqueda.Enabled = True
    NuevoRegM = False
ElseIf ModificadoM = True Then
    nuevo.Enabled = True
    modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    ultimo.Enabled = True
    siguiente.Enabled = True
    anterior.Enabled = True
    Actualizar.Enabled = False
    busqueda.Enabled = True
    ModificadoM = False
End If
parametro.Enabled = True
End Sub

Private Sub col_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    uniope.SetFocus
End If
End Sub
Private Sub Combo6_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub
Private Sub Combo7_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    asiactcenedu.SetFocus
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)

End Sub

Private Sub continuar_Click()
SSTab1.Tab = 1
sexo.SetFocus
End Sub

Private Sub continuar1_Click()
SSTab1.Tab = 2
afisegsocfam.SetFocus
End Sub

Private Sub continuar2_Click()
SSTab1.Tab = 3
loc.SetFocus
End Sub

Private Sub continuar3_Click()
SSTab1.Tab = 4
ninvivpapmam.SetFocus
End Sub

Private Sub continuar4_Click()
SSTab1.Tab = 5
datoc.Visible = False
datot.Visible = False
End Sub

Private Sub cuinindurdia_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    continuar4_Click
End If
End Sub

Private Sub DATAHER_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub



Private Sub datoc_Change()
malla.TextMatrix(malla.RowSel, malla.ColSel) = datoc.Text
End Sub

Private Sub datoc_Click()
malla.TextMatrix(malla.RowSel, malla.ColSel) = datoc.List(datoc.ListIndex)
End Sub

Private Sub datof_Change()
malla.TextMatrix(malla.RowSel, malla.ColSel) = datof.Value
End Sub

Private Sub datof_LostFocus()
Dim ANIOs As Integer
ANIOs = Format(Date, "yyyy")

If datof.Year > ANIOs Then
    MsgBox "Fecha no permitida!", vbExclamation, "Matricula"
    datof.Visible = True
    datof.SetFocus
    Exit Sub
End If

End Sub

Private Sub datot_Change()
malla.TextMatrix(malla.RowSel, malla.ColSel) = datot.Text
End Sub

Private Sub dep_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    mun.SetFocus
End If
End Sub

Private Sub depnac_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    munnac.SetFocus
End If
End Sub

Private Sub depprofam_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    paiprofam.SetFocus
End If
End Sub
Private Sub docide_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sexfam.SetFocus
End If
End Sub

Private Sub edaninvivpapmam_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub

Private Sub edafam_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    fecnacfam.SetFocus
End If
End Sub

Private Sub edad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    depnac.SetFocus
End If
End Sub

Private Sub edaninvivpapmad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    cuinindurdia.SetFocus
End If
End Sub

Private Sub eliminar_Click()
If bd.Recordset.RecordCount > 0 Then
If MsgBox("Está seguro de querer eliminar el registro?", vbYesNo + vbQuestion, "Eliminar Registro") = vbYes Then
   'modifica el registro en inscripciones para asiganarle que ya esta matrriculado
   On Error Resume Next
    ConexionBD1 matricula, "select * from inscripciones where numdoc='" & bd.Recordset!numdoc & "'"
    Dim modi
    modi = BD1.Recordset.EditMode
    BD1.Recordset!Matriculado = 0
    BD1.Recordset.Update
   On Error Resume Next
   ConexionBD1 matricula, "delete from hermanos where numdoc='" & numdoc.Text & "'"
   bd.Recordset.Delete
   If bd.Recordset.RecordCount > 0 Then
        bd.Recordset.MoveFirst
        bd.Refresh
        ma.Caption = bd.Recordset.RecordCount & " Matriculados"
        mostrarcampos
   Else
        Unload Me
        matricula.Show
   End If
End If
'llenar numero documento niño
ConexionBD1 matricula, "select numdoc from inscripciones where matriculado=0"
numdoc.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        numdoc.AddItem BD1.Recordset!numdoc
        BD1.Recordset.MoveNext
    Next i
End If

End If
End Sub

Private Sub entrem_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    numdoc.SetFocus
End If
End Sub
Private Sub est_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    forpagviv.SetFocus
End If
End Sub
Private Sub estciv_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    tipdisfam.SetFocus
End If
End Sub
Private Sub fecdighojsir_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    obs.SetFocus
End If
End Sub

Private Sub fecllebogfam_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    est.SetFocus
End If
End Sub

Private Sub fecdighojsir_LostFocus()
Dim ANIOs As Integer
ANIOs = Format(Date, "yyyy")

If fecdighojsir.Year > ANIOs Then
    MsgBox "Fecha no permitida!", vbExclamation, "Matricula"
    fecdighojsir.SetFocus
    Exit Sub
End If
End Sub

Private Sub feclle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyAscii = 13 Then
    continuar3.SetFocus
End If
End Sub

Private Sub fecnacfam_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub

Private Sub fecnacfam_LostFocus()
SSTab1.Tab = 2
End Sub
Private Sub LlenarCombos()
menu.estado.Panels(4).Text = "Cargando..."
'llenar col
ConexionBD1 matricula, "select * from parametrizacion where tippar=1" & " order by dato;"
col.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        col.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar unidad operativa
ConexionBD1 matricula, "select * from parametrizacion where tippar=24" & " order by dato;"
uniope.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        uniope.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar solicitud servicio
ConexionBD1 matricula, "select * from parametrizacion where tippar=25" & " order by dato;"
persolser.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        persolser.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar remitido por:
ConexionBD1 matricula, "select * from parametrizacion where tippar=26" & " order by dato;"
rempor.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        rempor.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar forma de pago
ConexionBD1 matricula, "select * from parametrizacion where tippar=27" & " order by dato;"
forpagviv.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        forpagviv.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar departamento
ConexionBD1 matricula, "select * from parametrizacion where tippar=14" & " order by dato;"
dep.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        dep.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar municipio
ConexionBD1 matricula, "select * from parametrizacion where tippar=18" & " order by dato;"
mun.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        mun.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If

'llenar modalidad
ConexionBD1 matricula, "select * from parametrizacion where tippar=17" & " order by dato;"
modal.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        modal.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar numero documento niño
ConexionBD1 matricula, "select numdoc from inscripciones where matriculado=0"
numdoc.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        numdoc.AddItem BD1.Recordset!numdoc
        BD1.Recordset.MoveNext
    Next i
End If

'llenar departamento
ConexionBD1 matricula, "select * from parametrizacion where tippar=14" & " order by dato;"
depnac.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        depnac.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar municipio
ConexionBD1 matricula, "select * from parametrizacion where tippar=18" & " order by dato;"
munnac.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        munnac.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar tipo de discapacidad
ConexionBD1 matricula, "select * from parametrizacion where tippar=2" & " order by dato;"
tipdisest.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        tipdisest.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar problemas asociados
ConexionBD1 matricula, "select * from parametrizacion where tippar=21" & " order by dato;"
proaso.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        proaso.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar localidad
ConexionBD1 matricula, "select * from parametrizacion where tippar=4" & " order by dato;"
loc.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        loc.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If

End Sub

Private Sub feclle_LostFocus()
Dim ANIOs As Integer
ANIOs = Format(Date, "yyyy")

If feclle.Year > ANIOs Then
    MsgBox "Fecha no permitida!", vbExclamation, "Matricula"
    feclle.SetFocus
    Exit Sub
End If
End Sub

Private Sub fecnac_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    edad.SetFocus
End If
End Sub

Private Sub LlenarMalla()
For i = 0 To 20
    malla.ColWidth(i) = 2000
Next i

malla.TextMatrix(0, 0) = "Tipo ID"
malla.TextMatrix(0, 1) = "Num. Doc."
malla.TextMatrix(0, 2) = "Apellidos"
malla.TextMatrix(0, 3) = "Nombres"
malla.ColWidth(4) = 1000
malla.TextMatrix(0, 4) = "Sexo"
malla.TextMatrix(0, 5) = "Fecha Nac."
malla.TextMatrix(0, 6) = "Edad Aprox."
malla.TextMatrix(0, 7) = "Estado Civil"
malla.TextMatrix(0, 8) = "Tipo Discapacidad"
malla.TextMatrix(0, 9) = "Parentesco"
malla.TextMatrix(0, 10) = "Nivel Estudio"
malla.ColWidth(11) = 1000
malla.TextMatrix(0, 11) = "Asiste Centro Edu."
malla.TextMatrix(0, 12) = "Actividad"
malla.TextMatrix(0, 13) = "Posición Ocupacional"
malla.TextMatrix(0, 14) = "Ingresos"
malla.TextMatrix(0, 15) = "Forma Percibir Ingresos"
malla.ColWidth(16) = 1000
malla.TextMatrix(0, 16) = "Afiliado Salud"
malla.TextMatrix(0, 17) = "Reg. Seg. Salud"
malla.TextMatrix(0, 18) = "Calidad Beneficiario"
malla.ColWidth(19) = 1000
malla.TextMatrix(0, 19) = "Vinculado Sec. Salud"
malla.ColWidth(20) = 1000
malla.TextMatrix(0, 20) = "Agresor Violencia IntraFam."
End Sub
Private Sub Form_Activate()
datoc.Visible = False
FormularioActivo = True
LlenarMalla
fecmat = Format(Date, "dd/mm/yyyy")
malla.Row = 1
malla.col = 1

ma.Caption = bd.Recordset.RecordCount & " Matriculados"
numdoc.SetFocus
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If NuevoRegM = True Then
    If MsgBox("Esta agregando un nuevo registro" & vbCrLf & "Desea continuar?", vbYesNo + vbQuestion, "Matricula") = vbYes Then
        Cancel = True
    Else
        
        NuevoRegM = False
    End If
End If
End Sub

Private Sub Form_Resize()
Me.WindowState = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormularioActivo = False
menu.estado.Visible = True
menu.estado.Panels(4).Text = "Menú Principal"
End Sub

Private Sub forpagviv_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    dep.SetFocus
End If
End Sub

Private Sub forpering_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    nomdilfor.SetFocus
End If
End Sub


Private Sub graasp_KeyPress(KeyAscii As Integer)
KeyAscii = 0
KeyAscii = Validar_letra(KeyAscii)

End Sub


Private Sub guardar_Click()
guardarregistro
busqueda.Enabled = True
bd.Refresh
ma.Caption = bd.Recordset.RecordCount & " Matriculados"
eliminar.Enabled = True
NuevoRegM = False
SSTab1.Tab = 0
End Sub
Function guardarregistro()
Dim NUM_DOC
'On Error Resume Next
NUM_DOC = numdoc.Text
bd.Recordset.AddNew
    bd.Recordset!numfor = numfor
    bd.Recordset!col = col.Text
    bd.Recordset!uniope = uniope.Text
    bd.Recordset!modal = modal.Text
    bd.Recordset!submod = submod.Text
    bd.Recordset!fecmat = fecmat
    bd.Recordset!persolser = persolser.Text
    bd.Recordset!rempor = rempor.Text
    bd.Recordset!prorem = prorem.Text
    bd.Recordset!entrem = entrem.Text
    bd.Recordset!numdoc = numdoc.Text
    bd.Recordset!depnac = depnac.Text
    bd.Recordset!munnac = munnac.Text
    bd.Recordset!Painac = Painac.Text
    bd.Recordset!tipdisest = tipdisest.Text
    bd.Recordset!niveduben = nivestalc.Text
    bd.Recordset!asiactcenedu = asiactcenedu.Text
    bd.Recordset!proaso = proaso.Text
    bd.Recordset!afisegsocben = afisegsocfam.Text
    bd.Recordset!regsegsocben = regsegsocfam.Text
    bd.Recordset!calben = calbenfam.Text
    bd.Recordset!vinsecsalben = vinsecsalfam.Text
    bd.Recordset!numficsis = numficsis.Text
    bd.Recordset!punsis = punsis.Text
    bd.Recordset!loc = loc.Text
    bd.Recordset!forpagviv = forpagviv.Text
    bd.Recordset!dptoprofam = dep.Text
    bd.Recordset!munprofam = mun.Text
    bd.Recordset!paiprofam = pais.Text
    bd.Recordset!fecllebogfam = feclle.Value
    bd.Recordset!ninvivpapmam = ninvivpapmam.Text
    bd.Recordset!ninvivperpadmadotr = ninvivperpadmadotr.Text
    bd.Recordset!vivperpadmad = vivpermpadmad.Text
    bd.Recordset!edaninvivpapmad = edaninvivpapmad.Text
    bd.Recordset!cuinindurdia = cuinindurdia.Text
    bd.Recordset!graasp = graasp.Text
    bd.Recordset!nomdilform = nomdilfor.Text
    bd.Recordset!nomfundighojsir = nomfundighojsir.Text
    bd.Recordset!fecdighojsir = fecdighojsir.Value
    bd.Recordset!obs = obs.Text
    bd.Recordset!mat = 1
bd.Recordset.Update
'modifica el registro en inscripciones para asiganarle que ya esta matrriculado
ConexionBD1 matricula, "select * from inscripciones where numdoc='" & NUM_DOC & "'"
Dim modi
modi = BD1.Recordset.EditMode
BD1.Recordset!Matriculado = 1
BD1.Recordset.Update

'guardar en la tabla hermanos
ConexionBD2 matricula, "select * from hermanos"
For i = 1 To 8 'filas
    If malla.TextMatrix(i, 0) <> "" Then
        bd2.Recordset.AddNew
            bd2.Recordset!numdoc = numdoc.Text
            bd2.Recordset!tipdocher = malla.TextMatrix(i, 0)
            bd2.Recordset!numdocher = malla.TextMatrix(i, 1)
            bd2.Recordset!apeher = malla.TextMatrix(i, 2)
            bd2.Recordset!nomher = malla.TextMatrix(i, 3)
            bd2.Recordset!sexher = malla.TextMatrix(i, 4)
            bd2.Recordset!fecnacher = malla.TextMatrix(i, 5)
            bd2.Recordset!edaproher = Val(malla.TextMatrix(i, 6))
            bd2.Recordset!estcivher = malla.TextMatrix(i, 7)
            bd2.Recordset!tipdisher = malla.TextMatrix(i, 8)
            bd2.Recordset!parher = malla.TextMatrix(i, 9)
            bd2.Recordset!esther = malla.TextMatrix(i, 10)
            bd2.Recordset!asiacteduher = malla.TextMatrix(i, 11)
            bd2.Recordset!actocuher = malla.TextMatrix(i, 12)
            bd2.Recordset!posocuher = malla.TextMatrix(i, 13)
            bd2.Recordset!ingmesher = malla.TextMatrix(i, 14)
            bd2.Recordset!forperingher = malla.TextMatrix(i, 15)
            bd2.Recordset!afisalher = malla.TextMatrix(i, 16)
            bd2.Recordset!regsegsalher = malla.TextMatrix(i, 17)
            bd2.Recordset!calbenher = malla.TextMatrix(i, 18)
            bd2.Recordset!vinsecsalher = malla.TextMatrix(i, 19)
            bd2.Recordset!agrviointher = malla.TextMatrix(i, 20)
        bd2.Recordset.Update
    End If
Next i
'genera el carnet
Matriculado = True
ConexionDocu matricula, "select * from listadodeespera where numdoc ='" & numdoc.Text & "';"

certificados.nombrecarne.Text = docu.Recordset!prinom & " " & docu.Recordset!segnom
certificados.apecarne.Text = docu.Recordset!priape & " " & docu.Recordset!segape
certificados.nivelcarne.Text = graasp.Text
certificados.carne.Visible = True
certificados.Show
carnete = True

'habilitar controles
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


Private Sub JeweledButton1_Click()

End Sub


Private Sub loc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    est.SetFocus
End If
End Sub

Private Sub malla_RowColChange()
datoc.Visible = False
datof.Visible = False
datot.Visible = False

Select Case malla.col
    Case 0:
        datoc.Width = malla.ColWidth(0)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        'llenar tipo documento
        ConexionBD2 matricula, "select * from parametrizacion where tippar=6"
        datoc.Clear
        If bd2.Recordset.RecordCount > 0 Then
            bd2.Recordset.MoveFirst
            For i = 1 To bd2.Recordset.RecordCount
                datoc.AddItem bd2.Recordset!dato
                bd2.Recordset.MoveNext
            Next i
        End If
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 1:
        datot.Width = malla.ColWidth(1)
        datot.Top = malla.CellTop + malla.Top
        datot.Left = malla.CellLeft + malla.Left
        'datot.Text = ""
        datot.Locked = False
        datot.Visible = True
        datot.SetFocus
        datot.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 2:
        datot.Width = malla.ColWidth(2)
        datot.Top = malla.CellTop + malla.Top
        datot.Left = malla.CellLeft + malla.Left
        'datot.Text = ""
        datot.Locked = False
        datot.Visible = True
        datot.SetFocus
        datot.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 3:
        'texto
        datot.Width = malla.ColWidth(3)
        datot.Top = malla.CellTop + malla.Top
        datot.Left = malla.CellLeft + malla.Left
        'datot.Text = ""
        datot.Locked = False
        datot.Visible = True
        datot.SetFocus
        datot.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 4:
        'sexo
        datoc.Width = malla.ColWidth(4)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        datoc.Clear
        datoc.AddItem "F"
        datoc.AddItem "M"
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 5:
        'fecha
        datof.Width = malla.ColWidth(5)
        datof.Top = malla.CellTop + malla.Top
        datof.Left = malla.CellLeft + malla.Left
        datof.Visible = True
        datof.SetFocus
    Case 6:
        'texto
        datot.Width = malla.ColWidth(6)
        datot.Top = malla.CellTop + malla.Top
        datot.Left = malla.CellLeft + malla.Left
        'datot.Text = ""
        datot.Locked = False
        datot.Visible = True
        datot.SetFocus
        datot.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 7:
        'estado civil
        datoc.Width = malla.ColWidth(7)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        'llenar tipo documento
        ConexionBD2 matricula, "select * from parametrizacion where tippar=28"
        datoc.Clear
        If bd2.Recordset.RecordCount > 0 Then
            bd2.Recordset.MoveFirst
            For i = 1 To bd2.Recordset.RecordCount
                datoc.AddItem bd2.Recordset!dato
                bd2.Recordset.MoveNext
            Next i
        End If
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 8:
        'tipo discapacidad
        datoc.Width = malla.ColWidth(8)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        'llenar tipo documento
        ConexionBD2 matricula, "select * from parametrizacion where tippar=2"
        datoc.Clear
        If bd2.Recordset.RecordCount > 0 Then
            bd2.Recordset.MoveFirst
            For i = 1 To bd2.Recordset.RecordCount
                datoc.AddItem bd2.Recordset!dato
                bd2.Recordset.MoveNext
            Next i
        End If
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 9:
        'parentesco
        datoc.Width = malla.ColWidth(9)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        'llenar tipo documento
        ConexionBD2 matricula, "select * from parametrizacion where tippar=29"
        datoc.Clear
        If bd2.Recordset.RecordCount > 0 Then
            bd2.Recordset.MoveFirst
            For i = 1 To bd2.Recordset.RecordCount
                datoc.AddItem bd2.Recordset!dato
                bd2.Recordset.MoveNext
            Next i
        End If
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 10:
        'estudios hermanos
        datoc.Width = malla.ColWidth(10)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        'llenar tipo documento
        ConexionBD2 matricula, "select * from parametrizacion where tippar=3"
        datoc.Clear
        If bd2.Recordset.RecordCount > 0 Then
            bd2.Recordset.MoveFirst
            For i = 1 To bd2.Recordset.RecordCount
                datoc.AddItem bd2.Recordset!dato
                bd2.Recordset.MoveNext
            Next i
        End If
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 11:
        'asisste actualmente a c.e
        datoc.Width = malla.ColWidth(11)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        datoc.Clear
        datoc.AddItem "SI"
        datoc.AddItem "NO"
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 12:
        'actividad ocupacional
        datoc.Width = malla.ColWidth(12)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        'llenar tipo documento
        ConexionBD2 matricula, "select * from parametrizacion where tippar=12"
        datoc.Clear
        If bd2.Recordset.RecordCount > 0 Then
            bd2.Recordset.MoveFirst
            For i = 1 To bd2.Recordset.RecordCount
                datoc.AddItem bd2.Recordset!dato
                bd2.Recordset.MoveNext
            Next i
        End If
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 13:
        'estudios hermanos
        datoc.Width = malla.ColWidth(13)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        'llenar tipo documento
        ConexionBD2 matricula, "select * from parametrizacion where tippar=9"
        datoc.Clear
        If bd2.Recordset.RecordCount > 0 Then
            bd2.Recordset.MoveFirst
            For i = 1 To bd2.Recordset.RecordCount
                datoc.AddItem bd2.Recordset!dato
                bd2.Recordset.MoveNext
            Next i
        End If
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 14:
        'ingreso mensual hermano
        'texto
        datot.Width = malla.ColWidth(14)
        datot.Top = malla.CellTop + malla.Top
        datot.Left = malla.CellLeft + malla.Left
        'datot.Text = ""
        datot.Locked = False
        datot.Visible = True
        datot.SetFocus
        datot.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 15:
        'forma de percibir ingresos
        datoc.Width = malla.ColWidth(15)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        ConexionBD2 matricula, "select * from parametrizacion where tippar=27"
        datoc.Clear
        If bd2.Recordset.RecordCount > 0 Then
            bd2.Recordset.MoveFirst
            For i = 1 To bd2.Recordset.RecordCount
                datoc.AddItem bd2.Recordset!dato
                bd2.Recordset.MoveNext
            Next i
        End If
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 16:
        'afiliado sec. salud
        datoc.Width = malla.ColWidth(16)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        datoc.Clear
        datoc.AddItem "SI"
        datoc.AddItem "NO"
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 17:
        'reg. seguridad social
        datoc.Width = malla.ColWidth(17)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        datoc.Clear
        datoc.AddItem "EPS"
        datoc.AddItem "ARS"
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 18:
        'calidad del beneficiario
        datoc.Width = malla.ColWidth(18)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        datoc.Clear
        datoc.AddItem "Directo"
        datoc.AddItem "Beneficiario Dependiente"
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 19:
        'vinculado secretaria de salud
        datoc.Width = malla.ColWidth(19)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        datoc.Clear
        datoc.AddItem "SI"
        datoc.AddItem "NO"
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
    Case 20:
        'agresor
        datoc.Width = malla.ColWidth(20)
        datoc.Top = malla.CellTop + malla.Top
        datoc.Left = malla.CellLeft + malla.Left
        datoc.Clear
        datoc.AddItem "SI"
        datoc.AddItem "NO"
        datoc.Visible = True
        datoc.SetFocus
        datoc.Text = malla.TextMatrix(malla.RowSel, malla.ColSel)
End Select
End Sub

Private Sub malla_Scroll()
datoc.Visible = False
datof.Visible = False
datot.Visible = False
End Sub

Private Sub mandis_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    tipdisest.SetFocus
End If
End Sub

Private Sub mod_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub

Private Sub modal_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    submod.SetFocus
End If
End Sub

Private Sub modificar_Click()
If bd.Recordset.RecordCount > 0 Then
ModificadoM = True
modificar.Enabled = False
Actualizar.Enabled = True
nuevo.Enabled = False
eliminar.Enabled = False
busqueda.Enabled = False
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False
parametro.Enabled = False
'desbloquear cajas
ModificarCajas matricula
fecnac.Enabled = False

End If
End Sub

Private Sub mostrar_Click()
datoc.Visible = False
datot.Visible = False
datof.Visible = False
'muestra en la malla los datos de los hermanos del niño
Dim numero As Integer
ConexionBD2 matricula, "select * from hermanos where numdoc='" & numdoc.Text & "'"
numero = bd2.Recordset.RecordCount
If bd2.Recordset.RecordCount > 0 Then
    bd2.Recordset.MoveFirst
End If
If numero > 0 Then
    For i = 1 To numero
        
            malla.TextMatrix(i, 0) = bd2.Recordset!tipdocher
            malla.TextMatrix(i, 1) = bd2.Recordset!numdocher
            malla.TextMatrix(i, 2) = bd2.Recordset!apeher
            malla.TextMatrix(i, 3) = bd2.Recordset!nomher
            malla.TextMatrix(i, 4) = bd2.Recordset!sexher
            If IsNull(bd2.Recordset!fecnacher) = False Then
                malla.TextMatrix(i, 5) = bd2.Recordset!fecnacher
            End If
            malla.TextMatrix(i, 6) = bd2.Recordset!edaproher
            malla.TextMatrix(i, 7) = bd2.Recordset!estcivher
            malla.TextMatrix(i, 8) = bd2.Recordset!tipdisher
            malla.TextMatrix(i, 9) = bd2.Recordset!parher
            malla.TextMatrix(i, 10) = bd2.Recordset!esther
            malla.TextMatrix(i, 11) = bd2.Recordset!asiacteduher
            malla.TextMatrix(i, 12) = bd2.Recordset!actocuher
            malla.TextMatrix(i, 13) = bd2.Recordset!posocuher
            malla.TextMatrix(i, 14) = bd2.Recordset!ingmesher
            malla.TextMatrix(i, 15) = bd2.Recordset!forperingher
            malla.TextMatrix(i, 16) = bd2.Recordset!afisalher
            malla.TextMatrix(i, 17) = bd2.Recordset!regsegsalher
            malla.TextMatrix(i, 18) = bd2.Recordset!calbenher
            malla.TextMatrix(i, 19) = bd2.Recordset!vinsecsalher
            malla.TextMatrix(i, 20) = bd2.Recordset!agrviointher
            If bd2.Recordset.RecordCount > 0 Then
                bd2.Recordset.MoveNext
            End If
    Next i
End If
End Sub

Private Sub mun_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    pais.SetFocus
End If
End Sub

Private Sub munnac_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    Painac.SetFocus
End If
End Sub

Private Sub munprofam_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    fecllebogfam.SetFocus
End If
End Sub
Private Sub ninvivperpapmad_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub

Private Sub nieestalcfam_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    act.SetFocus
End If
End Sub

Private Sub ninvivpapmam_Click()
If ninvivpapmam.Text = "SI" Then
    ninvivperpadmadotr.Enabled = False
    vivpermpadmad.Enabled = False
    edaninvivpapmad.Enabled = False
    cuinindurdia.SetFocus
ElseIf ninvivpapmam.Text = "NO" Then
    ninvivperpadmadotr.Enabled = True
    vivpermpadmad.Enabled = True
    edaninvivpapmad.Enabled = True
    ninvivperpadmadotr.SetFocus
End If
End Sub

Private Sub ninvivpapmam_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 And ninvivpapmam.Text = "NO" Then
    ninvivperpadmadotr.SetFocus
ElseIf KeyAscii = 13 And ninvivpapmam.Text = "SI" Then
    cuinindurdia.SetFocus
End If
End Sub

Private Sub ninvivperpadmad_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub

Private Sub ninvivperpadmadotr_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    vivpermpadmad.SetFocus
End If
End Sub

Private Sub nivestalc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    asiactcenedu.SetFocus
End If
End Sub

Private Sub nivestfam_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub

Private Sub nomdilfor_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    nomfundighojsir.SetFocus
End If
End Sub

Private Sub nomfundighojsir_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    fecdighojsir.SetFocus
End If
End Sub

Private Sub nuevo_Click()
NuevoRegM = True
'genera el autonumerico del formulario

Matriculado = False
'deshabilitar controles
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False
nuevo.Enabled = False
guardar.Enabled = True
modificar.Enabled = False
eliminar.Enabled = False
parametro.Enabled = False

'deshabilitar cajas y combos
col.Locked = False
uniope.Locked = False
modal.Locked = False
submod.Locked = False
'numfor.Locked = False
persolser.Locked = False
rempor.Locked = False
prorem.Locked = False
depnac.Locked = False
munnac.Locked = False
Painac.Locked = False
mandis.Locked = False
proaso.Locked = False
numficsis.Locked = False
punsis.Locked = False
loc.Locked = False
est.Locked = False
forpagviv.Locked = False
vinsecsalfam.Locked = False
entrem.Locked = False
'sexfam.Locked = False
'edafam.Locked = False
'estciv.Locked = False
'apredufam.Locked = False
'posocu.Locked = False
regsegsocfam.Locked = False
calbenfam.Locked = False
'limpia las cajas
cajasM matricula
'genera el numero de matricula
ConexionBD1 matricula, "select max(numfor)as nm from matricula"

If IsNull(BD1.Recordset!nm) = False Then
    numfor = Val(BD1.Recordset!nm) + 1
Else
    numfor = 1
End If

SSTab1.Tab = 0
busqueda.Enabled = False
col.SetFocus
'genera el usuario digito matricula
nomdilfor.Text = usuario
nomdilfor.Locked = True
Painac.Text = "Colombia"
pais.Text = "Colombia"
End Sub

Private Sub numdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    If numdoc.Text <> "" Then
        ConexionBD1 matricula, "select numdoc from matricula where numdoc='" & numdoc.Text & "'"
        If BD1.Recordset.RecordCount > 0 Then
            MsgBox "Estudiante ya matriculado!", vbInformation, "Matriculas"
            Exit Sub
        End If
        ConexionBD1 matricula, "select numdoc from inscripciones where numdoc='" & numdoc.Text & "'"
        If BD1.Recordset.RecordCount > 0 Then
            ConexionBD1 matricula, "select * from listadodeespera where numdoc='" & numdoc.Text & "'"
            'cargamos datos y bloqueamos cajas
            If BD1.Recordset.RecordCount > 0 Then
                tipdocnin.Text = BD1.Recordset!tipdoc
                tipdocnin.Locked = True
                nomnin.Text = BD1.Recordset!prinom & " " & BD1.Recordset!segnom
                nomnin.Locked = True
                apenin.Text = BD1.Recordset!priape & " " & BD1.Recordset!segape
                apenin.Locked = True
                sexo.Text = BD1.Recordset!sex
                sexo.Locked = True
                If IsNull(BD1.Recordset!fecnac) = False Then
                    fecnac.Value = BD1.Recordset!fecnac
                End If
                fecnac.Enabled = False
                edad.Text = BD1.Recordset!eda
                edad.Locked = True
                'genera el grado del niño
                If BD1.Recordset!meses >= 3 And BD1.Recordset!meses <= 12 Then
                    graasp.Text = "Salacuna"
                ElseIf BD1.Recordset!meses >= 13 And BD1.Recordset!meses <= 24 Then
                    graasp.Text = "Caminadores"
                ElseIf BD1.Recordset!meses >= 25 And BD1.Recordset!meses <= 30 Then
                    graasp.Text = "Párvulos"
                ElseIf BD1.Recordset!meses >= 31 And BD1.Recordset!meses <= 36 Then
                    graasp.Text = "Prekinder"
                ElseIf BD1.Recordset!meses >= 37 And BD1.Recordset!meses <= 48 Then
                    graasp.Text = "Prekinder"
                ElseIf BD1.Recordset!meses >= 49 And BD1.Recordset!meses <= 60 Then
                    graasp.Text = "Kinder"
                End If
                parjeffam.Text = BD1.Recordset!parfam
                parjeffam.Locked = True
                dir.Text = BD1.Recordset!dir
                dir.Locked = True
                bar.Text = BD1.Recordset!bar
                bar.Locked = True
                tel.Text = BD1.Recordset!tel
                tel.Locked = True
                'conectamos bd para cargar datos de inscripciones
                ConexionBD1 matricula, "select * from inscripciones where numdoc='" & numdoc.Text & "'"
                tipviv.Text = BD1.Recordset!tipviv
                tipviv.Locked = True
                conviv.Text = BD1.Recordset!conviv
                conviv.Locked = True
                tenviv.Text = BD1.Recordset!tenviv
                tenviv.Locked = True
                'llenamos la malla con los datos de los padres
                If BD1.Recordset!nompad <> Null Or BD1.Recordset!nompad <> "" Then
                    malla.TextMatrix(1, 6) = BD1.Recordset!edapad
                    malla.TextMatrix(1, 14) = BD1.Recordset!ingmenpad
                    malla.TextMatrix(1, 10) = BD1.Recordset!nivacapad
                End If
                If BD1.Recordset!nommad <> Null Or BD1.Recordset!nommad <> "" Then
                    malla.TextMatrix(2, 6) = BD1.Recordset!edamad
                    malla.TextMatrix(2, 14) = BD1.Recordset!ingmenmad
                    malla.TextMatrix(2, 10) = BD1.Recordset!nivacamad
                End If
                If BD1.Recordset!nompad = Null Or BD1.Recordset!nompad = "" Then 'And bd1.Recordset!nommad <> Null Or bd1.Recordset!nommad <> "" Then
                    malla.TextMatrix(1, 6) = BD1.Recordset!edamad
                    malla.TextMatrix(1, 14) = BD1.Recordset!ingmenmad
                    malla.TextMatrix(1, 10) = BD1.Recordset!nivacamad
                End If
                
            Else
                MsgBox "El documento no ha sido encontrado!" & vbCrLf & "El niño no se puede matricular!", vbInformation, "Matriculas"
                Exit Sub
            End If
        End If
    End If
    
End If
End Sub
Function avanzar()
 If tecla = 13 Then
  SendKeys "{tab}"
  KeyAscii = 0: tecla = 0
 End If
End Function
Private Sub numficsis_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    punsis.SetFocus
End If
End Sub

Private Sub numfor_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)

End Sub

Private Sub obs_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
End Sub

Private Sub opciones_Click()
PopupMenu optiones, 2
End Sub

Private Sub Painac_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    mandis.SetFocus
End If
End Sub

Private Sub paiprofam_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    continuar3_Click
End If
End Sub

Private Sub pais_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    feclle.SetFocus
End If
End Sub

Private Sub parametro_Click()
Para = True
ingresos.Show
End Sub

Private Sub parjeffam_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    nivestalc.SetFocus
End If
End Sub

Private Sub Persolser_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    rempor.SetFocus
End If
End Sub

Private Sub posocu_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    forpering.SetFocus
End If
End Sub

Private Sub primero_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveFirst
    mostrarcampos
    nd = numdoc.Text
End If
numdoc.SetFocus
End Sub

Private Sub proaso_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    continuar1_Click
End If
End Sub


Private Sub prorem_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    entrem.SetFocus
End If
End Sub
Private Sub regsegsoc_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub

Private Sub punsis_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    continuar2_Click
End If
End Sub

Private Sub regsegsocfam_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    calbenfam.SetFocus
End If
End Sub

Private Sub regsegsociben_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    calben.SetFocus
End If
End Sub

Private Sub rempor_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    prorem.SetFocus
End If
End Sub
Private Sub sex_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Private Sub sexfam_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    estciv.SetFocus
End If
End Sub

Private Sub sexo_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)

End Sub

Private Sub siguiente_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveNext
    If bd.Recordset.EOF Then
        bd.Recordset.MoveLast
    End If
    mostrarcampos
    nd = numdoc.Text
End If
numdoc.SetFocus
End Sub


Private Sub Form_Load()
LlenarCombos
menu.estado.Visible = False
'crear la conexion al bd
ConexionBD matricula, "select * from matricula"
If bd.Recordset.RecordCount > 0 Then
    mostrarcampos
End If
'deshabilitar cajas y combos
Deshabilitar matricula
End Sub

Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 5: datoc.Visible = False
            datot.Visible = False
            'carga datos de los padres
            
End Select
End Sub

Private Sub submod_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    persolser.SetFocus
End If
End Sub

Private Sub tipdisest_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    parjeffam.SetFocus
End If
End Sub

Private Sub tipdisfam_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    parjeffam.SetFocus
End If
End Sub
Private Sub tipdocfam_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub
Private Sub tipdoc_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    docide.SetFocus
End If
End Sub

Private Sub ultimo_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveLast
    mostrarcampos
    nd = numdoc.Text
End If
numdoc.SetFocus
End Sub

Private Sub uniope_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    modal.SetFocus
End If
End Sub

Private Sub vinsecsalben_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    afisegsocben.SetFocus
End If
End Sub

Private Sub vinsecsalfam_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    numficsis.SetFocus
End If
End Sub
Private Sub vivpermpadmad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    edaninvivpapmad.SetFocus
End If
End Sub

Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "matriculas.htm"
End Sub
