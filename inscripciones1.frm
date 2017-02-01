VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form inscripciones 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inscripciones"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "inscripciones1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10860
   Begin Jardin.xphelp xphelp1 
      Height          =   315
      Left            =   10200
      Top             =   3960
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin MSAdodcLib.Adodc bd 
      Height          =   330
      Left            =   7680
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc BD1 
      Height          =   330
      Left            =   7680
      Top             =   5640
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos del Niño"
      TabPicture(0)   =   "inscripciones1.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "xpgroupbox1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fr"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "continuar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "docu"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Datos de la Familia"
      TabPicture(1)   =   "inscripciones1.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "xpgroupbox4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Datos de las Condiciones de la Vivienda"
      TabPicture(2)   =   "inscripciones1.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image2"
      Tab(2).Control(1)=   "xpgroupbox8"
      Tab(2).ControlCount=   2
      Begin MSAdodcLib.Adodc docu 
         Height          =   375
         Left            =   5520
         Top             =   720
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
      Begin JeweledBut.JeweledButton continuar 
         Height          =   375
         Left            =   7560
         TabIndex        =   1
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
         MICON           =   "inscripciones1.frx":019E
         BC              =   8438015
         FC              =   0
      End
      Begin Jardin.xpgroupbox xpgroupbox8 
         Height          =   3255
         Left            =   -72000
         TabIndex        =   2
         Top             =   1200
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5741
         Caption         =   "Vivienda"
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
         Begin VB.ComboBox tipviv 
            Height          =   315
            ItemData        =   "inscripciones1.frx":01BA
            Left            =   2520
            List            =   "inscripciones1.frx":01CA
            TabIndex        =   7
            Top             =   840
            Width           =   1815
         End
         Begin VB.ComboBox tenviv 
            Height          =   315
            ItemData        =   "inscripciones1.frx":01F2
            Left            =   2520
            List            =   "inscripciones1.frx":01FF
            TabIndex        =   6
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox conviv 
            Height          =   315
            ItemData        =   "inscripciones1.frx":0228
            Left            =   2520
            List            =   "inscripciones1.frx":0238
            TabIndex        =   5
            Top             =   1440
            Width           =   2295
         End
         Begin VB.ComboBox estviv 
            Height          =   315
            ItemData        =   "inscripciones1.frx":0269
            Left            =   2520
            List            =   "inscripciones1.frx":0273
            TabIndex        =   4
            Top             =   2040
            Width           =   1815
         End
         Begin VB.ComboBox serpub 
            Height          =   315
            ItemData        =   "inscripciones1.frx":028D
            Left            =   2520
            List            =   "inscripciones1.frx":02A6
            TabIndex        =   3
            Top             =   2640
            Width           =   2535
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tenencia de la Vivienda"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   2040
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condiciones de la Vivenda"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   2265
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de la Vivienda"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   1635
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado de la Vivienda"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   2160
            Width           =   1845
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Servicios Públicos"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   2760
            Width           =   1530
         End
      End
      Begin Jardin.xpgroupbox fr 
         Height          =   3135
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5530
         Caption         =   "Niño"
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
         Begin VB.TextBox numher 
            Height          =   285
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   19
            Top             =   2160
            Width           =   735
         End
         Begin VB.OptionButton xpradiobutton1 
            Caption         =   "SI"
            Height          =   255
            Left            =   2160
            TabIndex        =   18
            Top             =   960
            Width           =   855
         End
         Begin VB.OptionButton xpradiobutton2 
            Caption         =   "NO"
            Height          =   255
            Left            =   3480
            TabIndex        =   17
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox prealgenf 
            Height          =   645
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   1320
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox lugocufam 
            Height          =   285
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   15
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox lugnac 
            Height          =   285
            Left            =   2160
            MaxLength       =   20
            TabIndex        =   14
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label87 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar de Nacimiento"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   1755
         End
         Begin VB.Label Label92 
            BackStyle       =   0  'Transparent
            Caption         =   "El niño(a) presenta alguna Enfermedad"
            ForeColor       =   &H80000007&
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label84 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar que Ocupa en la Familia"
            ForeColor       =   &H80000007&
            Height          =   495
            Left            =   120
            TabIndex        =   21
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label Label93 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Hermanos"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   2160
            Width           =   1455
         End
      End
      Begin Jardin.xpgroupbox xpgroupbox4 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   24
         Top             =   720
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7646
         Caption         =   "Familia"
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
         Begin VB.ComboBox ninviv 
            Height          =   315
            ItemData        =   "inscripciones1.frx":0317
            Left            =   1920
            List            =   "inscripciones1.frx":032A
            TabIndex        =   25
            Top             =   240
            Width           =   1815
         End
         Begin Jardin.xpgroupbox xpgroupbox6 
            Height          =   3255
            Left            =   4320
            TabIndex        =   26
            Top             =   480
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   5741
            Caption         =   "Madre"
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
            Begin VB.TextBox nommad 
               Height          =   285
               Left            =   1800
               MaxLength       =   30
               TabIndex        =   34
               Top             =   240
               Width           =   2175
            End
            Begin VB.TextBox ocumad 
               Height          =   285
               Left            =   1800
               MaxLength       =   15
               TabIndex        =   33
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox nomempmad 
               Height          =   285
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   32
               Top             =   1680
               Width           =   1815
            End
            Begin VB.TextBox edamad 
               Height          =   285
               Left            =   1800
               MaxLength       =   2
               TabIndex        =   31
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox ingmenmad 
               Height          =   285
               Left            =   1800
               MaxLength       =   7
               TabIndex        =   30
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox telempmad 
               Height          =   285
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   29
               Top             =   2040
               Width           =   1815
            End
            Begin VB.ComboBox nivacamad 
               Height          =   315
               ItemData        =   "inscripciones1.frx":036C
               Left            =   1800
               List            =   "inscripciones1.frx":036E
               TabIndex        =   28
               Top             =   2400
               Width           =   1815
            End
            Begin VB.TextBox otringmad 
               Height          =   285
               Left            =   1800
               MaxLength       =   7
               TabIndex        =   27
               Top             =   2760
               Width           =   1815
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "$"
               ForeColor       =   &H80000007&
               Height          =   195
               Index           =   3
               Left            =   1650
               TabIndex        =   48
               Top             =   2850
               Width           =   105
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "$"
               ForeColor       =   &H80000007&
               Height          =   195
               Index           =   2
               Left            =   1650
               TabIndex        =   47
               Top             =   1400
               Width           =   105
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nivel Académico"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   46
               Top             =   2520
               Width           =   1410
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Telefono Empresa"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   45
               Top             =   2160
               Width           =   1545
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre Empresa"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   44
               Top             =   1800
               Width           =   1485
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Edad"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   43
               Top             =   720
               Width           =   420
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   42
               Top             =   360
               Width           =   675
            End
            Begin VB.Label Label97 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ocupación"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   41
               Top             =   1080
               Width           =   885
            End
            Begin VB.Label Label60 
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre de la Empresa Madre"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000007&
               Height          =   375
               Left            =   -3360
               TabIndex        =   40
               Top             =   2880
               Width           =   1815
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Otros Ingresos"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   39
               Top             =   2880
               Width           =   1275
            End
            Begin VB.Label Label98 
               BackStyle       =   0  'Transparent
               Caption         =   "Edad Madre"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000007&
               Height          =   255
               Left            =   -3360
               TabIndex        =   38
               Top             =   2280
               Width           =   1095
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ingreso Mensual"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   37
               Top             =   1440
               Width           =   1410
            End
            Begin VB.Label Label99 
               BackStyle       =   0  'Transparent
               Caption         =   "Teléfono de la empresa Madre"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000007&
               Height          =   375
               Left            =   -3360
               TabIndex        =   36
               Top             =   3480
               Width           =   2295
            End
            Begin VB.Label Label63 
               BackStyle       =   0  'Transparent
               Caption         =   "Nivel Académico Madre"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000007&
               Height          =   255
               Left            =   -3360
               TabIndex        =   35
               Top             =   3960
               Width           =   2175
            End
         End
         Begin Jardin.xpgroupbox xpgroupbox5 
            Height          =   3255
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   5741
            Caption         =   "Padre"
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
            Begin VB.TextBox edapad 
               Height          =   285
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   57
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox telemppad 
               Height          =   285
               Left            =   1920
               MaxLength       =   10
               TabIndex        =   56
               Top             =   2040
               Width           =   1815
            End
            Begin VB.TextBox nomemppad 
               Height          =   285
               Left            =   1920
               MaxLength       =   25
               TabIndex        =   55
               Top             =   1680
               Width           =   1815
            End
            Begin VB.TextBox nompad 
               Height          =   285
               Left            =   1920
               MaxLength       =   30
               TabIndex        =   54
               Top             =   240
               Width           =   2055
            End
            Begin VB.TextBox ocupad 
               Height          =   285
               Left            =   1920
               MaxLength       =   15
               TabIndex        =   53
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox ingmenpad 
               Height          =   285
               Left            =   1920
               MaxLength       =   7
               TabIndex        =   52
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox otringpad 
               Height          =   285
               Left            =   1920
               MaxLength       =   7
               TabIndex        =   51
               Top             =   2800
               Width           =   1815
            End
            Begin VB.ComboBox nivacapad 
               Height          =   315
               ItemData        =   "inscripciones1.frx":0370
               Left            =   1920
               List            =   "inscripciones1.frx":0372
               TabIndex        =   50
               Top             =   2400
               Width           =   1815
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "$"
               ForeColor       =   &H80000007&
               Height          =   195
               Index           =   1
               Left            =   1680
               TabIndex        =   67
               Top             =   2880
               Width           =   105
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "$"
               ForeColor       =   &H80000007&
               Height          =   195
               Index           =   0
               Left            =   1680
               TabIndex        =   66
               Top             =   1420
               Width           =   105
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               BackStyle       =   0  'Transparent
               Caption         =   "Edad"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   65
               Top             =   720
               Width           =   420
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ingreso Mensual"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   64
               Top             =   1440
               Width           =   1410
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ocupación"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   63
               Top             =   1080
               Width           =   885
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   62
               Top             =   360
               Width           =   675
            End
            Begin VB.Label Label61 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Otros Ingresos"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   61
               Top             =   2880
               Width           =   1275
            End
            Begin VB.Label Label94 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Telèfono Empresa"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   60
               Top             =   2160
               Width           =   1545
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre Empresa"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   59
               Top             =   1800
               Width           =   1485
            End
            Begin VB.Label Label95 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nivel Académico"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   120
               TabIndex        =   58
               Top             =   2520
               Width           =   1410
            End
         End
         Begin JeweledBut.JeweledButton continue 
            Height          =   375
            Left            =   7320
            TabIndex        =   68
            Top             =   3840
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
            MICON           =   "inscripciones1.frx":0374
            BC              =   8438015
            FC              =   0
         End
         Begin VB.Label Label83 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "El niño(a)vive con"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   1545
         End
      End
      Begin Jardin.xpgroupbox xpgroupbox1 
         Height          =   1575
         Left            =   120
         TabIndex        =   70
         Top             =   480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2778
         Caption         =   "Consultar Cupo"
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
         Begin VB.ComboBox numdoc 
            Height          =   315
            Left            =   1560
            TabIndex        =   71
            Top             =   360
            Width           =   2535
         End
         Begin JeweledBut.JeweledButton buscardoc 
            Height          =   495
            Left            =   2520
            TabIndex        =   72
            Top             =   960
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   873
            TX              =   "Consultar Listado en Espera"
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
            MICON           =   "inscripciones1.frx":0390
            BC              =   12632256
            FC              =   0
            Picture         =   "inscripciones1.frx":04FE
         End
         Begin VB.Label numins 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1560
            TabIndex        =   75
            Top             =   1200
            Width           =   60
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "No. Inscripción:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   1200
            Width           =   1350
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "No Documento"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   480
            Width           =   1260
         End
      End
      Begin VB.Image Image2 
         Height          =   1965
         Left            =   -74760
         Picture         =   "inscripciones1.frx":2208
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2370
      End
      Begin VB.Image Image1 
         Height          =   2565
         Left            =   6600
         Picture         =   "inscripciones1.frx":6EF7
         Top             =   600
         Width           =   1905
      End
   End
   Begin Jardin.xpgroupbox xpgroupbox2 
      Height          =   3615
      Left            =   9000
      TabIndex        =   76
      Top             =   120
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
         TabIndex        =   77
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
         MICON           =   "inscripciones1.frx":A119
         BC              =   8438015
         FC              =   0
         Picture         =   "inscripciones1.frx":A287
      End
      Begin JeweledBut.JeweledButton busqueda 
         Height          =   375
         Left            =   120
         TabIndex        =   78
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
         MICON           =   "inscripciones1.frx":CD91
         BC              =   8438015
         FC              =   0
         Picture         =   "inscripciones1.frx":CEFF
      End
      Begin JeweledBut.JeweledButton eliminar 
         Height          =   375
         Left            =   120
         TabIndex        =   79
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
         MICON           =   "inscripciones1.frx":D059
         BC              =   8438015
         FC              =   0
         Picture         =   "inscripciones1.frx":D1C7
      End
      Begin JeweledBut.JeweledButton guardar 
         Height          =   375
         Left            =   120
         TabIndex        =   80
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
         MICON           =   "inscripciones1.frx":D761
         BC              =   8438015
         FC              =   0
         Picture         =   "inscripciones1.frx":D8CF
      End
      Begin JeweledBut.JeweledButton Actualizar 
         Height          =   375
         Left            =   120
         TabIndex        =   81
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
         MICON           =   "inscripciones1.frx":134FD
         BC              =   8438015
         FC              =   0
         Picture         =   "inscripciones1.frx":1366B
      End
      Begin JeweledBut.JeweledButton modificar 
         Height          =   375
         Left            =   120
         TabIndex        =   82
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
         MICON           =   "inscripciones1.frx":13C05
         BC              =   8438015
         FC              =   0
         Picture         =   "inscripciones1.frx":13D73
      End
      Begin JeweledBut.JeweledButton parametro 
         Height          =   375
         Left            =   120
         TabIndex        =   83
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
         MICON           =   "inscripciones1.frx":13ECD
         BC              =   8438015
         FC              =   0
      End
   End
   Begin JeweledBut.JeweledButton salir 
      Height          =   375
      Left            =   9120
      TabIndex        =   84
      Top             =   5880
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
      MICON           =   "inscripciones1.frx":1403B
      BC              =   8438015
      FC              =   0
      Picture         =   "inscripciones1.frx":141A9
   End
   Begin Jardin.xpgroupbox xpgroupbox3 
      Height          =   735
      Left            =   120
      TabIndex        =   85
      Top             =   5640
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1296
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
         TabIndex        =   86
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
         MICON           =   "inscripciones1.frx":14303
         BC              =   8438015
         FC              =   0
         Picture         =   "inscripciones1.frx":14471
      End
      Begin JeweledBut.JeweledButton siguiente 
         Height          =   375
         Left            =   3480
         TabIndex        =   87
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
         MICON           =   "inscripciones1.frx":145CB
         BC              =   8438015
         FC              =   0
         Picture         =   "inscripciones1.frx":14739
      End
      Begin JeweledBut.JeweledButton ultimo 
         Height          =   375
         Left            =   5160
         TabIndex        =   88
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
         MICON           =   "inscripciones1.frx":14893
         BC              =   8438015
         FC              =   0
         Picture         =   "inscripciones1.frx":14A01
      End
      Begin JeweledBut.JeweledButton anterior 
         Height          =   375
         Left            =   1800
         TabIndex        =   89
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
         MICON           =   "inscripciones1.frx":14B5B
         BC              =   8438015
         FC              =   0
         Picture         =   "inscripciones1.frx":14CC9
      End
   End
   Begin JeweledBut.JeweledButton cancelar 
      Height          =   375
      Left            =   9120
      TabIndex        =   90
      Top             =   5400
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
      MICON           =   "inscripciones1.frx":14E23
      BC              =   8438015
      FC              =   0
      Picture         =   "inscripciones1.frx":14F91
   End
   Begin VB.Label numreg 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   9120
      TabIndex        =   92
      Top             =   4920
      Width           =   60
   End
   Begin VB.Label total 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   9120
      TabIndex        =   91
      Top             =   4560
      Width           =   60
   End
End
Attribute VB_Name = "inscripciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dato As String
Dim i As Integer

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
'bloquear cajas de texto y combos
numdoc.Locked = True
lugnac.Locked = True
prealgenf.Locked = True
lugocufam.Locked = True
nompad.Locked = True
ocupad.Locked = True
'ingmen.Locked = True
nomemppad.Locked = True
telemppad.Locked = True
nommad.Locked = True
ocumad.Locked = True
nomempmad.Locked = True
ingmenmad.Locked = True
ingmenpad.Locked = True
edamad.Locked = True
edamad.Locked = True
telempmad.Locked = True
otringmad.Locked = True
otringpad.Locked = True
numdoc.Locked = True
numher.Locked = True
'BD.Recordset.Update
If ModificadoI = True Then
    'modifica el registro en listado de espera para asiganarle que ya no esta inscrito
    MODIFICARR
    ModificadoI = False
End If
Deshabilitari inscripciones
End Sub
Sub MODIFICARR()
'modifica el REGISTRO EN INSCRIPCIONES
'ConexionBD1 inscripciones, "select * from inscripciones where numdoc='" & numdoc.Text & "'"
Dim modii
modii = bd.Recordset.EditMode
bd.Recordset!numdoc = numdoc.Text
bd.Recordset!numins = numins.Caption
bd.Recordset!lugnac = lugnac.Text
If xpradiobutton2.Value = True Then
    bd.Recordset!prealgenf = "NO"
ElseIf xpradiobutton2.Value = False Then
    bd.Recordset!prealgenf = "NO"
ElseIf xpradiobutton1.Value = True Then
    bd.Recordset!prealgenf = prealgenf.Text
End If
bd.Recordset!numher = Val(numher.Text)
bd.Recordset!lugocufam = lugocufam.Text
bd.Recordset!ninviv = ninviv.Text
bd.Recordset!nompad = nompad.Text
bd.Recordset!ocupad = ocupad.Text
bd.Recordset!ingmenpad = Val(ingmenpad.Text)
bd.Recordset!edapad = Val(edapad.Text)
bd.Recordset!nomemppad = nomemppad.Text
bd.Recordset!telemppad = Val(telemppad.Text)
bd.Recordset!nivacapad = nivacapad.Text
bd.Recordset!otringpad = otringpad.Text
bd.Recordset!nommad = nommad.Text
bd.Recordset!ocumad = ocumad.Text
bd.Recordset!nomempmad = nomempmad.Text
bd.Recordset!ingmenmad = Val(ingmenmad.Text)
bd.Recordset!edamad = edamad.Text
bd.Recordset!telempmad = telempmad.Text
bd.Recordset!nivacamad = nivacamad.Text
bd.Recordset!otringmad = Val(otringmad.Text)
bd.Recordset!tenviv = tenviv.Text
bd.Recordset!tipviv = tipviv.Text
bd.Recordset!conviv = conviv.Text
bd.Recordset!estviv = estviv.Text
bd.Recordset!serpub = serpub.Text
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
numdoc.SetFocus
End Sub

Private Sub bd_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub bd1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub buscardoc_Click()
'CONSULTA PARA VER SI YA ESTA UN NUMDOC INSCRITO
ConexionBD1 inscripciones, "select * from inscripciones where numdoc='" & numdoc.Text & "'"
If BD1.Recordset.RecordCount > 0 Then
    MsgBox "El niño con documento: " & numdoc.Text & " ya está inscrito!", vbInformation, "Incripciones"
    Exit Sub
End If

'consulta si el documento coincide con uno de la lista
'de espera

Dim Consulta_SQL As String
Dim i As Integer, enc As Integer
enc = 0
If numdoc.Text <> "" Then
dato = numdoc.Text
Consulta_SQL = "SELECT * FROM listadodeespera " + "WHERE numdoc = '" + dato + "';"
ConexionBD1 inscripciones, Consulta_SQL
  If BD1.Recordset.RecordCount > 0 Then
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = True
    fr.Visible = True
    lugnac.SetFocus
    primero.Enabled = False
    anterior.Enabled = False
    siguiente.Enabled = False
    ultimo.Enabled = False
    nuevo.Enabled = False
    guardar.Enabled = True
    modificar.Enabled = False
  Else
    MsgBox "No se encontraron coincidencias!", vbInformation, "Inscripciones"
    numdoc.Text = ""
    numdoc.SetFocus
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    fr.Visible = False
    primero.Enabled = True
    anterior.Enabled = True
    siguiente.Enabled = True
    ultimo.Enabled = True
    nuevo.Enabled = True
    guardar.Enabled = False
    modificar.Enabled = True
    Exit Sub
  End If
ConsultaF = True
listespera.Tag = numdoc.Text
listespera.Show
End If
End Sub

Private Sub busqueda_Click()
MB.Formulario = Me.Name
MB.Descripcion = "Inscripciones"
elBuscador.Show
End Sub

Private Sub cancelar_Click()
If bd.Recordset.RecordCount > 0 Then
mostrarcampos
End If
If NuevoRegI = True Then
    nuevo.Enabled = True
    modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    anterior.Enabled = True
    siguiente.Enabled = True
    ultimo.Enabled = True
    guardar.Enabled = False
    busqueda.Enabled = True
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = True
    NuevoRegI = False
ElseIf ModificadoI = True Then
    nuevo.Enabled = True
    modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    ultimo.Enabled = True
    siguiente.Enabled = True
    anterior.Enabled = True
    Actualizar.Enabled = False
    busqueda.Enabled = True
    ModificadoI = False
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = True
End If
parametro.Enabled = True
End Sub

Private Sub continuar_Click()
If SSTab1.TabVisible(1) = True Then
    SSTab1.Tab = 1
    ninviv.SetFocus
Else
    MsgBox "Ingrese un número de documento para poder continuar!", vbInformation, "Inscripciones"
    numdoc.SetFocus
End If
End Sub

Private Sub continue_Click()
SSTab1.Tab = 2
tenviv.SetFocus
End Sub

Private Sub conviv_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    estviv.SetFocus
End If
End Sub
Private Sub edamad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    ocumad.SetFocus
End If
End Sub
Private Sub edapad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    ocupad.SetFocus
End If
End Sub
Private Sub eliminar_Click()
If bd.Recordset.RecordCount > 0 Then
If MsgBox("Está seguro de querer eliminar el registro?", vbYesNo + vbQuestion, "Eliminar Registro") = vbYes Then
    'modifica el registro en listado de espera para asiganarle que ya no esta inscrito
    ConexionBD1 inscripciones, "select * from listadodeespera where numdoc='" & bd.Recordset!numdoc & "'"
    Dim modi
    modi = BD1.Recordset.EditMode
    BD1.Recordset!inscrito = 0
    BD1.Recordset.Update
    bd.Recordset.Delete
   If bd.Recordset.RecordCount > 0 Then
    bd.Recordset.MoveFirst
    bd.Refresh
    total = bd.Recordset.RecordCount & " inscritos."
    mostrarcampos
    Else
        Unload Me
        inscripciones.Show
   End If
End If
End If
End Sub

Private Sub estviv_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    serpub.SetFocus
End If
End Sub

Private Sub LlenarCombos()
menu.estado.Panels(4).Text = "Cargando..."
'llenar nivel academico padre
ConexionBD1 inscripciones, "select * from parametrizacion where tippar=3" & " order by dato;"
nivacapad.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        nivacapad.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar nivel academico madre
ConexionBD1 inscripciones, "select * from parametrizacion where tippar=3" & " order by dato;"
nivacamad.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        nivacamad.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar tenencia de la vivienda
ConexionBD1 inscripciones, "select * from parametrizacion where tippar=13" & " order by dato;"
tenviv.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        tenviv.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar tipo de la vivienda
ConexionBD1 inscripciones, "select * from parametrizacion where tippar=7" & " order by dato;"
tipviv.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        tipviv.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar condiciones de la vivienda
ConexionBD1 inscripciones, "select * from parametrizacion where tippar=20" & " order by dato;"
conviv.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        conviv.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar estado de la vivienda
ConexionBD1 inscripciones, "select * from parametrizacion where tippar=5" & " order by dato;"
estviv.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        estviv.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar servicios publicos
ConexionBD1 inscripciones, "select * from parametrizacion where tippar=8" & " order by dato;"
serpub.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        serpub.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
End Sub
Private Sub Form_Activate()
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height
'FormularioActivo = True
buscardoc.Enabled = False
If Para = True Then
    LlenarCombos
    Para = False
    If bd.Recordset.RecordCount > 0 Then
        mostrarcampos
    End If
End If
'menu.Enabled = True
'Deshabilitari inscripciones
End Sub


Private Sub Form_Initialize()
InitCommonControls
DoEvents
End Sub

Private Sub Form_Load()
'ConstruirMenus
LlenarCombos

menu.estado.Panels(4).Text = "Control de Inscripciones"
'bloquear cajas de texto y combos
numdoc.Locked = True
lugnac.Locked = True
prealgenf.Locked = True
lugocufam.Locked = True
nompad.Locked = True
ocupad.Locked = True
'ingmen.Locked = True
nomemppad.Locked = True
telemppad.Locked = True
nommad.Locked = True
ocumad.Locked = True
nomempmad.Locked = True
ingmenmad.Locked = True
ingmenpad.Locked = True
edamad.Locked = True
edamad.Locked = True
telempmad.Locked = True
otringmad.Locked = True
otringpad.Locked = True
numdoc.Locked = True
numher.Locked = True

On Error Resume Next
ConexionBD inscripciones, "select * from inscripciones where matriculado=0"
total = bd.Recordset.RecordCount & " inscritos."
bd.Recordset.MoveFirst

If bd.Recordset.RecordCount > 0 Then
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = True
    mostrarcampos
Else
    MsgBox "No existe ningún registro en inscripciones!", vbInformation, "Inscripciones"
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    Exit Sub
End If

End Sub
Function mostrarcampos()
numreg = bd.Recordset.AbsolutePosition & " registro."
'activarcajas
'mostrar listado de documentos en listado de espera
numdoc.Clear
ConexionDocu inscripciones, "select * from listadodeespera where inscrito=0 order by numdoc"
If docu.Recordset.RecordCount > 0 Then
    docu.Recordset.MoveFirst
    For i = 1 To docu.Recordset.RecordCount
        numdoc.AddItem Trim$(docu.Recordset!numdoc)
        docu.Recordset.MoveNext
    Next i
End If
numdoc.Text = bd.Recordset!numdoc
lugnac = bd.Recordset!lugnac

If bd.Recordset!prealgenf = "NO" Then
    xpradiobutton2.Value = True
    xpradiobutton1.Value = False
    prealgenf.Visible = False
ElseIf bd.Recordset!prealgenf <> "SI" Then
    prealgenf = bd.Recordset!prealgenf
    prealgenf.Visible = True
    xpradiobutton1.Value = True
    xpradiobutton2.Value = False
End If
If IsNull(bd.Recordset!numins) = False Then
    numins = bd.Recordset!numins
End If
numher = bd.Recordset!numher
lugocufam = bd.Recordset!lugocufam
ninviv = bd.Recordset!ninviv

nompad.Text = bd.Recordset!nompad
ocupad.Text = bd.Recordset!ocupad
ingmenpad.Text = bd.Recordset!ingmenpad
edapad.Text = bd.Recordset!edapad
nomemppad.Text = bd.Recordset!nomemppad
telemppad.Text = bd.Recordset!telemppad
nivacapad.Text = bd.Recordset!nivacapad
otringpad.Text = bd.Recordset!otringpad
nommad.Text = bd.Recordset!nommad
ocumad.Text = bd.Recordset!ocumad
nomempmad.Text = bd.Recordset!nomempmad
ingmenmad.Text = bd.Recordset!ingmenmad
edamad.Text = bd.Recordset!edamad
'If bd.Recordset!telempmad <> Null Then
    telempmad.Text = bd.Recordset!telempmad
'End If
'If bd.Recordset!nivacamad <> Null Then
    nivacamad.Text = bd.Recordset!nivacamad
'End If
'If bd.Recordset!otringmad <> Null Then
    otringmad.Text = bd.Recordset!otringmad
'End If
'If bd.Recordset!tenviv <> Null Then
    tenviv.Text = bd.Recordset!tenviv
'End If
'If bd.Recordset!tipviv <> Null Then
    tipviv.Text = bd.Recordset!tipviv
'End If
'If bd.Recordset!conviv <> Null Then
    conviv.Text = bd.Recordset!conviv
'End If
'If bd.Recordset!estviv <> Null Then
    estviv.Text = bd.Recordset!estviv
'End If
'If bd.Recordset!serpub <> Null Then
   serpub.Text = bd.Recordset!serpub
'End If
SSTab1.Tab = 0
numdoc.SetFocus
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If NuevoRegI = True Then
    If MsgBox("Esta agregando un nuevo registro" & vbCrLf & "Desea continuar?", vbYesNo + vbQuestion, "Inscripciones") = vbYes Then
        Cancel = True
    Else
        NuevoRegI = False
    End If
End If
End Sub

Private Sub Form_Resize()
'Me.Left = (menu.Width - Me.Width) / 2
'Me.Top = ((menu.Height - Me.Height) / 2) -  menu.estado.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
'FormularioActivo = False
menu.estado.Panels(4).Text = "Menú Principal"

DoEvents
End Sub

Private Sub guardar_Click()
guardarregistro
End Sub

Private Sub ingmenmad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    nomempmad.SetFocus
End If
End Sub
Private Sub ingmenpad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    nomemppad.SetFocus
End If
End Sub

Private Sub lugnac_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)

End Sub

Private Sub lugocufam_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    continuar_Click
End If
End Sub

Private Sub modificar_Click()
If bd.Recordset.RecordCount > 0 Then
'desbloquear cajas de texto y combos
ModificadoI = True
numdoc.Locked = False
lugnac.Locked = False
prealgenf.Locked = False
lugocufam.Locked = False
nompad.Locked = False
ocupad.Locked = False
'ingmen.Locked = False
nomemppad.Locked = False
telemppad.Locked = False
nommad.Locked = False
ocumad.Locked = False
nomempmad.Locked = False
ingmenmad.Locked = False
ingmenpad.Locked = False
edamad.Locked = False
edamad.Locked = False
telempmad.Locked = False
otringmad.Locked = False
otringpad.Locked = False
numdoc.Locked = False
numher.Locked = False

busqueda.Enabled = False
modificar.Enabled = False
parametro.Enabled = False
nuevo.Enabled = False
eliminar.Enabled = False

primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False
'Dim modif As Variant
'modif = BD.Recordset.EditMode
Actualizar.Enabled = True
fr.Visible = True
Habilitei inscripciones
End If
End Sub

Private Sub ninviv_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    nompad.SetFocus
End If
End Sub
Private Sub nivacamad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    otringmad.SetFocus
End If
End Sub
Private Sub nivacapad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    otringpad.SetFocus
End If
End Sub
Private Sub nomempmad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    telempmad.SetFocus
End If
End Sub
Private Sub nomemppad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    telemppad.SetFocus
End If
End Sub
Private Sub nommad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    edamad.SetFocus
End If
End Sub
Private Sub nompad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    edapad.SetFocus
End If
End Sub
Private Sub nuevo_Click()
NuevoRegI = True
buscardoc.Enabled = True
'deshabilitar controles
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False
nuevo.Enabled = False
parametro.Enabled = False
guardar.Enabled = True
modificar.Enabled = False
eliminar.Enabled = False

fr.Visible = False

'desbloquear cajas de texto
cajas inscripciones
Habilitei inscripciones

SSTab1.TabVisible(1) = False
SSTab1.TabVisible(2) = False
SSTab1.Tab = 0
'mostrar listado de documentos en listado de espera
numdoc.Clear
ConexionDocu inscripciones, "select * from listadodeespera WHERE INSCRITO=0 ORDER BY numdoc"
If docu.Recordset.RecordCount > 0 Then
    docu.Recordset.MoveFirst
    For i = 1 To docu.Recordset.RecordCount
        numdoc.AddItem Trim$(docu.Recordset!numdoc)
        docu.Recordset.MoveNext
    Next i
End If
busqueda.Enabled = False
numdoc.SetFocus
'genera el numero de inscripcion
ConexionBD1 inscripciones, "select max(numins)as numero from inscripciones"
If IsNull(BD1.Recordset!numero) = False Then
    numins.Caption = (BD1.Recordset!numero) + 1
Else
    numins.Caption = "1"
End If
numher.Locked = False
End Sub
Private Sub numdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    buscardoc_Click
End If

End Sub
Private Sub numher_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    lugocufam.SetFocus
End If
End Sub
Private Sub ocumad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    ingmenmad.SetFocus
End If
End Sub
Private Sub ocupad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    ingmenpad.SetFocus
End If
End Sub
Private Sub otringmad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    continue_Click
End If
End Sub

Private Sub otringpad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    nommad.SetFocus
End If
End Sub

Private Sub parametro_Click()
Para = True
ingresos.Show
End Sub

Private Sub prealgenf_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    numher.SetFocus
End If
End Sub

Private Sub primero_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveFirst
    mostrarcampos
End If

numdoc.SetFocus
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Private Sub serpub_KeyPress(KeyAscii As Integer)
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
numdoc.SetFocus
End Sub
Sub activarcajas()
'bloquear cajas de texto y combos
numdoc.Locked = False
lugnac.Locked = False
prealgenf.Locked = False
lugocufam.Locked = False
nompad.Locked = False
ocupad.Locked = False
'ingmen.Locked = True
nomemppad.Locked = False
telemppad.Locked = False
nommad.Locked = False
ocumad.Locked = False
nomempmad.Locked = False
ingmenmad.Locked = False
ingmenpad.Locked = False
edamad.Locked = False
edamad.Locked = False
telempmad.Locked = False
otringmad.Locked = False
otringpad.Locked = False
numdoc.Locked = False
numher.Locked = False
End Sub

Private Sub telempmad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    nivacamad.SetFocus
End If
End Sub


Private Sub telemppad_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    nivacapad.SetFocus
End If
End Sub

Private Sub tenviv_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    tipviv.SetFocus
End If
End Sub
Private Sub tipviv_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    conviv.SetFocus
End If
End Sub
Function guardarregistro()
Dim NUM_DOC

'validar datos
BD1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PROYECTO.mdb;Mode=Share Deny None;Persist Security Info=False;Jet OLEDB:Database Password=810205;Jet OLEDB:Engine Type=4"
BD1.RecordSource = "SELECT * FROM listadodeespera WHERE numdoc='" & Trim$(numdoc.Text) & "';"
BD1.Refresh


If BD1.Recordset.RecordCount = 0 Then
    MsgBox "Este número de documento no existe en lista de espera!", vbInformation, "Inscripciones"
    Exit Function
End If


If prealgenf.Visible = True And prealgenf.Text = "" Then
    MsgBox "Por favor indique la enfermedad del niño!", vbInformation, "Inscripciones"
    Exit Function
End If

If lugnac.Text = "" Or _
ninviv.Text = "" Or nommad.Text = "" Or edamad.Text = "" Or _
nivacamad.Text = "" Or _
tenviv.Text = "" Or tipviv.Text = "" Or conviv.Text = "" Or estviv.Text = "" Or serpub.Text = "" Then
    MsgBox "Falta datos por ingresar, por favor agreguelos!", vbExclamation, "Inscripciones"
    Exit Function
End If
ConexionBD inscripciones, "select * from inscripciones"
NUM_DOC = numdoc.Text
On Error Resume Next
bd.Recordset.AddNew
bd.Recordset!numdoc = numdoc.Text
bd.Recordset!numins = numins.Caption
bd.Recordset!lugnac = lugnac.Text
If xpradiobutton2.Value = True Then
    bd.Recordset!prealgenf = "NO"
ElseIf xpradiobutton2.Value = False Then
    bd.Recordset!prealgenf = "NO"
ElseIf xpradiobutton1.Value = True Then
    bd.Recordset!prealgenf = prealgenf.Text
End If
bd.Recordset!numher = Val(numher.Text)
bd.Recordset!lugocufam = lugocufam.Text
bd.Recordset!ninviv = ninviv.Text
bd.Recordset!nompad = nompad.Text
bd.Recordset!ocupad = ocupad.Text
bd.Recordset!ingmenpad = Val(ingmenpad.Text)
bd.Recordset!edapad = Val(edapad.Text)
bd.Recordset!nomemppad = nomemppad.Text
bd.Recordset!telemppad = Val(telemppad.Text)
bd.Recordset!nivacapad = nivacapad.Text
bd.Recordset!otringpad = otringpad.Text
bd.Recordset!nommad = nommad.Text
bd.Recordset!ocumad = ocumad.Text
bd.Recordset!nomempmad = nomempmad.Text
bd.Recordset!ingmenmad = Val(ingmenmad.Text)
bd.Recordset!edamad = edamad.Text
bd.Recordset!telempmad = telempmad.Text
bd.Recordset!nivacamad = nivacamad.Text
bd.Recordset!otringmad = Val(otringmad.Text)
bd.Recordset!tenviv = tenviv.Text
bd.Recordset!tipviv = tipviv.Text
bd.Recordset!conviv = conviv.Text
bd.Recordset!estviv = estviv.Text
bd.Recordset!serpub = serpub.Text
bd.Recordset.Update
bd.Refresh
'modifica el registro en listado de espera para asiganarle que ya esta inscrito
ConexionBD1 inscripciones, "select * from listadodeespera where numdoc='" & NUM_DOC & "'"
Dim modi
modi = BD1.Recordset.EditMode
BD1.Recordset!inscrito = 1
BD1.Recordset.Update
On Error Resume Next
ConexionBD inscripciones, "select * from inscripciones where matriculado=0"
If bd.Recordset.RecordCount > 0 Then
    total = bd.Recordset.RecordCount & " inscritos."
    bd.Recordset.MoveFirst
    mostrarcampos
End If
Deshabilitari inscripciones
busqueda.Enabled = True
bd.Refresh
total = bd.Recordset.RecordCount & " inscritos."
NuevoRegI = False
nuevo.Enabled = True
modificar.Enabled = True
eliminar.Enabled = True
primero.Enabled = True
anterior.Enabled = True
siguiente.Enabled = True
ultimo.Enabled = True
parametro.Enabled = True
guardar.Enabled = False

'bloquear cajas de texto y combos
numdoc.Locked = True
lugnac.Locked = True
prealgenf.Locked = True
lugocufam.Locked = True
nompad.Locked = True
ocupad.Locked = True
'ingmen.Locked = True
nomemppad.Locked = True
telemppad.Locked = True
nommad.Locked = True
ocumad.Locked = True
nomempmad.Locked = True
ingmenmad.Locked = True
ingmenpad.Locked = True
edamad.Locked = True
edamad.Locked = True
telempmad.Locked = True
otringmad.Locked = True
otringpad.Locked = True
numdoc.Locked = True
numher.Locked = True

SSTab1.Tab = 0

End Function

Private Sub ultimo_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveLast
    mostrarcampos
End If
numdoc.SetFocus
End Sub

Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "inscripciones.htm"
End Sub

Private Sub xpradiobutton1_Click()
prealgenf.Visible = True
prealgenf.SetFocus
End Sub

Private Sub xpradiobutton2_Click()
prealgenf.Visible = False
End Sub


