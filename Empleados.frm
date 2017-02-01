VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form empleado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empleados"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Empleados.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9615
   Begin MSAdodcLib.Adodc bd 
      Height          =   330
      Left            =   8040
      Top             =   4080
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
   Begin MSAdodcLib.Adodc bd1 
      Height          =   330
      Left            =   8040
      Top             =   3720
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
      Height          =   5535
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7575
      _ExtentX        =   13785
      _ExtentY        =   11245
      Caption         =   "Datos Empleado"
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
      Begin VB.ComboBox BAR 
         Height          =   315
         Left            =   4080
         TabIndex        =   45
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ComboBox pro 
         Height          =   315
         Left            =   4080
         TabIndex        =   44
         Top             =   3000
         Width           =   2175
      End
      Begin VB.ComboBox nivestnofor 
         Height          =   315
         ItemData        =   "Empleados.frx":628A
         Left            =   240
         List            =   "Empleados.frx":628C
         TabIndex        =   42
         Top             =   5040
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker fecnac 
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   3600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50266113
         CurrentDate     =   38118
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecvin 
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50266113
         CurrentDate     =   38118
      End
      Begin VB.TextBox numdoc 
         Height          =   285
         Left            =   240
         MaxLength       =   12
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox dir 
         Height          =   285
         Left            =   240
         MaxLength       =   35
         TabIndex        =   4
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox tel 
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox exp 
         Height          =   315
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox ape 
         Height          =   315
         Left            =   4080
         MaxLength       =   25
         TabIndex        =   3
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox cor 
         Height          =   315
         Left            =   4080
         MaxLength       =   25
         OLEDropMode     =   2  'Automatic
         TabIndex        =   6
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox nom 
         Height          =   315
         Left            =   240
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox cel 
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   7
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox car 
         Height          =   315
         ItemData        =   "Empleados.frx":628E
         Left            =   240
         List            =   "Empleados.frx":6290
         TabIndex        =   8
         Top             =   3600
         Width           =   2415
      End
      Begin VB.ComboBox nivestfor 
         Height          =   315
         ItemData        =   "Empleados.frx":6292
         Left            =   4080
         List            =   "Empleados.frx":6294
         TabIndex        =   11
         Top             =   4320
         Width           =   2535
      End
      Begin Jardin.xphelp xphelp1 
         Height          =   315
         Left            =   6960
         Top             =   5040
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Formato Fecha: dia/mes/año"
         Height          =   195
         Left            =   4080
         TabIndex        =   43
         Top             =   5160
         Width           =   2475
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Documento Identidad"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   2190
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombres"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Expedida en"
         Height          =   255
         Left            =   4080
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos"
         Height          =   195
         Left            =   4080
         TabIndex        =   22
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barrio"
         Height          =   195
         Left            =   4080
         TabIndex        =   21
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correo Electrónico"
         Height          =   195
         Left            =   4080
         TabIndex        =   20
         Top             =   2160
         Width           =   1590
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Celular"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   3360
         Width           =   525
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Vinculación"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   3960
         Width           =   1785
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel de Estudios No Formal"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   4680
         Width           =   2385
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Profesión"
         Height          =   195
         Left            =   4080
         TabIndex        =   15
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Nacimiento"
         Height          =   195
         Left            =   4080
         TabIndex        =   14
         Top             =   3360
         Width           =   1770
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel de Estudios Formal"
         Height          =   195
         Left            =   4080
         TabIndex        =   13
         Top             =   4080
         Width           =   2100
      End
   End
   Begin Jardin.xpgroupbox xpgroupbox7 
      Height          =   3615
      Left            =   7800
      TabIndex        =   28
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
         TabIndex        =   29
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
         MICON           =   "Empleados.frx":6296
         BC              =   8438015
         FC              =   0
         Picture         =   "Empleados.frx":6404
      End
      Begin JeweledBut.JeweledButton busqueda 
         Height          =   375
         Left            =   120
         TabIndex        =   30
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
         MICON           =   "Empleados.frx":8F0E
         BC              =   8438015
         FC              =   0
         Picture         =   "Empleados.frx":907C
      End
      Begin JeweledBut.JeweledButton eliminar 
         Height          =   375
         Left            =   120
         TabIndex        =   31
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
         MICON           =   "Empleados.frx":91D6
         BC              =   8438015
         FC              =   0
         Picture         =   "Empleados.frx":9344
      End
      Begin JeweledBut.JeweledButton guardar 
         Height          =   375
         Left            =   120
         TabIndex        =   32
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
         MICON           =   "Empleados.frx":98DE
         BC              =   8438015
         FC              =   0
         Picture         =   "Empleados.frx":9A4C
      End
      Begin JeweledBut.JeweledButton Actualizar 
         Height          =   375
         Left            =   120
         TabIndex        =   33
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
         MICON           =   "Empleados.frx":F67A
         BC              =   8438015
         FC              =   0
         Picture         =   "Empleados.frx":F7E8
      End
      Begin JeweledBut.JeweledButton modificar 
         Height          =   375
         Left            =   120
         TabIndex        =   34
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
         MICON           =   "Empleados.frx":FD82
         BC              =   8438015
         FC              =   0
         Picture         =   "Empleados.frx":FEF0
      End
      Begin JeweledBut.JeweledButton parametro 
         Height          =   375
         Left            =   120
         TabIndex        =   47
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
         MICON           =   "Empleados.frx":1004A
         BC              =   8438015
         FC              =   0
      End
   End
   Begin JeweledBut.JeweledButton salir 
      Height          =   375
      Left            =   7920
      TabIndex        =   35
      Top             =   6240
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
      MICON           =   "Empleados.frx":101B8
      BC              =   8438015
      FC              =   0
      Picture         =   "Empleados.frx":10326
   End
   Begin Jardin.xpgroupbox xpgroupbox8 
      Height          =   855
      Left            =   120
      TabIndex        =   36
      Top             =   5760
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
         TabIndex        =   37
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
         MICON           =   "Empleados.frx":10480
         BC              =   8438015
         FC              =   0
         Picture         =   "Empleados.frx":105EE
      End
      Begin JeweledBut.JeweledButton siguiente 
         Height          =   375
         Left            =   3480
         TabIndex        =   38
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
         MICON           =   "Empleados.frx":10748
         BC              =   8438015
         FC              =   0
         Picture         =   "Empleados.frx":108B6
      End
      Begin JeweledBut.JeweledButton ultimo 
         Height          =   375
         Left            =   5160
         TabIndex        =   39
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
         MICON           =   "Empleados.frx":10A10
         BC              =   8438015
         FC              =   0
         Picture         =   "Empleados.frx":10B7E
      End
      Begin JeweledBut.JeweledButton anterior 
         Height          =   375
         Left            =   1800
         TabIndex        =   40
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
         MICON           =   "Empleados.frx":10CD8
         BC              =   8438015
         FC              =   0
         Picture         =   "Empleados.frx":10E46
      End
   End
   Begin JeweledBut.JeweledButton cancelar 
      Height          =   375
      Left            =   7920
      TabIndex        =   48
      Top             =   5760
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
      MICON           =   "Empleados.frx":10FA0
      BC              =   8438015
      FC              =   0
      Picture         =   "Empleados.frx":1110E
   End
   Begin VB.Label numreg 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7920
      TabIndex        =   46
      Top             =   4320
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   8040
      Picture         =   "Empleados.frx":11268
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label em 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   7920
      TabIndex        =   41
      Top             =   3960
      Width           =   60
   End
End
Attribute VB_Name = "empleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Actualizar_Click()

'habilitar controles
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
If ModificadoE = True Then
    bd.Recordset.Delete
    bd.Refresh
    guardarregistro
    ModificadoE = False
    Bloquear
End If
End Sub
Function avanzar()
 If tecla = 13 Then
  SendKeys "{tab}"
  KeyAscii = 0: tecla = 0
 End If
End Function

Private Sub anterior_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MovePrevious
    If bd.Recordset.BOF Then
        bd.Recordset.MoveFirst
    End If
    mostrarcampos
End If
numdoc.SetFocus
Bloquear
End Sub

Private Sub ape_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    dir.SetFocus
End If
End Sub

Private Sub bar_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If

KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    tel.SetFocus
End If
End Sub

Private Sub bd_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub bd1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub busqueda_Click()
MB.Formulario = Me.Name
MB.Descripcion = "Empleados"
elBuscador.Show
End Sub

Private Sub cancelar_Click()
If bd.Recordset.RecordCount > 0 Then
    mostrarcampos
End If
If NuevoRegE = True Then
    nuevo.Enabled = True
    modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    anterior.Enabled = True
    siguiente.Enabled = True
    ultimo.Enabled = True
    guardar.Enabled = False
    busqueda.Enabled = True
    NuevoRegE = False
ElseIf ModificadoE = True Then
    nuevo.Enabled = True
    modificar.Enabled = True
    eliminar.Enabled = True
    primero.Enabled = True
    ultimo.Enabled = True
    siguiente.Enabled = True
    anterior.Enabled = True
    Actualizar.Enabled = False
    busqueda.Enabled = True
    ModificadoE = False
End If
 parametro.Enabled = True
 Bloquear
End Sub

Private Sub car_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If

KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    fecnac.SetFocus
End If
End Sub

Private Sub cel_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    pro.SetFocus
End If
End Sub

Private Sub cor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cel.SetFocus
End If
End Sub

Private Sub dir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    bar.SetFocus
End If
End Sub

Private Sub eliminar_Click()
If bd.Recordset.RecordCount > 0 Then
If MsgBox("Está seguro de querer eliminar el registro?", vbYesNo + vbQuestion, "Eliminar Registro") = vbYes Then
   bd.Recordset.Delete
   If bd.Recordset.RecordCount > 0 Then
    bd.Recordset.MoveFirst
    bd.Refresh
    em.Caption = bd.Recordset.RecordCount & " Empleados"
    mostrarcampos
    Else
        Unload Me
        empleado.Show
   End If
End If
End If
End Sub

Private Sub exp_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    nom.SetFocus
End If
End Sub



Private Sub fecnac_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    fecvin.SetFocus
End If
End Sub



Private Sub fecnac_LostFocus()
Dim ANIOs As Integer
ANIOs = Format(Date, "yyyy")

If fecnac.Year > ANIOs Then
    MsgBox "Fecha no permitida!", vbExclamation, "Pagos"
    fecnac.SetFocus
    Exit Sub
End If
End Sub

Private Sub fecvin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    nivestfor.SetFocus
End If
End Sub
Private Sub LlenarCombos()
'llenar cargo
ConexionBD1 empleado, "select * from parametrizacion where tippar=11" & " order by dato;"
car.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        car.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar BARRIO
ConexionBD1 empleado, "select * from parametrizacion where tippar=16" & " order by dato;"
bar.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        bar.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar profesion
ConexionBD1 empleado, "select * from parametrizacion where tippar=23" & " order by dato;"
pro.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        pro.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar estudio formal
ConexionBD1 empleado, "select * from parametrizacion where tippar=3" & " order by dato;"
nivestfor.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        nivestfor.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
'llenar estudio no formal
ConexionBD1 empleado, "select * from parametrizacion where tippar=19" & " order by dato;"
nivestnofor.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        nivestnofor.AddItem BD1.Recordset!dato
        BD1.Recordset.MoveNext
    Next i
End If
End Sub

Private Sub Form_Activate()
FormularioActivo = True
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height


'bloquear cajas y combos
exp.Locked = True
nom.Locked = True
ape.Locked = True
dir.Locked = True
bar.Locked = True
tel.Locked = True
cor.Locked = True
cel.Locked = True
pro.Locked = True
car.Locked = True
nivestfor.Locked = True
nivestnofor.Locked = True
numdoc.Locked = True
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
menu.estado.Panels(4).Text = "Control de los Empleados del Jardín"
On Error Resume Next
ConexionBD empleado, "select * from empleado"
If bd.Recordset.RecordCount > 0 Then
   mostrarcampos
   em.Caption = bd.Recordset.RecordCount & " Empleados"
End If

End Sub
Function mostrarcampos()
numreg = bd.Recordset.AbsolutePosition & " registro."
numdoc.Text = bd.Recordset!numdoc
exp.Text = bd.Recordset!exp
nom.Text = bd.Recordset!nom
ape.Text = bd.Recordset!ape
dir.Text = bd.Recordset!dir
bar.Text = bd.Recordset!bar
If IsNull(bd.Recordset!tel) = False Then
    tel.Text = bd.Recordset!tel
End If
If IsNull(bd.Recordset!cor) = False Then
    cor.Text = bd.Recordset!cor
End If
If IsNull(bd.Recordset!cel) = False Then
    cel.Text = bd.Recordset!cel
End If
If IsNull(bd.Recordset!pro) = False Then
    pro.Text = bd.Recordset!pro
End If
If IsNull(bd.Recordset!car) = False Then
    car.Text = bd.Recordset!car
End If
If IsNull(bd.Recordset!fecvin) = False Then
    fecvin.Value = bd.Recordset!fecvin
End If
If IsNull(bd.Recordset!fecnac) = False Then
    fecnac.Value = bd.Recordset!fecnac
End If
If IsNull(bd.Recordset!nivestfor) = False Then
    nivestfor.Text = bd.Recordset!nivestfor
End If
If IsNull(bd.Recordset!niveestnofor) = False Then
    nivestnofor.Text = bd.Recordset!niveestnofor
End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If NuevoRegE = True Then
    If MsgBox("Esta agregando un nuevo registro" & vbCrLf & "Desea continuar?", vbYesNo + vbQuestion, "Empleados") = vbYes Then
        Cancel = True
    Else
        NuevoRegE = False
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
bd.Recordset.Close
Set bd.Recordset = Nothing
End Sub

Private Sub guardar_Click()
guardarregistro

End Sub
Function guardarregistro()
Dim ANIOs1, anios2 As Integer
ANIOs1 = fecnac.Year
anios2 = (fecvin.Year - ANIOs1)

If fecvin.Year <= ANIOs1 Then
    MsgBox "Fecha no permitida. La fecha de vinculación debe ser por los menos" & vbCrLf & "mayor 15 años a la de nacimiento!", vbExclamation, "Empleados"
    fecvin.SetFocus
    Exit Function
End If
If anios2 < 15 Then
    MsgBox "Fecha no permitida. La fecha de vinculación debe ser por los menos" & vbCrLf & "mayor 15 años a la de nacimiento!", vbExclamation, "Empleados"
    fecvin.SetFocus
    Exit Function
End If
If numdoc.Text = "" Or nom.Text = "" Or ape.Text = "" Or _
dir.Text = "" Or pro.Text = "" Or car.Text = "" Or _
nivestfor.Text = "" Then
    MsgBox "Hace falta datos por ingresar!", vbInformation, "Empleados"
    Exit Function
End If

On Error Resume Next
bd.Recordset.AddNew
bd.Recordset!numdoc = numdoc.Text
bd.Recordset!exp = exp.Text
bd.Recordset!nom = nom.Text
bd.Recordset!ape = ape.Text
bd.Recordset!dir = dir.Text
bd.Recordset!bar = bar.Text
bd.Recordset!tel = tel.Text
bd.Recordset!cor = cor.Text
bd.Recordset!cel = cel.Text
bd.Recordset!pro = pro.Text
bd.Recordset!car = car.Text
bd.Recordset!fecvin = fecvin.Value
bd.Recordset!fecnac = fecnac.Value
bd.Recordset!nivestfor = nivestfor.Text
bd.Recordset!niveestnofor = nivestnofor.Text
bd.Recordset.Update
'habilitar controles
nuevo.Enabled = True
modificar.Enabled = True
eliminar.Enabled = True
primero.Enabled = True
anterior.Enabled = True
busqueda.Enabled = True
siguiente.Enabled = True
ultimo.Enabled = True
guardar.Enabled = False
parametro.Enabled = True
bd.Refresh
em.Caption = bd.Recordset.RecordCount & " Empleados"
eliminar.Enabled = True
NuevoRegE = False
Bloquear
End Function

Private Sub modificar_Click()
If bd.Recordset.RecordCount > 0 Then
ModificadoE = True
'desbloquear cajas y combos
parametro.Enabled = False
exp.Locked = False
nom.Locked = False
ape.Locked = False
dir.Locked = False
bar.Locked = False
tel.Locked = False
cor.Locked = False
cel.Locked = False
pro.Locked = False
car.Locked = False
nivestfor.Locked = False
nivestnofor.Locked = False
numdoc.Locked = False

'deshabilitar controles
modificar.Enabled = False
Actualizar.Enabled = True
nuevo.Enabled = False
eliminar.Enabled = False
busqueda.Enabled = False
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False

End If
End Sub

Private Sub nivestfor_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If

KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    nivestnofor.SetFocus
End If
End Sub

Private Sub nivestnofor_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If

KeyAscii = Validar_letra(KeyAscii)
End Sub

Private Sub nom_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    ape.SetFocus
End If
End Sub

Private Sub nuevo_Click()
NuevoRegE = True
'deshabilitar controles
primero.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
ultimo.Enabled = False
nuevo.Enabled = False
guardar.Enabled = True
modificar.Enabled = False
eliminar.Enabled = False
busqueda.Enabled = False
parametro.Enabled = False
'desbloquear cajas y combos
exp.Locked = False
nom.Locked = False
ape.Locked = False
dir.Locked = False
bar.Locked = False
tel.Locked = False
cor.Locked = False
cel.Locked = False
pro.Locked = False
car.Locked = False
nivestfor.Locked = False
nivestnofor.Locked = False
numdoc.Locked = False


'limpiar
exp = ""
nom = ""
ape = ""
dir = ""
bar = ""
tel = ""
cor = ""
cel = ""
pro = ""
car = ""


nivestfor = ""
nivestnofor = ""
numdoc = ""
numdoc.SetFocus
End Sub

Private Sub numdocide_KeyPress(KeyAscii As Integer)
tecla = KeyAscii
avanzar
End Sub

Private Sub numdoc_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    ConexionBD1 empleado, "select * from empleado where numdoc=" & numdoc.Text
    If BD1.Recordset.RecordCount > 0 Then
        MsgBox "Empleado ya existente!", vbInformation, "Empleados"
        numdoc.SetFocus
        Exit Sub
    End If
    exp.SetFocus
End If
End Sub

Private Sub parametro_Click()
Para = True
ingresos.Show
End Sub

Private Sub primero_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveFirst
    mostrarcampos
End If
Bloquear
numdoc.SetFocus
End Sub
Sub Bloquear()
'bloquear cajas y combos
exp.Locked = True
nom.Locked = True
ape.Locked = True
dir.Locked = True
bar.Locked = True
tel.Locked = True
cor.Locked = True
cel.Locked = True
pro.Locked = True
car.Locked = True
nivestfor.Locked = True
nivestnofor.Locked = True
numdoc.Locked = True
End Sub
Private Sub pro_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If

KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    car.SetFocus
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
numdoc.SetFocus
Bloquear
End Sub
Private Sub tel_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    cor.SetFocus
End If
End Sub

Private Sub ultimo_Click()
If bd.Recordset.RecordCount <> 0 Then
    bd.Recordset.MoveLast
    mostrarcampos
End If
numdoc.SetFocus
Bloquear
End Sub

Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "empleados.htm"
End Sub
