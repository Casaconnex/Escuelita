VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form certificados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Certificados, Constancias, Diplomas y Carnet's"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "certificados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc bd 
      Height          =   330
      Left            =   360
      Top             =   4320
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8640
      Top             =   5040
   End
   Begin Jardin.xphelp xphelp1 
      Height          =   315
      Left            =   1920
      Top             =   3120
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin Jardin.xpgroupbox xpgroupbox1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   4895
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
      Begin JeweledBut.JeweledButton certi 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Certificados"
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
         MICON           =   "certificados.frx":2AFA
         BC              =   8438015
         FC              =   0
         Picture         =   "certificados.frx":2C68
      End
      Begin JeweledBut.JeweledButton comu 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Constancias"
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
         MICON           =   "certificados.frx":4972
         BC              =   8438015
         FC              =   0
         Picture         =   "certificados.frx":4AE0
      End
      Begin JeweledBut.JeweledButton salir 
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   2280
         Width           =   1935
         _ExtentX        =   3413
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
         MICON           =   "certificados.frx":75EA
         BC              =   8438015
         FC              =   0
         Picture         =   "certificados.frx":7758
      End
      Begin JeweledBut.JeweledButton diplos 
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Diplomas"
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
         MICON           =   "certificados.frx":78B2
         BC              =   8438015
         FC              =   0
         Picture         =   "certificados.frx":7A20
      End
      Begin JeweledBut.JeweledButton carnets 
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Carnet's"
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
         MICON           =   "certificados.frx":A52A
         BC              =   8438015
         FC              =   0
         Picture         =   "certificados.frx":A698
      End
   End
   Begin Jardin.xpgroupbox grupocerti 
      Height          =   4455
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7858
      Caption         =   "Datos Certificados"
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
      Begin VB.ComboBox numdocc 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox certihora2 
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "5:00 PM"
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox certicargo 
         Height          =   285
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   16
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox certinomex 
         Height          =   285
         Left            =   2520
         MaxLength       =   35
         TabIndex        =   15
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox certihora1 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "7:00 AM"
         Top             =   1920
         Width           =   975
      End
      Begin MSComCtl2.DTPicker certifecha 
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19988481
         CurrentDate     =   38132
      End
      Begin VB.TextBox certigrado 
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox certinombre 
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Top             =   960
         Width           =   3375
      End
      Begin JeweledBut.JeweledButton vistaprevia 
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   3960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         TX              =   "Imprimir"
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
         MICON           =   "certificados.frx":D1A2
         BC              =   8438015
         FC              =   0
         Picture         =   "certificados.frx":D310
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento del niño:"
         Height          =   195
         Left            =   240
         TabIndex        =   61
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         Height          =   195
         Left            =   3555
         TabIndex        =   18
         Top             =   2040
         Width           =   105
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo de quien expide:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   3600
         Width           =   2025
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de quien expide:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Expedición:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horario:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grado:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del niño:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1485
      End
   End
   Begin Jardin.xpgroupbox constancias 
      Height          =   4815
      Left            =   2520
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8493
      Caption         =   "Datos Constancias"
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
      Begin MSAdodcLib.Adodc bd1 
         Height          =   330
         Left            =   720
         Top             =   4440
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.TextBox ctiempoe 
         Height          =   285
         Left            =   2400
         TabIndex        =   40
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox ccargoe 
         Height          =   285
         Left            =   2400
         TabIndex        =   37
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox clugare 
         Height          =   285
         Left            =   4200
         TabIndex        =   36
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox ccedulae 
         Height          =   285
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   33
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox clugar 
         Height          =   285
         Left            =   4200
         TabIndex        =   31
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox cnombrex 
         Height          =   285
         Left            =   2400
         TabIndex        =   24
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox ccedula 
         Height          =   285
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   23
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox cnombree 
         Height          =   285
         Left            =   2400
         TabIndex        =   21
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox ccargoex 
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   1320
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker cfechaex 
         Height          =   375
         Left            =   2400
         TabIndex        =   22
         Top             =   3720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19988481
         CurrentDate     =   38132
      End
      Begin JeweledBut.JeweledButton cvista 
         Height          =   375
         Left            =   3840
         TabIndex        =   25
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         TX              =   "Imprimir"
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
         MICON           =   "certificados.frx":D8AA
         BC              =   8438015
         FC              =   0
         Picture         =   "certificados.frx":DA18
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo laborando:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   3360
         Width           =   1620
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo empleado:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cedula:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   2400
         Width           =   675
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "de:"
         Height          =   195
         Left            =   3840
         TabIndex        =   34
         Top             =   2400
         Width           =   285
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre empleado:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   1650
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cedula:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "de:"
         Height          =   195
         Left            =   3840
         TabIndex        =   29
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Expedición:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   3840
         Width           =   1545
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de quien expide:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo de quien expide:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   2025
      End
   End
   Begin Jardin.xpgroupbox diploma 
      Height          =   3375
      Left            =   2520
      TabIndex        =   42
      Top             =   240
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4895
      Caption         =   "Datos Diplomas"
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
      Begin VB.ComboBox numdocd 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox apediplo 
         Height          =   285
         Left            =   1560
         TabIndex        =   52
         Top             =   1800
         Width           =   4815
      End
      Begin MSComCtl2.DTPicker fecdiplo 
         Height          =   375
         Left            =   1560
         TabIndex        =   48
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19988481
         CurrentDate     =   38140
      End
      Begin VB.TextBox nombrediplo 
         Height          =   285
         Left            =   1560
         TabIndex        =   45
         Top             =   1320
         Width           =   4815
      End
      Begin JeweledBut.JeweledButton vistadiplo 
         Height          =   375
         Left            =   4200
         TabIndex        =   49
         Top             =   2880
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         TX              =   "Imprimir"
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
         MICON           =   "certificados.frx":DFB2
         BC              =   8438015
         FC              =   0
         Picture         =   "certificados.frx":E120
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos:"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombres:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   2400
         Width           =   570
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Otorgado a:"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   1035
      End
   End
   Begin Jardin.xpgroupbox carne 
      Height          =   4455
      Left            =   2520
      TabIndex        =   53
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7858
      Caption         =   "Datos Carnet's"
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
         Left            =   2400
         TabIndex        =   66
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox profecarne 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1215
         MaxLength       =   20
         TabIndex        =   57
         Top             =   3280
         Width           =   2965
      End
      Begin VB.TextBox nivelcarne 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   840
         MaxLength       =   15
         TabIndex        =   56
         Top             =   2920
         Width           =   3345
      End
      Begin VB.TextBox apecarne 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   360
         MaxLength       =   20
         TabIndex        =   55
         Top             =   2540
         Width           =   3820
      End
      Begin VB.TextBox nombrecarne 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   54
         Top             =   2160
         Width           =   3080
      End
      Begin JeweledBut.JeweledButton vistacarne 
         Height          =   375
         Left            =   3600
         TabIndex        =   58
         Top             =   3960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         TX              =   "Imprimir"
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
         MICON           =   "certificados.frx":E6BA
         BC              =   8438015
         FC              =   0
         Picture         =   "certificados.frx":E828
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Documento:"
         Height          =   195
         Left            =   240
         TabIndex        =   65
         Top             =   360
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   3015
         Left            =   120
         Picture         =   "certificados.frx":108AA
         Stretch         =   -1  'True
         Top             =   840
         Width           =   5655
      End
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000018&
      Caption         =   "Formato Fecha: dia/mes/año"
      Height          =   435
      Left            =   120
      TabIndex        =   60
      Top             =   3120
      Width           =   1395
   End
   Begin VB.Image Image2 
      Height          =   2055
      Left            =   4800
      Picture         =   "certificados.frx":7D5D1C
      Stretch         =   -1  'True
      Top             =   900
      Width           =   1455
   End
   Begin VB.Label aviso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Por favor no guarde los cambios al cerrar Word!!!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2520
      TabIndex        =   59
      Top             =   5520
      Width           =   4890
   End
End
Attribute VB_Name = "certificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Certificado As Word.Application
Private Constancia As Word.Application
Private CDiploma As Word.Application
Private CCarnet As Word.Application

Private Sub apecarne_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    nivelcarne.SetFocus
End If
End Sub

Private Sub apediplo_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
   fecdiplo.SetFocus
End If
End Sub

Private Sub carnets_Click()
diploma.Visible = False
grupocerti.Visible = False
constancias.Visible = False
carne.Visible = True
nombrecarne.SetFocus
'llenar numero documento
ConexionBD1 certificados, "select * from matricula"
numdoc.Clear
If BD1.Recordset.RecordCount > 0 Then
    BD1.Recordset.MoveFirst
    For i = 1 To BD1.Recordset.RecordCount
        numdoc.AddItem BD1.Recordset!numdoc
        BD1.Recordset.MoveNext
    Next i
End If
End Sub

Private Sub ccargoe_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    ctiempoe.SetFocus
End If
End Sub

Private Sub ccargoex_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    cnombree.SetFocus
End If
End Sub

Private Sub ccedula_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    If ccedula.Text = ccedulae.Text Then
        MsgBox "El mismo empleado no puede darse una constancia!", vbInformation, "Constancias"
        Exit Sub
    End If
    'carga datos del empleado
    On Error Resume Next
    ConexionBD1 certificados, "select * from empleado where numdoc=" & ccedula.Text
    If BD1.Recordset.RecordCount > 0 Then
        cnombrex.Text = BD1.Recordset!nom & " " & BD1.Recordset!ape
        clugar.Text = BD1.Recordset!exp
        ccargoex.Text = BD1.Recordset!car
    Else
        MsgBox "La cedula que ingreso no se encuentra registrada!", vbInformation, "Constancias"
        ccedulae.SetFocus
        Exit Sub
    End If
    
    clugar.SetFocus
End If
    
End If
End Sub

Private Sub ccedulae_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    'carga datos del empleado
    On Error Resume Next
    ConexionBD1 certificados, "select * from empleado where numdoc=" & ccedulae.Text
    If BD1.Recordset.RecordCount > 0 Then
        cnombree.Text = BD1.Recordset!nom & " " & BD1.Recordset!ape
        clugare.Text = BD1.Recordset!exp
        ccargoe.Text = BD1.Recordset!car
        Dim anos, meses
        anos = Int((DateValue(Format(Date, "dd/mm/yyyy")) - DateValue(BD1.Recordset!fecvin)) / 365)
        meses = Mid((DateValue(Format(Date, "dd/mm/yyyy")) - DateValue(BD1.Recordset!fecvin)) / 365, 3, 1)
        If anos <> 0 And meses <> 0 Then
            ctiempoe.Text = anos & " años y " & meses & " meses"
        ElseIf anos = 0 And meses <> 0 Then
            ctiempoe.Text = meses & " meses"
        ElseIf meses = 0 And anos <> 0 Then
            ctiempoe.Text = anos & " años"
        End If
    Else
        MsgBox "La cedula que ingreso no se encuentra registrada!", vbInformation, "Constancias"
        ccedulae.SetFocus
        Exit Sub
    End If
    
    clugare.SetFocus
End If
End Sub

Private Sub certi_Click()
constancias.Visible = False
grupocerti.Visible = True
diploma.Visible = False
carne.Visible = False
ConexionBD certificados, "select numdoc from matricula"
numdocc.Clear
If bd.Recordset.RecordCount > 0 Then
    bd.Recordset.MoveFirst
    For i = 1 To bd.Recordset.RecordCount
        numdocc.AddItem bd.Recordset!numdoc
        bd.Recordset.MoveNext
    Next i
End If
numdocc.SetFocus
End Sub

Private Sub certicargo_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    vistaprevia_Click
End If
End Sub

Private Sub certifecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    certinomex.SetFocus
End If
End Sub

Private Sub certigrado_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    certihora1.SetFocus
End If
End Sub

Private Sub certihora1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    certihora2.SetFocus
End If
End Sub

Private Sub certihora2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    certifecha.SetFocus
End If
End Sub

Private Sub certinombre_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    certigrado.SetFocus
End If
End Sub

Private Sub certinomex_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    certicargo.SetFocus
End If
End Sub
Private Sub cfechaex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cvista_Click
End If
End Sub

Private Sub clugar_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    ccargoex.SetFocus
End If
End Sub

Private Sub clugare_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    ccargoe.SetFocus
End If
End Sub

Private Sub cnombree_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    ccedulae.SetFocus
End If
End Sub

Private Sub cnombrex_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    ccedula.SetFocus
End If
End Sub

Private Sub comu_Click()
grupocerti.Visible = False
constancias.Visible = True
diploma.Visible = False
carne.Visible = False
cnombrex.SetFocus
cfechaex.Value = Format(Date, "dd/mm/yyyy")
End Sub
Private Sub ctiempo_Click()
cfechaex.SetFocus
End Sub

Private Sub ctiempo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cfechaex.SetFocus
End If
End Sub

Private Sub ctiempoe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ctiempo.SetFocus
End If
End Sub

Private Sub cursodiplo_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    fecdiplo.SetFocus
End If
End Sub

Private Sub cvista_Click()
If cnombrex.Text = "" Or ccedula.Text = "" Or clugar.Text = "" Or _
ccargoex.Text = "" Or cnombree.Text = "" Or ccedulae.Text = "" Or clugare.Text = "" Or _
ccargoe.Text = "" Or ctiempoe.Text = "" Then
    MsgBox "Falta datos por ingresar!", vbInformation, "Constancias"
    Exit Sub
End If
If ccedula.Text = ccedulae.Text Then
        MsgBox "El mismo empleado no puede darse una constancia!", vbInformation, "Constancias"
        Exit Sub
End If
If MsgBox("Desea imprimir?", vbYesNo + vbQuestion, "Constancias") = vbYes Then
    On Error Resume Next
    Set Constancia = New Word.Application
    Constancia.Documents.Open App.Path & "\certi.doc"
    Constancia.Selection.Paragraphs.Alignment = wdAlignParagraphCenter
    Constancia.WordBasic.Bold
    Constancia.Selection.Font.Size = 16
    Constancia.Selection.Font.Name = "Verdana"
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert "A QUIEN INTERESE"
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.Selection.Paragraphs.Alignment = wdAlignParagraphJustify
    Constancia.Selection.Font.Size = 12
    Constancia.WordBasic.Insert "Yo " & cnombrex.Text & " identificado(a) con número de cédula " & ccedula.Text & " de " & clugar.Text & ". Certifico que " & cnombree.Text & " identificado(a) con número de cédula " & ccedulae.Text & " de " & clugare.Text & ". Labora en la Corporación Centro de Desarrollo Comunitario CODEC desde hace " & ctiempoe.Text & " desempeñandose en el cargo de " & ccargoe.Text & "."
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert "Se expide a petición verbal del interesado a los " & cfechaex.Day & " días del mes de " & mes(cfechaex.Month) & " del año " & cfechaex.Year
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert "Cordialmente,"
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert "__________________________________"
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert cnombrex.Text
    Constancia.WordBasic.Insert vbCrLf
    Constancia.WordBasic.Insert ccargoex.Text
    Constancia.ActiveDocument.PrintOut
    Constancia.ActiveDocument.Close wdDoNotSaveChanges
    Constancia.Application.Quit
    'Constancia.Application.Visible = True
    'Constancia.ActiveDocument.PrintPreview
    DoEvents
    If Err.Number <> 0 Then
        MsgBox "Su versión de Microsoft Word no permite " & vbCrLf & "realizar esta acción!", vbInformation, "Constancias, Diplomas, Certificados..."
        Exit Sub
    End If
End If
End Sub

Private Sub diplos_Click()
ConexionBD certificados, "select numdoc from matricula order by numdoc"
diploma.Visible = True
grupocerti.Visible = False
constancias.Visible = False
carne.Visible = False
numdocd.Clear
If bd.Recordset.RecordCount > 0 Then
    bd.Recordset.MoveFirst
    For i = 1 To bd.Recordset.RecordCount
        numdocd.AddItem bd.Recordset!numdoc
        bd.Recordset.MoveNext
    Next i
End If
numdocd.SetFocus

End Sub

Private Sub fecdiplo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    vistadiplo_Click
End If
End Sub

Private Sub Form_Activate()
If Matriculado = False Then
    FormularioActivo = True
ElseIf Matriculado = True Then
    carnets_Click
End If
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height
End Sub


Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
menu.estado.Panels(4).Text = "Generador de Carnet's, Certificados, Diplomas y Constancias"
End Sub

Private Sub Form_Resize()
Me.Left = (menu.Width - Me.Width) / 2
Me.Top = ((menu.Height - Me.Height) / 2) - menu.estado.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormularioActivo = False
menu.estado.Panels(4).Text = "Menú Principal"
carnete = False
End Sub

Private Sub nivelcarne_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    profecarne.SetFocus
End If
End Sub

Private Sub nombrecarne_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    apecarne.SetFocus
End If
End Sub

Private Sub nombrediplo_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    apediplo.SetFocus
End If
End Sub



Private Sub numdoc_Click()
ConexionBD1 certificados, "select * from listadodeespera where numdoc='" & numdoc.Text & "'"
If BD1.Recordset.RecordCount > 0 Then
    nombrecarne.Text = BD1.Recordset!prinom & " " & BD1.Recordset!segnom
    apecarne.Text = BD1.Recordset!priape & " " & BD1.Recordset!segape
End If
ConexionBD1 certificados, "select * from matricula where numdoc='" & numdoc.Text & "'"
If BD1.Recordset.RecordCount > 0 Then
    nivelcarne.Text = BD1.Recordset!graasp
End If
End Sub

Private Sub numdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
ConexionBD1 certificados, "select * from listadodeespera where numdoc='" & numdoc.Text & "'"
If BD1.Recordset.RecordCount > 0 Then
    nombrecarne.Text = BD1.Recordset!prinom & " " & BD1.Recordset!segnom
    apecarne.Text = BD1.Recordset!priape & " " & BD1.Recordset!segape
End If
ConexionBD1 certificados, "select * from matricula where numdoc='" & numdoc.Text & "'"
If BD1.Recordset.RecordCount > 0 Then
    nivelcarne.Text = BD1.Recordset!graasp
End If
End If
End Sub

Private Sub numdocc_Click()
ConexionBD certificados, "select * from listadodeespera where numdoc='" & numdocc.Text & "'"
If bd.Recordset.RecordCount > 0 Then
    certinombre.Text = bd.Recordset!prinom & " " & bd.Recordset!segnom & " " & bd.Recordset!priape & " " & bd.Recordset!segape
End If
ConexionBD certificados, "select * from matricula where numdoc='" & numdocc.Text & "'"
If bd.Recordset.RecordCount > 0 Then
    certigrado.Text = bd.Recordset!graasp
End If
End Sub



Private Sub numdocd_Click()
ConexionBD certificados, "select * from listadodeespera where numdoc='" & numdocd.Text & "'"
If bd.Recordset.RecordCount > 0 Then
    nombrediplo.Text = bd.Recordset!prinom & " " & bd.Recordset!segnom
    apediplo.Text = bd.Recordset!priape & " " & bd.Recordset!segape
    fecdiplo.SetFocus
End If
End Sub

Private Sub profecarne_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    vistacarne_Click
End If
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
If i = 1 Then
    aviso.ForeColor = &H80FF&    'naranja
ElseIf i = 2 Then
    aviso.ForeColor = &HFF& 'rojo
ElseIf i = 3 Then
    aviso.ForeColor = &H0& 'negro
End If
i = i + 1
If i > 3 Then
    i = 1
End If
End Sub

Private Sub vistacarne_Click()
If nombrecarne.Text = "" Or apecarne.Text = "" Or nivelcarne.Text = "" Or profecarne.Text = "" Then
    MsgBox "Falta datos por ingresar!", vbInformation, "Diplomas"
    Exit Sub
End If
On Error Resume Next
Set CCarnet = New Word.Application
'CCarnet.Application.Visible = True
CCarnet.Documents.Open App.Path & "\carnet.doc"
For i = 1 To 4
    CCarnet.WordBasic.Insert Chr(13)
Next i
CCarnet.Selection.Font.Name = "Comic Sans MS"
CCarnet.Selection.Font.Size = 12
CCarnet.WordBasic.Insert "            " & UCase(nombrecarne.Text)
CCarnet.WordBasic.Insert Chr(13)
CCarnet.WordBasic.Insert Chr(13)
CCarnet.Selection.Font.Name = "Comic Sans MS"
CCarnet.Selection.Font.Size = 12
CCarnet.WordBasic.Insert "      " & UCase(apecarne.Text)
CCarnet.WordBasic.Insert Chr(13)
CCarnet.Selection.Font.Name = "Comic Sans MS"
CCarnet.Selection.Font.Size = 12
CCarnet.WordBasic.Insert "          " & UCase(nivelcarne.Text)
CCarnet.WordBasic.Insert Chr(13)
CCarnet.Selection.Font.Name = "Comic Sans MS"
CCarnet.Selection.Font.Size = 12
CCarnet.WordBasic.Insert "              " & UCase(profecarne.Text)
CCarnet.ActiveDocument.PrintOut
DoEvents
CCarnet.ActiveDocument.Close wdDoNotSaveChanges
If Err.Number <> 0 Then
        MsgBox "Su versión de Microsoft Word no permite " & vbCrLf & "realizar esta acción!", vbInformation, "Constancias, Diplomas, Certificados..."
        Exit Sub
    End If
End Sub

Private Sub vistadiplo_Click()
If numdocd.ListIndex <> -1 Then
If nombrediplo.Text = "" Or apediplo.Text = "" Then
    MsgBox "Falta datos por ingresar!", vbInformation, "Diplomas"
    Exit Sub
End If
If MsgBox("Desea imprimir?", vbYesNo + vbQuestion, "Diplomas") = vbYes Then
On Error Resume Next
Set CDiploma = New Word.Application
CDiploma.Documents.Open App.Path & "\diplo.doc"
CDiploma.Selection.Paragraphs.Alignment = wdAlignParagraphCenter
For i = 1 To 22
    CDiploma.WordBasic.Insert Chr(13)
Next i
CDiploma.Selection.Font.Name = "Comic Sans MS"
CDiploma.Selection.Font.Size = 18
CDiploma.WordBasic.Insert "OTORGADO A:"
CDiploma.WordBasic.Insert Chr(13)
CDiploma.WordBasic.Insert Chr(13)
CDiploma.Selection.Font.Name = "Whimsy TT"
CDiploma.Selection.Font.Size = 30
CDiploma.WordBasic.Insert UCase(nombrediplo.Text)
CDiploma.WordBasic.Insert Chr(13)
CDiploma.Selection.Font.Name = "Whimsy TT"
CDiploma.Selection.Font.Size = 30
CDiploma.WordBasic.Insert UCase(apediplo.Text)
CDiploma.Selection.Font.Name = "Comic Sans MS"
CDiploma.Selection.Font.Size = 18
CDiploma.WordBasic.Insert Chr(13)
CDiploma.WordBasic.Insert Chr(13)
CDiploma.WordBasic.Insert Chr(13)
CDiploma.Selection.Font.Name = "Comic Sans MS"
CDiploma.Selection.Font.Size = 18
CDiploma.WordBasic.Insert "POR HABER JUGADO Y DISFRUTADO SU"
CDiploma.WordBasic.Insert Chr(13)
CDiploma.Selection.Font.Name = "Comic Sans MS"
CDiploma.Selection.Font.Size = 18
CDiploma.WordBasic.Insert "KINDER"
CDiploma.WordBasic.Insert Chr(13)
CDiploma.WordBasic.Insert Chr(13)
CDiploma.WordBasic.Insert Chr(13)
CDiploma.Selection.Font.Name = "Comic Sans MS"
CDiploma.Selection.Font.Size = 18
CDiploma.WordBasic.Insert "_________________            _________________"
CDiploma.WordBasic.Insert "    DIRECTORA                     EDUCADORA   "
CDiploma.WordBasic.Insert Chr(13)
CDiploma.WordBasic.Insert Chr(13)
CDiploma.WordBasic.Insert Chr(13)
CDiploma.WordBasic.Insert Chr(13)
CDiploma.Selection.Font.Name = "Comic Sans MS"
CDiploma.Selection.Font.Size = 10
CDiploma.WordBasic.Insert "BOGOTA D.C. " & UCase(mes(fecdiplo.Month)) & " " & fecdiplo.Day & " DE " & fecdiplo.Year
CDiploma.ActiveDocument.PrintOut
CDiploma.ActiveDocument.Close wdDoNotSaveChanges
CDiploma.Application.Quit
'CDiploma.ActiveDocument.PrintPreview
'CDiploma.Application.Visible = True
DoEvents
If Err.Number <> 0 Then
        MsgBox "Su versión de Microsoft Word no permite " & vbCrLf & "realizar esta acción!", vbInformation, "Constancias, Diplomas, Certificados..."
        Exit Sub
    End If
End If
End If
End Sub

Private Sub vistaprevia_Click()
If numdocc.ListIndex <> -1 Then
    If certinombre.Text = "" Or certigrado.Text = "" Or _
    certihora1.Text = "" Or certihora2.Text = "" Or _
    certinomex.Text = "" Or certicargo.Text = "" Then
        MsgBox "Falta datos por ingresar!", vbInformation, "Certificados"
        Exit Sub
    End If
    If MsgBox("Desea imprimir?", vbYesNo + vbQuestion, "Certificados") = vbYes Then
    On Error Resume Next
    Set Certificado = New Word.Application
       Certificado.Documents.Open App.Path & "\certi.doc"
    Certificado.Selection.Paragraphs.Alignment = wdAlignParagraphCenter
    Certificado.WordBasic.Bold
    Certificado.Selection.Font.Size = 16
    Certificado.Selection.Font.Name = "Verdana"
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert "CERTIFICO QUE:"
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.Selection.Paragraphs.Alignment = wdAlignParagraphJustify
    Certificado.Selection.Font.Size = 12
    Certificado.WordBasic.Insert "El niño " & UCase(certinombre.Text) & ", se encuentra matriculado en esta Institución en el grado " & certigrado.Text & ", con un horario de " & certihora1.Text & " a " & certihora2.Text & " de lunes a viernes."
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert "Se expide en Bogotá D.C. a los " & certifecha.Day & " días del mes de " & mes(certifecha.Month) & " del año " & certifecha.Year & ", a solicitud verbal del interesado."
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert "Coordialmente,"
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert "__________________________________"
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert certinomex.Text
    Certificado.WordBasic.Insert vbCrLf
    Certificado.WordBasic.Insert certicargo.Text
    Certificado.ActiveDocument.PrintOut
    Certificado.ActiveDocument.Close wdDoNotSaveChanges
    Certificado.Application.Quit
    'Certificado.ActiveDocument.PrintPreview
    'Certificado.Application.Visible = True
    DoEvents
    If Err.Number <> 0 Then
        MsgBox "Su versión de Microsoft Word no permite " & vbCrLf & "realizar esta acción!", vbInformation, "Constancias, Diplomas, Certificados..."
        Exit Sub
    End If
    End If
End If
End Sub


Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "certificados.htm"
End Sub
