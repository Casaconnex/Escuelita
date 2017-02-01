VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox forpagviv 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   3120
      List            =   "Form1.frx":0010
      TabIndex        =   12
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox munprofam 
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox paiprofam 
      Height          =   285
      Left            =   3120
      TabIndex        =   10
      Text            =   "Colombia"
      Top             =   6960
      Width           =   2175
   End
   Begin VB.ComboBox depprofam 
      Height          =   315
      ItemData        =   "Form1.frx":0039
      Left            =   3120
      List            =   "Form1.frx":0097
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   6240
      Width           =   2415
   End
   Begin VB.ComboBox calben 
      Height          =   315
      ItemData        =   "Form1.frx":01B9
      Left            =   2160
      List            =   "Form1.frx":01C6
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox regsegsociben 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox afisegsocben 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox agrviointfam 
      Height          =   315
      ItemData        =   "Form1.frx":01F8
      Left            =   3120
      List            =   "Form1.frx":0202
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin MSComCtl2.DTPicker fecllebogfam 
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   49020929
      CurrentDate     =   38117
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "País Procedencia Familia"
      Height          =   375
      Left            =   840
      TabIndex        =   17
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Municipio Procedencia Familia"
      Height          =   375
      Left            =   840
      TabIndex        =   16
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha llegada Bogotá Familia"
      Height          =   495
      Left            =   840
      TabIndex        =   15
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Departamento Procedencia Familia"
      Height          =   495
      Left            =   840
      TabIndex        =   14
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de Pago Vivienda"
      Height          =   195
      Left            =   840
      TabIndex        =   13
      Top             =   3840
      Width           =   2085
   End
   Begin VB.Label Label64 
      BackStyle       =   0  'Transparent
      Caption         =   "Calidad del beneficiario"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label63 
      BackStyle       =   0  'Transparent
      Caption         =   "Regimen seguridad social beneficiario"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label62 
      BackStyle       =   0  'Transparent
      Caption         =   "Afiliado a la seguridad social beneficiario"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "Agresor violencia Intrafimiar"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
