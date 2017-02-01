VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form CARGA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00BDF5FD&
   BorderStyle     =   0  'None
   Caption         =   "Form9"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc bd 
      Height          =   330
      Left            =   3120
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox picpgb2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   480
      ScaleHeight     =   19
      ScaleMode       =   0  'User
      ScaleWidth      =   431
      TabIndex        =   2
      Top             =   4200
      Width           =   6465
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   240
      Top             =   4560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows 2000, XP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5400
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Programa para el control de inscripciones, matriculas, pagos e inventario de material didáctico de un Jardín Infantil."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   3
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Left            =   120
      Picture         =   "CARGA.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2655
   End
   Begin VB.Image imgpgb1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      Picture         =   "CARGA.frx":E1552
      Top             =   4560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jardín Artístico Comunitario La Escuelita"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   1
      Left            =   195
      TabIndex        =   1
      Top             =   45
      Width           =   6870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jardín Artístico Comunitario La Escuelita"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   450
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6870
   End
End
Attribute VB_Name = "CARGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim distance As Integer

Private Sub bd_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub Form_Activate()
On Error Resume Next
ConexionBD CARGA, "select * from acceso"
Me.MousePointer = 11
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
    distance = 4
    Horizontal Me, RGB(253, 245, 189), RGB(249, 221, 147)
    picpgb2.PaintPicture imgpgb1, 0, 0, 4, 19, 0, 0, 4, 19
    picpgb2.PaintPicture imgpgb1, 4, 0, picpgb2.Width - 9, 19, 4, 0, 10, 19
    picpgb2.PaintPicture imgpgb1, picpgb2.Width - 5, 0, 5, 19, 14, 0, 5, 19
End Sub
Private Sub Form_Terminate()
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    For i = 1 To 2
        picpgb2.PaintPicture imgpgb1.Picture, distance, 4, 8, 12, 23, 5, 8, 12
        distance = distance + 10
    Next i
    If distance > picpgb2.Width - 5 Then
        Timer1.Enabled = False
        Unload Me
        FONDO.Show
    End If
End Sub


