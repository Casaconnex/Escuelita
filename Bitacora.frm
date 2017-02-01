VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Bitacora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visor de sucesos"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Bitacora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   477
   StartUpPosition =   2  'CenterScreen
   Begin Jardin.xphelp xphelp1 
      Height          =   315
      Left            =   6480
      Top             =   120
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin JeweledBut.JeweledButton Eliminar 
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   5520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      TX              =   "Eliminar Registros"
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
      MICON           =   "Bitacora.frx":08CA
      BC              =   8438015
      FC              =   0
      Picture         =   "Bitacora.frx":0A38
   End
   Begin JeweledBut.JeweledButton salir 
      Height          =   375
      Left            =   5520
      TabIndex        =   1
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
      MICON           =   "Bitacora.frx":0FD2
      BC              =   8438015
      FC              =   0
      Picture         =   "Bitacora.frx":1140
   End
   Begin MSAdodcLib.Adodc BD 
      Height          =   330
      Left            =   1200
      Top             =   5400
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
   Begin MSDataGridLib.DataGrid GRID 
      Bindings        =   "Bitacora.frx":129A
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8070
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "INGRESOS AL SISTEMA"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1514,986
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1514,986
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esta es una vista de los ingresos al sistema realizados por cada usuario."
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   6225
   End
End
Attribute VB_Name = "Bitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bd_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub eliminar_Click()
If MsgBox("Esta seguro(a) de realizar esta operación?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
    If BD.Recordset.RecordCount > 0 Then
        BD.Recordset.MoveFirst
        For i = 1 To BD.Recordset.RecordCount
            On Error Resume Next
            BD.Recordset.Delete
            BD.Recordset.MoveNext
        Next i
        BD.Refresh
        Unload Me
    End If
End If
If Err.Number <> 0 Then
    MsgBox Err.Description, vbCritical, "Error"
End If
End Sub

Private Sub Form_Activate()
FormularioActivo = True
ConexionBD Bitacora, "SELECT * FROM LOG"
GRID.Columns(0).Caption = "Usuario"
GRID.Columns(1).Caption = "Fecha"
GRID.Columns(2).Caption = "Hora"
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
menu.estado.Panels(4).Text = "Ingresos al sistema"
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormularioActivo = False
menu.estado.Panels(4).Text = "Menú Principal"
BD.Recordset.Close
Set BD.Recordset = Nothing
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "visor.htm"
End Sub
