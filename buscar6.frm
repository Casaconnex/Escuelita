VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form buscarmateriald 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "elBuscador de SISJACE"
   ClientHeight    =   2670
   ClientLeft      =   2730
   ClientTop       =   1785
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   178
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Jardin.xp_canvas forma 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
      Caption         =   "elBuscador de SISJACE"
      Icon            =   "buscar6.frx":0000
      Fixed_Single    =   -1  'True
      Begin MSAdodcLib.Adodc bd 
         Height          =   330
         Left            =   1560
         Top             =   480
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
      Begin VB.TextBox termino 
         BackColor       =   &H00C5FAFE&
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   240
         MaxLength       =   40
         TabIndex        =   1
         Top             =   2160
         Width           =   2415
      End
      Begin Jardin.xptopbuttons cerrar 
         Height          =   315
         Left            =   5505
         Top             =   75
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin Jardin.xphelp ayuda 
         Height          =   315
         Left            =   5175
         Top             =   75
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   3600
         Top             =   480
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Digite lo que quiere buscar:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   2385
      End
      Begin VB.Image Image2 
         Height          =   630
         Left            =   3240
         MouseIcon       =   "buscar6.frx":0452
         MousePointer    =   99  'Custom
         Picture         =   "buscar6.frx":05B0
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   675
      End
      Begin VB.Image Image1 
         Height          =   915
         Left            =   240
         Picture         =   "buscar6.frx":0B71
         Stretch         =   -1  'True
         Top             =   480
         Width           =   3615
      End
      Begin VB.Image Image3 
         Height          =   1845
         Left            =   3840
         Picture         =   "buscar6.frx":1AE7
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image Image4 
         Height          =   630
         Left            =   3240
         MouseIcon       =   "buscar6.frx":4971
         MousePointer    =   99  'Custom
         Picture         =   "buscar6.frx":4ACF
         Stretch         =   -1  'True
         Top             =   1800
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image Image5 
         Height          =   630
         Left            =   3240
         MouseIcon       =   "buscar6.frx":5B41
         MousePointer    =   99  'Custom
         Picture         =   "buscar6.frx":5C9F
         Stretch         =   -1  'True
         Top             =   1800
         Visible         =   0   'False
         Width           =   660
      End
   End
End
Attribute VB_Name = "buscarmateriald"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dato As String, Consulta_SQL As String



Private Sub ayuda_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "buscar.htm"
End Sub

Private Sub fecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Image2_Click
End If
End Sub

Private Sub List1_Click()
Select Case List1.ListIndex
    Case 0: ParametroBusqueda = "numdes"
            termino.Visible = True
            fecha.Visible = False
            termino.SetFocus
            EsNumero = True
    Case 1: ParametroBusqueda = "fecdes"
            termino.Visible = False
            fecha.Visible = True
            fecha.SetFocus
            EsNumero = False
    Case 2: ParametroBusqueda = "ref"
            termino.Visible = True
            fecha.Visible = False
            termino.SetFocus
            EsNumero = True
    Case 3: ParametroBusqueda = "numdoc"
            termino.Visible = True
            fecha.Visible = False
            termino.SetFocus
            EsNumero = True
    Case 4: ParametroBusqueda = "sede"
            termino.Visible = True
            fecha.Visible = False
            termino.SetFocus
            EsNumero = False
End Select
End Sub

Private Sub termino_KeyPress(KeyAscii As Integer)

    KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    Image2_Click
End If
End Sub

Private Sub Timer1_Timer()
Image5_Click
Timer1.Enabled = False

End Sub
Private Sub Image2_Click()
Dim campo As String
Dim enc As Integer
enc = 0

If termino.Text = "" Then
    MsgBox "No ha ingresado el termino a buscar!", vbExclamation, "elBuscador"
    Exit Sub
End If
'consulta la bd
campo = ParametroBusqueda
dato = termino.Text


Consulta_SQL = "SELECT * FROM materialdespachado where " & campo & "=" & dato
bd.RecordSource = Consulta_SQL
bd.Refresh
If bd.Recordset.RecordCount > 1 Then
    mdespachado.anterior.Visible = True
    mdespachado.siguiente.Visible = True
    mostrarcampos
    enc = 1
ElseIf bd.Recordset.RecordCount > 0 Then
    mostrarcampos
    enc = 1
    Me.Hide
End If

If enc = 0 Then
  MsgBox "No se encontraron coincidencias", vbInformation, "elBuscador"
  termino.SetFocus
End If
End Sub
Private Sub cerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
ConexionBD buscarmateriald, "select * from materialdespachado"

End Sub
Private Sub forma_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Image4.Visible = False
Image5.Visible = False
End Sub
Function mostrarcampos()
mdespachado.numdes.Text = bd.Recordset!numdes
mdespachado.fecdes.Value = bd.Recordset!fecdes
mdespachado.ref.Text = bd.Recordset!ref
mdespachado.can.Text = bd.Recordset!can
'mdespachado.desper.Text = bd.Recordset!desper
mdespachado.numdoc.Text = bd.Recordset!numdoc
mdespachado.sede.Text = bd.Recordset!sede
End Function

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Image4.Visible = True
End Sub

Private Sub Image4_Click()
Image4.Visible = False
Image5.Visible = True

End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = True
End Sub

Private Sub Image5_Click()
Image2_Click
End Sub
Private Sub termino_GotFocus()
termino.SelStart = 0
termino.SelLength = Len(termino.Text)
End Sub


