VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form buscarmatricula 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "elBuscador de SISJACE"
   ClientHeight    =   3735
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
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   ShowInTaskbar   =   0   'False
   Begin Jardin.xp_canvas forma 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6588
      Caption         =   "elBuscador de SISJACE"
      Icon            =   "buscar3.frx":0000
      Fixed_Single    =   -1  'True
      Begin MSAdodcLib.Adodc bd1 
         Height          =   330
         Left            =   3000
         Top             =   600
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
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   5
         Top             =   3240
         Width           =   3015
      End
      Begin Jardin.xpgroupbox xpgroupbox1 
         Height          =   1455
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2566
         Caption         =   "Buscar por:"
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
         Begin VB.ListBox List1 
            BackColor       =   &H00C5FAFE&
            Height          =   1035
            ItemData        =   "buscar3.frx":0452
            Left            =   120
            List            =   "buscar3.frx":045C
            TabIndex        =   4
            Top             =   240
            Width           =   2415
         End
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
      Begin JeweledBut.JeweledButton siguiente 
         Height          =   495
         Left            =   960
         TabIndex        =   2
         ToolTipText     =   "Siguiente Registro"
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         TX              =   ""
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
         MICON           =   "buscar3.frx":0481
         BC              =   8438015
         FC              =   0
         Picture         =   "buscar3.frx":05EF
      End
      Begin JeweledBut.JeweledButton anterior 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Anterior Registro"
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         TX              =   ""
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
         MICON           =   "buscar3.frx":0A41
         BC              =   8438015
         FC              =   0
         Picture         =   "buscar3.frx":0BAF
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   5400
         Top             =   1680
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Digite lo que quiere buscar:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1800
         TabIndex        =   7
         Top             =   3000
         Width           =   2385
      End
      Begin VB.Image Image2 
         Height          =   630
         Left            =   4920
         MouseIcon       =   "buscar3.frx":1001
         MousePointer    =   99  'Custom
         Picture         =   "buscar3.frx":115F
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   675
      End
      Begin VB.Image Image1 
         Height          =   915
         Left            =   240
         Picture         =   "buscar3.frx":1720
         Stretch         =   -1  'True
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C5FAFE&
         Caption         =   "Seleccione en la lista el parámetro de búsqueda que desee..."
         Height          =   855
         Left            =   3120
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Image Image3 
         Height          =   1845
         Left            =   3840
         Picture         =   "buscar3.frx":2696
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image Image4 
         Height          =   630
         Left            =   4920
         MouseIcon       =   "buscar3.frx":5520
         MousePointer    =   99  'Custom
         Picture         =   "buscar3.frx":567E
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image Image5 
         Height          =   630
         Left            =   4920
         MouseIcon       =   "buscar3.frx":66F0
         MousePointer    =   99  'Custom
         Picture         =   "buscar3.frx":684E
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C5FAFE&
         BackStyle       =   1  'Opaque
         Height          =   1095
         Left            =   3000
         Top             =   1560
         Width           =   1815
      End
   End
End
Attribute VB_Name = "buscarmatricula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim dato As String, Consulta_SQL As String
Private Sub anterior_Click()
bd.Recordset.MovePrevious
If bd.Recordset.BOF = True Then
      bd.Recordset.MoveNext
      MsgBox "Este el primer registro!", vbInformation, "elBuscador"
    Else
     mostrarcampos
End If
End Sub

Private Sub ayuda_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "buscar.htm"
End Sub

Private Sub cerrar_Click()
Unload Me
End Sub

Private Sub fec_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Image2_Click
End If
End Sub

Private Sub Form_Load()
ConexionBD buscarmatricula, "select * from matricula"
End Sub

Private Sub forma_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Image4.Visible = False
Image5.Visible = False
End Sub

Private Sub Image2_Click()
Dim campo As String
Dim enc As Integer
enc = 0

If fec.Visible = True Then
    termino.Text = fec.Value
End If

If List1.ListIndex = -1 Then
    MsgBox "No ha seleccionado un parámetro de búsqueda!", vbExclamation, "elBuscador"
    Exit Sub
ElseIf termino.Text = "" Then
    MsgBox "No ha ingresado el termino a buscar!", vbExclamation, "elBuscador"
    Exit Sub
End If
'consulta la bd
campo = ParametroBusqueda
dato = termino.Text

If campo = "fecmat" Then
    Consulta_SQL = "SELECT * FROM matricula where " & campo & "=" & dato '+ "';"
ElseIf campo = "numdoc" Then
    Consulta_SQL = "SELECT * FROM matricula where " & campo & "='" & dato + "';"
ElseIf campo = "numfor" Then
    Consulta_SQL = "SELECT * FROM matricula where " & campo & "=" & dato '+ "';"
End If
 

bd.RecordSource = Consulta_SQL
bd.Refresh
If bd.Recordset.RecordCount > 1 Then
    anterior.Visible = True
    siguiente.Visible = True
    mostrarcampos
    enc = 1
    Me.Hide
ElseIf bd.Recordset.RecordCount > 0 Then
    mostrarcampos
    enc = 1
    Me.Hide
End If

If enc = 0 Then
  MsgBox "No se encontraron coincidencias", vbInformation, "elBuscador"
  If fec.Visible = True Then
    fec.SetFocus
  Else
    termino.SetFocus
  End If
End If
End Sub

Function mostrarcampos()
'muestra todos los campos de matricula
    matricula.numfor.Text = bd.Recordset!numfor
    matricula.col.Text = bd.Recordset!col
    matricula.uniope.Text = bd.Recordset!uniope
    matricula.modal.Text = bd.Recordset!modal
    matricula.submod.Text = bd.Recordset!submod
    matricula.fecmat.Value = bd.Recordset!fecmat
    matricula.persolser.Text = bd.Recordset!persolser
    matricula.rempor.Text = bd.Recordset!rempor
    matricula.prorem.Text = bd.Recordset!prorem
    matricula.entrem.Text = bd.Recordset!entrem
    matricula.numdoc.Text = bd.Recordset!numdoc
    matricula.depnac.Text = bd.Recordset!depnac
    matricula.munnac.Text = bd.Recordset!munnac
    matricula.Painac.Text = bd.Recordset!Painac
    matricula.tipdisest.Text = bd.Recordset!tipdisest
    matricula.nivestalc.Text = bd.Recordset!niveduben
    matricula.asiactcenedu.Text = bd.Recordset!asiactcenedu
    matricula.proaso.Text = bd.Recordset!proaso
    matricula.afisegsocfam.Text = bd.Recordset!afisegsocben
    matricula.regsegsocfam.Text = bd.Recordset!regsegsocben
    matricula.calbenfam.Text = bd.Recordset!calben
    matricula.vinsecsalfam.Text = bd.Recordset!vinsecsalben
    matricula.numficsis.Text = bd.Recordset!numficsis
    matricula.punsis.Text = bd.Recordset!punsis
    matricula.loc.Text = bd.Recordset!loc
    matricula.forpagviv.Text = bd.Recordset!forpagviv
    matricula.dep.Text = bd.Recordset!dptoprofam
    matricula.mun.Text = bd.Recordset!munprofam
    matricula.pais.Text = bd.Recordset!paiprofam
    matricula.feclle.Value = bd.Recordset!fecllebogfam
    matricula.ninvivpapmam.Text = bd.Recordset!ninvivpapmam
    matricula.ninvivperpadmadotr.Text = bd.Recordset!ninvivperpadmadotr
    matricula.vivpermpadmad.Text = bd.Recordset!vivperpadmad
    matricula.edaninvivpapmad.Text = bd.Recordset!edaninvivpapmad
    matricula.cuinindurdia.Text = bd.Recordset!cuinindurdia
    matricula.graasp.Text = bd.Recordset!graasp
    If IsNull(bd.Recordset!nomdilform) = False Then
        matricula.nomdilfor.Text = bd.Recordset!nomdilform
    End If
    matricula.nomfundighojsir.Text = bd.Recordset!nomfundighojsir
    matricula.fecdighojsir.Value = bd.Recordset!fecdighojsir
    matricula.obs.Text = bd.Recordset!obs
    'muestra los campos de listado de espera
    ConexionBD1 buscarmatricula, "select * from listadodeespera where numdoc='" & bd.Recordset!numdoc & "'"
    If bd1.Recordset.RecordCount > 0 Then
        matricula.tipdocnin.Text = bd1.Recordset!tipdoc
        matricula.nomnin.Text = bd1.Recordset!prinom & " " & bd1.Recordset!segnom
        matricula.apenin.Text = bd1.Recordset!priape & " " & bd1.Recordset!segape
        matricula.sexo.Text = bd1.Recordset!sex
        matricula.fecnac.Value = bd1.Recordset!fecnac
        matricula.edad.Text = bd1.Recordset!eda
        matricula.parjeffam.Text = bd1.Recordset!parfam
        matricula.dir.Text = bd1.Recordset!dir
        matricula.bar.Text = bd1.Recordset!bar
        matricula.tel.Text = bd1.Recordset!tel
    End If
        'conectamos bd para cargar datos de inscripciones
        ConexionBD1 buscarmatricula, "select * from inscripciones where numdoc='" & bd.Recordset!numdoc & "'"
        matricula.tipviv.Text = bd1.Recordset!tipviv
        matricula.conviv.Text = bd1.Recordset!conviv
        matricula.tenviv.Text = bd1.Recordset!tenviv
matricula.SSTab1.Tab = 0

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

Private Sub List1_Click()
Select Case List1.ListIndex
    Case 0: ParametroBusqueda = "numdoc"
            fec.Visible = False
            termino.Visible = True
            termino.Text = ""
            Label1.Visible = True
            termino.SetFocus
            Cualquiera = 3
            
    Case 1: ParametroBusqueda = "fecmat"
            fec.Visible = True
            termino.Visible = False
            Label1.Visible = False
            fec.SetFocus
            Cualquiera = 2
    Case 2: ParametroBusqueda = "numfor"
            fec.Visible = False
            termino.Visible = True
            termino.Text = ""
            Label1.Visible = True
            termino.SetFocus
            Cualquiera = 1
    
End Select
End Sub



Private Sub siguiente_Click()
bd.Recordset.MoveNext
If bd.Recordset.EOF = True Then
      bd.Recordset.MovePrevious
      MsgBox "Este es el último registro", vbInformation, "elBuscador"
    Else
    mostrarcampos
End If
End Sub

Private Sub termino_GotFocus()
termino.SelStart = 0
termino.SelLength = Len(termino.Text)
End Sub

Private Sub termino_KeyPress(KeyAscii As Integer)

If Cualquiera = 1 Then
    KeyAscii = Validar_numero(KeyAscii)
ElseIf Cualquiera = 2 Then
    KeyAscii = Validar_letra(KeyAscii)
ElseIf Cualquiera = 3 Then
    KeyAscii = KeyAscii
End If

If KeyAscii = 13 Then
    Image2_Click
End If
End Sub

Private Sub Timer1_Timer()
Image5_Click
Timer1.Enabled = False
End Sub

