VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form backup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Backup de la Base de Datos"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   140
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox destino 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   360
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Guardar Copia de Seguridad..."
      FileName        =   "Proyecto"
      InitDir         =   "C:\"
   End
   Begin JeweledBut.JeweledButton cancelar 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      TX              =   "&Cancelar"
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
      MPTR            =   0
      MICON           =   "backup.frx":0000
      BC              =   8438015
      FC              =   0
      Picture         =   "backup.frx":001C
   End
   Begin JeweledBut.JeweledButton copia 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      TX              =   "&Realizar Copia"
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
      MPTR            =   0
      MICON           =   "backup.frx":0176
      BC              =   8438015
      FC              =   0
   End
   Begin JeweledBut.JeweledButton buscar 
      Height          =   420
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "De click aquí para escoger el destino..."
      Top             =   720
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   741
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
      MPTR            =   0
      MICON           =   "backup.frx":0192
      BC              =   12632256
      FC              =   0
      Picture         =   "backup.frx":01AE
   End
   Begin Jardin.xphelp xphelp1 
      Height          =   315
      Left            =   2400
      Top             =   1560
      Width           =   315
      _extentx        =   556
      _extenty        =   556
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Se va a realizar una copia de seguridad de la Base de  Datos. "
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   4710
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destino:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   720
   End
End
Attribute VB_Name = "backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buscar_Click()
Dialogo.Filter = "Base de Datos |*.mdb"
If BackupR = True Then
    Dialogo.ShowSave
ElseIf BackupR = False Then
    Dialogo.ShowOpen
End If
destino.Text = Dialogo.FileName
If destino.Text = "Proyecto" Then
    destino.Text = ""
End If
End Sub

Private Sub cancelar_Click()
Unload Me
End Sub

Private Sub copia_Click()
If destino.Text = "" Then
    MsgBox "La ruta no es valida!", vbExclamation, Me.Caption
    Exit Sub
End If

On Error GoTo error_copia
If BackupR = True Then
    FileCopy App.Path + "\proyecto.mdb", destino.Text
    MsgBox "Copia de seguridad realizada exitosamente!", vbInformation, "Backup"
    Unload Me
ElseIf BackupR = False Then
    If MsgBox("Al realizar la restauración perderá los datos ingresados despues" & vbCrLf & "de haber realizado la última copia de seguridad." & vbCrLf & vbCrLf & "Está seguro(a)?", vbQuestion + vbYesNo, "Advertencia!") = vbYes Then
        FileCopy destino.Text, App.Path + "\proyecto.mdb"
        MsgBox "Restauración realizada exitosamente!", vbInformation, "Backup"
        Unload Me
    Else
        Exit Sub
    End If
End If

copias:
    Exit Sub
error_copia:
    MsgBox Err.Description, vbCritical, "Error Backup"
    Resume copias
End Sub

Private Sub Form_Activate()
FormularioActivo = True
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
menu.estado.Panels(4).Text = "Generación de Backup y/o restauración de la Base de Datos"
menu.Enabled = False
If BackupR = True Then
    dos = "Destino:"
    copia.Caption = "Realizar Backup"
    Dialogo.DialogTitle = "Guardar Copia de Seguridad..."
    backup.Caption = "Backup de la Base de Datos"
ElseIf BackupR = False Then
    dos = "Origen:"
    copia.Caption = "Restaurar Backup"
    Dialogo.DialogTitle = "Restaurar Copia de Seguridad..."
    backup.Caption = "Restauración de la Base de Datos"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormularioActivo = False
menu.Enabled = True
menu.estado.Panels(4).Text = "Menú Principal"
End Sub

Private Sub xphelp1_Click()
chmHelp.HelpFile = App.Path + "\jardin.chm"
chmHelp.DisplayTopic "copiaseguridad.htm"
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
