VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form acercade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de Jardín"
   ClientHeight    =   3690
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5775
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "acercade.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2546.904
   ScaleMode       =   0  'User
   ScaleWidth      =   5423.023
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton cmdSysInfo 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "&Información"
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
      MICON           =   "acercade.frx":0E42
      BC              =   8438015
      FC              =   0
   End
   Begin JeweledBut.JeweledButton cmdok 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "&Aceptar"
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
      MICON           =   "acercade.frx":0E5E
      BC              =   8438015
      FC              =   0
      Picture         =   "acercade.frx":0E7A
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   120
      Picture         =   "acercade.frx":0FD4
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows 2000, XP"
      Height          =   195
      Left            =   4080
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"acercade.frx":E2526
      Height          =   1215
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jardín Artístico Comunitario La Escuelita"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión 1.0"
      Height          =   195
      Left            =   4680
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"acercade.frx":E25FC
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   240
      TabIndex        =   0
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "acercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Opciones de seguridad de clave del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos ROOT de clave del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en valor nulo
Const REG_DWORD = 4                      ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Intentar obtener ruta de acceso y nombre del programa de Info. del sistema a partir del Registro...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Intentar obtener sólo ruta del programa de Info. del sistema a partir del Registro...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validar la existencia de versión conocida de 32 bits del archivo
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error: no se puede encontrar el archivo...
        Else
            GoTo SysInfoErr
        End If
    ' Error: no se puede encontrar la entrada del Registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "La información del sistema no está disponible en este momento", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contador de bucle
    Dim rc As Long                                          ' Código de retorno
    Dim hKey As Long                                        ' Controlador de una clave de Registro abierta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo de datos de una clave de Registro
    Dim tmpVal As String                                    ' Almacenamiento temporal para un valor de clave de Registro
    Dim KeyValSize As Long                                  ' Tamaño de variable de clave de Registro
    '------------------------------------------------------------
    ' Abrir clave de registro bajo KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abrir clave de Registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Error de controlador...
    
    tmpVal = String$(1024, 0)                             ' Asignar espacio de variable
    KeyValSize = 1024                                       ' Marcar tamaño de variable
    
    '------------------------------------------------------------
    ' Obtener valor de clave de Registro...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtener o crear valor de clave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar errores
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agregar cadena terminada en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Encontrado valor nulo, se va a quitar de la cadena
    Else                                                    ' En WinNT las cadenas no terminan en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize)                   ' No se ha encontrado valor nulo, sólo se va a extraer la cadena
    End If
    '------------------------------------------------------------
    ' Determinar tipo de valor de clave para conversión...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Buscar tipos de datos...
    Case REG_SZ                                             ' Tipo de datos String de clave de Registro
        KeyVal = tmpVal                                     ' Copiar valor de cadena
    Case REG_DWORD                                          ' Tipo de datos Double Word de clave del Registro
        For i = Len(tmpVal) To 1 Step -1                    ' Convertir cada bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Generar valor carácter a carácter
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word a cadena
    End Select
    
    GetKeyValue = True                                      ' Se ha devuelto correctamente
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
    Exit Function                                           ' Salir
    
GetKeyError:      ' Borrar después de que se produzca un error...
    KeyVal = ""                                             ' Establecer valor a cadena vacía
    GetKeyValue = False                                     ' Fallo de retorno
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
End Function

Private Sub Form_Activate()
FormularioActivo = True
menu.Enabled = False
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
menu.estado.Panels(4).Text = "Información acerca del programa"
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormularioActivo = False
menu.Enabled = True
menu.estado.Panels(4).Text = "Menú Principal"
End Sub

