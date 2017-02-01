Attribute VB_Name = "Declaraciones"
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
'PERMITE COLOCAR IMAGENES EN LOS MENUS
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Const MF_BYPOSITION = &H400&


Public usuario As String
Public passw As String
Public perfil As Integer
Public CambiarNomUsuario As String
Public ParametroBusqueda As String
Public FormularioActivo As Boolean
Public BackupR As Boolean
Public i As Integer
Public j As Integer
'variable paera la ayuda en linea
Public chmHelp As New cHtmlHelp
Public Tabla As String
Public ConsultaF As Boolean
Public EsNumero As Boolean
Public Matriculado As Boolean
Public Cualquiera As Integer
Public NuevoRegL As Boolean
Public NuevoRegI As Boolean
Public NuevoRegM As Boolean
Public NuevoRegE As Boolean
Public nd As String
Public PagoActual As Integer
Public ModificadoE As Boolean
Public ModificadoL As Boolean
Public ModificadoI As Boolean
Public ModificadoM As Boolean
Public ModificadoMat As Boolean
Public ModificadoP As Boolean
Public EmpleadoVal As Boolean
'para reporte de grados
Public Enum CantMax
    SalaCuna = 20
    Caminadores = 50
    Parvulos = 30
    Prekinder1 = 30
    Prekinder2 = 35
    Kinder = 35
End Enum
'para el motor de busqueda
Public Type Motor
    Formulario As String
    Descripcion As String
End Type
Public MB As Motor

Public Control As String
Public descargar As Boolean

Public Para As Boolean

Public ObjetoActual

Public carnete As Boolean
Public perfilito As Integer
