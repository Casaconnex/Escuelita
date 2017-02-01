VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHtmlHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Asignar el valor -1 a la constante ES_DEBUG
'para mostrar los valores devueltos por la función HtmlHelp
#Const ES_DEBUG = -1

'Para usar con VB5 los mensajes emergentes
'-----------------------------------------
'Tomado del tutorial de HTML Help de http://www.vbexplorer.com/
Private Type HH_IDPAIR
    dwControlId As Long
    dwTopicId As Long
End Type

'This array should contain the number of controls that have
'context-sensitive help, plus one more for a zero-terminating
'pair.
Private ids(0 To 1) As HH_IDPAIR

Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hWnd As Long) As Long
'-----------------------------------------

'Propiedades públicas
Public hWnd As Long         'El hWnd del form que llama a la ayuda
Public HelpFile As String   'El fichero de ayuda con extensión chm

'----------------------------------------------------------------------------------
'En Microsoft TechNet puedes encontrar este artículo:
'HOWTO: Use HTML Help API in a Visual Basic 5.0 Application
'PSS ID Number: Q183434
'
'Aunque la definición de la Enumeración y la primera declaración
'es de las news
'----------------------------------------------------------------------------------
'
'Htmlhelp consts
Private Enum HH_COMMAND
    HH_DISPLAY_TOPIC = &H0
    HH_HELP_FINDER = &H0        ' WinHelp equivalent
    HH_DISPLAY_TOC = &H1        ' Muestra la tabla de contenidos
    HH_DISPLAY_INDEX = &H2      ' Muestra la solapa Index
    HH_DISPLAY_SEARCH = &H3     ' Muestra la solapa de búsqueda (Search)
    HH_SET_WIN_TYPE = &H4
    HH_GET_WIN_TYPE = &H5
    HH_GET_WIN_HANDLE = &H6
    HH_GET_INFO_TYPES = &H7     ' not currently implemented
    HH_SET_INFO_TYPES = &H8     ' not currently implemented
    HH_SYNC = &H9
    HH_ADD_NAV_UI = &HA         ' not currently implemented
    HH_ADD_BUTTON = &HB         ' not currently implemented
    HH_GETBROWSER_APP = &HC     ' not currently implemented
    HH_KEYWORD_LOOKUP = &HD
    HH_DISPLAY_TEXT_POPUP = &HE ' display string resource id
                                ' or text in a popup window
    HH_HELP_CONTEXT = &HF       ' display mapped numeric value
                                ' in dwData
    HH_TP_HELP_CONTEXTMENU      ' Text pop-up help, similar to
                                ' WinHelp's HELP_CONTEXTMENU.
                                '   HH_TP_HELP_CONTEXTMENU = &H10
    HH_TP_HELP_WM_HELP = &H11   ' text pop-up help, similar to
                                ' WinHelp's HELP_WM_HELP.
    HH_CLOSE_ALL = &H12         ' close all windows opened directly
                                ' or indirectly by the caller
    HH_ALINK_LOOKUP = &H13      ' ALink version of HH_KEYWORD_LOOKUP
End Enum

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type POINTAPI
    X   As Long
    Y   As Long
End Type
    
'
'Estructura para usar con el comando HH_DISPLAY_TEXT_POPUP
'Sacada de la página HTML Help API Reference del site de MS.
'
'typedef struct tagHH_POPUP         Member         Description
'{                                  ------            -----------
'    int       cbStruct;            cbStruct       Specifies the size of the structure. This value must always be filled in before passing the structure to the HtmlHelp API.
'    HINSTANCE hinst;               hinst          Instance handle of the program or .dll file to retrieve the string resource from. Ignored if idString is zero, or a filename was passed to the HtmlHelp API.
'    UINT      idString;            idString       Specifies either a resource ID or a topic number in a text file.
'    LPCTSTR   pszText;             pszText        Specifies the text to display if idString is zero.
'    POINT     pt;                  pt             Specifies where the top center of the pop-up should be located.
'    COLORREF  clrForeground;       clrForeground  Specifies the RGB value to use for the foreground color. If this value is -1, the system color for window text will be used.
'    COLORREF  clrBackground;       clrBackground  Specifies the RGB value to use for the background color. If this value is -1, the system color for window background will be used.
'    RECT      rcMargins;           rcMargins      Specifies the margins to use on the left, top, right, and bottom sides of the pop-up. The default for all rectangle members is -1.
'    LPCTSTR   pszFont;             pszFont        Specifies an optional text string containing: facename[,point size[,charset[,color[, BOLD ITALIC UNDERLINE]]]]. This string specifies a font, its point size, its character set, its color, and various attributes. Options can be omitted by simply specifying the next comma. For example, to create a bold, 10-pt MS Sans Serif font, the parameter value would be: MS Sans Serif, 10, , , BOLD
'} HH_POPUP;
'
Private Type tagHH_POPUP
    cbStruct        As Long     '=Len(HH_POPUP) Longitud de la estructura
    hinst           As Long     '
    idString        As Long     '
    pszText         As String   'Texto a mostrar
    pt              As POINTAPI 'Posición de la ventana popup, en pixels
    clrForeground   As Long     'Color de la letra, -1 es el valor por defecto
    clrBackground   As Long     'Color de fondo, -1 es el valor por defecto
    rcMargins       As RECT     'Tamaño de la ventana, -1 son los valores por defecto
    pszFont         As String   'Tipo de letra a usar: "Arial,9"
                                'facename[,point size[,charset[,color[, BOLD ITALIC UNDERLINE]]]].
End Type

'
'Estructura para mostrar la solapa Search y un texto a buscar
'sacada del fichero de ayuda api.chm de MS.
'
'typedef struct tagHH_FTS_QUERY     Member             Description
'{                                  ------              -----------
'     int      cbStruct;            cbStruct           Specifies the size of the structure.
'     BOOL     fUniCodeStrings;     fUniCodeStrings    TRUE if all strings are Unicode.
'     LPCTSTR  pszSearchQuery;      pszSearchQuery     String containing the search query.
'     LONG     iProximity;          iProximity         Word proximity.
'     BOOL     fStemmedSearch;      fStemmedSearch     TRUE for StemmedSearch only.
'     BOOL     fTitleOnly;          fTitleOnly         TRUE for Title search only.
'     BOOL     fExecute;            fExecute           TRUE to initiate the search.
'     LPCTSTR  pszWindow;           pszWindow          Window to display in.
'} HH_FTS_QUERY;
'
Private Type tagHH_FTS_QUERY
    cbStruct        As Long     '=Len(HH_FTS_QUERY) Longitud de la estructura.
    fUniCodeStrings As Long     'TRUE if all strings are Unicode.
    pszSearchQuery  As String   'String containing the search query.
    iProximity      As Long     'Word proximity.
    fStemmedSearch  As Long     'TRUE for StemmedSearch only.
    fTitleOnly      As Long     'TRUE for Title search only.
    fExecute        As Long     'TRUE to initiate the search.
    pszWindow       As String   'Window to display in.
End Type

'
'HtmlHelp api call
'Con esta funciona perfectamente, los LONGs
'Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
    (ByVal hwndCaller As Long, ByVal pszFile As String, _
    ByVal uCommand As HH_COMMAND, ByVal dwData As Long) As Long

'NOTA: Si se usa esta forma, hay que indicar el último parámetro
'      con la palabra ByVal delante...
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
    (ByVal hWndCaller As Long, ByVal pszFile As String, _
    ByVal uCommand As HH_COMMAND, dwData As Any) As Long


Public Function DisplayContext(Optional ByVal TopicNum As Long = 0&) As Long
    'Mostrar el tópico de ayuda,
    'si no se indica se mostrará el tópico por defecto
    '
    Dim sHelpFile As String
    
    sHelpFile = SoloCHM
    DisplayContext = HtmlHelp(hWnd, sHelpFile, HH_HELP_CONTEXT, ByVal TopicNum)
    
    #If ES_DEBUG Then
        Debug.Print "DisplayContext= " & DisplayContext
    #End If
End Function

Public Function DisplayTopic(Optional ByVal TopicName As String = "") As Long
    'Mostrar el tópico de ayuda,
    'si no se indica el valor de TopicName, se mostrará el tópico por defecto.
    'TopicName deberá usarse como suele hacerse con los links normales:
    '   nombrePágina.htm#NombreMarcador o nombrePagina.htm
    '
    Dim sHelpFile As String
        
    sHelpFile = SoloCHM
    If Len(TopicName) = 0 Then
        DisplayTopic = HtmlHelp(hWnd, sHelpFile, HH_DISPLAY_TOPIC, ByVal 0&)
    Else
        DisplayTopic = HtmlHelp(hWnd, sHelpFile, HH_DISPLAY_TOPIC, ByVal TopicName)
    End If
    
    #If ES_DEBUG Then
        Debug.Print "DisplayTopic= " & DisplayTopic
    #End If
End Function

Public Function PopUp(ByVal Text As String) As Long
    'Para mostrar una ventana PopUp con el texto indicado
    Dim HH_POPUP As tagHH_POPUP
    Dim elForm As Form
    
    On Local Error Resume Next
    
    'Asignar el form activo
    Set elForm = Screen.ActiveForm
    
    With HH_POPUP
        .cbStruct = Len(HH_POPUP)
        .clrBackground = &HA3F8FE
        .clrForeground = -1
        .pszFont = "Verdana,10"
        .pszText = Text
        'posicionar la ventana de PopUp
        .pt.X = (elForm.Left + 360) \ Screen.TwipsPerPixelX
        .pt.Y = (elForm.Top + (elForm.Height \ 2) + 240) \ Screen.TwipsPerPixelY
        .rcMargins.Bottom = -1
        .rcMargins.Left = -1
        .rcMargins.Right = -1
        .rcMargins.Top = -1
    End With
    PopUp = HtmlHelp(hWnd, HelpFile, HH_DISPLAY_TEXT_POPUP, HH_POPUP)
    
    #If ES_DEBUG Then
        Debug.Print "PopUp= " & PopUp
    #End If
    
    Err = 0
End Function

Public Function DisplaySearch() As Long
'                              Optional ByVal sBuscar As String = ""
'                              Optional ByVal laVentana As String = "") As Long
    'Muestra la solapa Search (Búsqueda)
    '
    Dim sHelpFile As String
    Dim HH_FTS_QUERY As tagHH_FTS_QUERY
    
    With HH_FTS_QUERY
        .cbStruct = Len(HH_FTS_QUERY)
        .fStemmedSearch = 0&
        .fTitleOnly = 0&
        .fUniCodeStrings = 0&
        .iProximity = 0&
        .pszSearchQuery = "" 'sBuscar   'sólo se mostrará si no se ha efectuado ninguna búsqueda
        .pszWindow = "Default"      'laVentana 'esto no tiene efecto
        .fExecute = 1&              'esto parece que no tiene efecto en VB
    End With
    
    sHelpFile = SoloCHM
    DisplaySearch = HtmlHelp(hWnd, sHelpFile, HH_DISPLAY_SEARCH, HH_FTS_QUERY)
    
    #If ES_DEBUG Then
        Debug.Print "DisplaySearch= " & DisplaySearch
    #End If
End Function

Public Function DisplayTOC() As Long
    'Muestra la tabla de contenidos (TOC)
    Dim sHelpFile As String
    
    sHelpFile = SoloCHM
    
        DisplayTOC = HtmlHelp(hWnd, sHelpFile, HH_DISPLAY_TOC, ByVal 0&)
    
    #If ES_DEBUG Then
        Debug.Print "DiplayToc= " & DisplayTOC
    #End If
End Function

Public Function DisplayIndex(Optional ByVal sKeyWord As String = "") As Long
    'Muestra la solapa Index,
    'si se especifica una palabra clave, se mostrará esa clave          (27/Dic/98)
    Dim sHelpFile As String
    
    sHelpFile = SoloCHM
    If Len(sKeyWord) = 0 Then
        DisplayIndex = HtmlHelp(hWnd, sHelpFile, HH_DISPLAY_INDEX, ByVal 0&)
    Else
        DisplayIndex = HtmlHelp(hWnd, sHelpFile, HH_DISPLAY_INDEX, ByVal sKeyWord)
    End If
    
    #If ES_DEBUG Then
        Debug.Print "DisplayIndex= " & DisplayIndex
    #End If
End Function

Private Function SoloCHM() As String
    'Devuelve sólo el fichero base de la ayuda
    
    'Hay que quitarle el ::/xxx que tenga
    Dim i As Long
    Dim sHelpFile As String
    
    sHelpFile = HelpFile
    i = InStr(sHelpFile, "::/")
    If i Then
        sHelpFile = Trim$(Left$(sHelpFile, i - 1))
    End If
    SoloCHM = sHelpFile
End Function

Public Function HelpContextPop(Optional ByVal elControl As Control) As Long
    'Se pasará el control en el que se ha pulsado F1
    'Deberá tener asignado el valor del HelpContextID
    'Esta función es para usar con VB5 o con un formulario en el que no se
    'ha especificado el WhatThisHelp...
    Dim vControl As Control
    
    On Local Error Resume Next
    
    If elControl Is Nothing Then
        Set vControl = Screen.ActiveControl
    Else
        Set vControl = elControl
    End If
    
    ids(0).dwTopicId = CLng(vControl.HelpContextID)
    ids(0).dwControlId = GetDlgCtrlID(vControl.hWnd)
    ' The last pair in the array must contain zeros (0)
    ids(1).dwControlId = 0
    ids(1).dwTopicId = 0
    
    If Err = 0 Then
        HelpContextPop = HtmlHelp(vControl.hWnd, HelpFile, HH_TP_HELP_WM_HELP, ids(0))
    End If
    
    Err = 0
End Function

Public Function HelpContext(Optional ByVal elControl As Control, _
                            Optional ByVal laVentana As String = "") As Long
    'elControl  será el control en el que se ha pulsado F1
    '           si no se especifica, se tomará el control activo.
    '           Ese control deberá tener asignado el valor del HelpContextID
    'laVentana  será la ventana que se usará para mostrar la ayuda
    '
    Dim vControl As Control
    Dim TopicNum As Long
    Dim sHelpFile As String
    
    On Local Error Resume Next
    
    If elControl Is Nothing Then
        Set vControl = Screen.ActiveControl
    Else
        Set vControl = elControl
    End If
    
    TopicNum = CLng(vControl.HelpContextID)
    
    If Err = 0 Then
        sHelpFile = SoloCHM
        If Len(laVentana) Then
            sHelpFile = sHelpFile & " >" & laVentana
        End If
        HelpContext = HtmlHelp(hWnd, sHelpFile, HH_HELP_CONTEXT, ByVal TopicNum)
    End If
    
    Err = 0
End Function



