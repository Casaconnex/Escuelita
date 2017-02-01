VERSION 5.00
Begin VB.Form FONDO 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form9"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   1080
      Picture         =   "FONDO.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "FONDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
acceso.Show
End Sub

Private Sub Form_Load()
Image1.Left = 0
Image1.Top = 0
Image1.Height = Screen.Height
Image1.Width = Screen.Width
End Sub
