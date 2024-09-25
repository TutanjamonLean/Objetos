VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21075
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   21075
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   5280
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11280
      TabIndex        =   1
      Top             =   3600
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8040
      TabIndex        =   0
      Top             =   3600
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Medir As String
Dim objMascotas1 As New Mascotas
Private Sub Command1_Click()
        
'    objMascotas1.edad = 69
'
'    objMascotas1.peso = 60
'
'    objMascotas1.raza = "ario"
'
'    objMascotas1.especie = "antisemita"
'
'    objMascotas1.color = "humilde"
'
'    objMascotas1.amigable = False
'
'    objMascotas1.tamaño = "enano"
'
'    objMascotas1.talle = 60
    
End Sub

Private Sub Command2_Click()
    
    Print objMascotas1.Getedad

    
'    Print objMascotas1.edad
'
'    Print objMascotas1.peso
'
'    Print objMascotas1.raza
'
'    Print objMascotas1.especie
'
'    Print objMascotas1.color
'
'    Print objMascotas1.amigable
'
'    Print objMascotas1.tamaño
'
'    Print objMascotas1.talle
    
End Sub

Private Sub Form_Activate()
    
    objMascotas1.constructor 15
    
End Sub

