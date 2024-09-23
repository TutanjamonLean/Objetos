VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mascotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private nombre As String
Private edad As Integer
Private peso As Double
Private raza As String
Private especie As String
Private color As String
Private amigable As Boolean
Private tamaño As String
Private talle As Double

Public Sub constructor(thisnombre As String)
    
    nombre = thisnombre
    
End Sub
