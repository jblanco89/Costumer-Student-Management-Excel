VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Pop-up de formulario
Private Sub CommandButton1_Click()

Registro_cliente.Show

End Sub

'Limpiar filtros
Private Sub Worksheet_Change(ByVal Target As Range)
If Not Intersect(Target, Range("C10:R10")) Is Nothing And Range("B4").Value = False Then
    TableFiltro
End If
End Sub

'Navegación resaltada
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

If Not Intersect(Target, Range("C10:R108")) Is Nothing Then
    Range("A2").Value = Target.Row
End If
End Sub

