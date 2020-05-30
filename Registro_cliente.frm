VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Registro_cliente 
   Caption         =   "Registro"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10170
   OleObjectBlob   =   "Registro_cliente.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Registro_cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Funcionalidades Formulario
'Botón aceptar
Private Sub aceptar_Click()

    If correo.Value <> "" Then
    
        Range("d1000000").End(xlUp).Offset(1, 0).Select
        ActiveCell.Value = correo.Value
        ActiveCell.Offset(0, -1) = nombre1.Value
        ActiveCell.Offset(0, 2) = curso1.Value
        ActiveCell.Offset(0, 3) = nivel_1.Value
        ActiveCell.Offset(0, 4) = empresa1.Value
        ActiveCell.Offset(0, 5) = horas1.Value
        ActiveCell.Offset(0, 6) = primer_pago1.Value
        ActiveCell.Offset(0, 7) = segundo_pago1.Value
        ActiveCell.Offset(0, 10) = fecha1_1.Value
        ActiveCell.Offset(0, 12) = docente1.Value
        ActiveCell.Offset(0, 11) = ActiveCell.Offset(0, 10).Value + 30
        ActiveCell.Offset(0, 11).NumberFormat = "dd/mm/yyyy"
        ActiveCell.Offset(0, 13) = ActiveCell.Offset(0, 6).Value * 0.3
        ActiveCell.Offset(0, 13).NumberFormat = "#.##0€"
        Registro_cliente.Hide
        
        
    Else
    
        MsgBox "Correo requerido"
        
    End If
    
    nombre1.Value = ""
    curso1.Value = ""
    nivel_1.Value = ""
    empresa1.Value = ""
    horas1.Value = ""
    primer_pago1.Value = ""
    segundo_pago1.Value = ""
    fecha1_1.Value = ""
    docente1.Value = ""
    correo.Value = ""
    correo.SetFocus
    
    
End Sub

'Botón cancelar
Private Sub CommandButton2_Click()

Registro_cliente.Hide

End Sub


Private Sub fecha_1_1_Change()

End Sub


Private Sub fecha1_Change()

End Sub

Private Sub correo_Change()

End Sub

Private Sub nombre1_Change()

End Sub

Private Sub UserForm_Initialize()
    nivel_1.AddItem "No Aplica"
    nivel_1.AddItem "Básico"
    nivel_1.AddItem "Intermedio"
    nivel_1.AddItem "Avanzado"
    nivel_1.AddItem "Profesional"
End Sub



