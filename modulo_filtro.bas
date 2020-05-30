Attribute VB_Name = "modulo_filtro"
'Código para el filtro avanzado
Sub TableFiltro()
Dim ContNombre, ContEmpresa, ContCurso, ContEmail, ContDocente As String
Dim FeFin, FeIn As Date
Dim Ufila As Long

With Hoja1
Ufila = .Range("C999999").End(xlUp).Row
If Ufila < 11 Then Ufila = 11
If .Range("C10").Value = "Nombre" Then ContNombre = Empty Else: ContNombre = .Range("C10").Value
If .Range("D10").Value = "Correo" Then ContEmail = Empty Else: ContEmail = .Range("D10")
If .Range("F10").Value = "Curso" Then ContCurso = Empty Else: ContCurso = .Range("F10")
If .Range("H10").Value = "Empresa" Then ContEmpresa = Empty Else: ContEmpresa = .Range("H10")
If .Range("P10").Value = "Docente" Then ContDocente = Empty Else: ContDocente = .Range("P10")
If .Range("N10").Value = "Fecha Inicio" Then FeIn = "01/01/1900" Else: FeIn = .Range("N10")
If .Range("O10").Value = "Fecha Fin" Then FeFin = "31/12/2035" Else: FeFin = .Range("O10")
.Range("C11:R" & Ufila).Select
Selection.AutoFilter
    With .Range("C11:R" & Ufila)
    
        If ContNombre <> Empty Then .AutoFilter Field:=1, Criteria1:="=*" & ContNombre & "*"
'        .AutoFilter Field:=12, Criteria1:=">=" & FeIn, Operator:=xlAnd, Criteria2:="<=" & FeFin
        
        If ContEmail <> Empty Then .AutoFilter Field:=2, Criteria1:="=*" & ContEmail & "*"
        If ContCurso <> Empty Then .AutoFilter Field:=4, Criteria1:="=*" & ContCurso & "*"
        If ContEmpresa <> Empty Then .AutoFilter Field:=6, Criteria1:="=*" & ContEmpresa & "*"
        If ContDocente <> Empty Then .AutoFilter Field:=14, Criteria1:="=*" & ContEmail & "*"
    End With
  .Range("11:11").EntireRow.Hidden = True
  
    

End With


End Sub

'Función que limpia el filtro
Sub ClearFilt()
With Hoja1
.Range("B4").Value = True
.AutoFilterMode = False
.Range("C10").Value = "Nombre"
.Range("D10").Value = "Correo"
.Range("F10").Value = "Curso"
.Range("H10").Value = "Empresa"
.Range("P10").Value = "Docente"
.Range("N10").Value = "Fecha Inicio"
.Range("O10").Value = "Fecha Fin"
.Range("B4").Value = False
End With
End Sub
