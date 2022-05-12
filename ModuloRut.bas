Attribute VB_Name = "ModuloRut"

' /*-------------------------------[ ModuloRut ]---------------------------------/
'  Autor       : Gonen09
'  Descripción : Módulo para validar y formatear RUT en Excel VBA
'  Versión     : 1.0
'  Fecha       : 07/03/2022
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

Function limpiarRut(rut As String) As String

   Dim nuevoRut As String
   
   nuevoRut = Trim(rut)
   nuevoRut = Replace(nuevoRut, " ", "")
   nuevoRut = Replace(nuevoRut, vbTab, "")

   limpiarRut = nuevoRut

End Function

Function quitarFormato(rut As String) As String

   Dim nuevoRut As String   

   nuevoRut = UCase(rut)
   nuevoRut = Replace(nuevoRut, ".", "")
   nuevoRut = Replace(nuevoRut, "-", "")

   quitarFormato = nuevoRut

End Function
   
'Verifica RUT sin formato
Function esRut(rut As String) As Boolean

   Dim numeroRut As String
   Dim dv As String
   
   If (rut <> "") Then
   
      If (Len(rut) > 7 And Len(rut) < 10) Then

         numeroRut = Left(rut, Len(rut) - 1)

         If (IsNumeric(numeroRut)) Then

            dv = Right(rut, 1)

            If (dv <> "" And dv <> " " And dv <> vbTab) Then

               If (isnumeric(dv)) Then 
      
                  If (dv > -1) And (dv < 10) Then
                     esRut = True
                     Exit Function
                  End If
            
               Else
               
                  If (dv = "K") Then
                     esRut = True
                     Exit Function
                  End If

               End If

            End If
         End If
      End If
   End If 

   esRut = False

End Function

'Módulo obtenido de la web, autor no encontrado, créditos a su autor.  
Public Function rutDigito(ByVal Rut As Long) As String

   Dim Digito As Integer
   Dim Contador As Integer
   Dim Multiplo As Integer
   Dim Acumulador As Integer

   Contador = 2
   Acumulador = 0
   
   While Rut <> 0
   
      Multiplo = (Rut Mod 10) * Contador
      Acumulador = Acumulador + Multiplo
      Rut = Rut \ 10
      Contador = Contador + 1
      
      If Contador = 8 Then
         Contador = 2
      End If
      
   Wend
   
   Digito = 11 - (Acumulador Mod 11)
   rutDigito = CStr(Digito)
   
   If Digito = 10 Then rutDigito = "K"
   If Digito = 11 Then rutDigito = "0"
   
End Function

Function verificaRut(rut As String) As Boolean

   Dim numeroRut As String
   Dim dve As String
   Dim dvs As String

   numeroRut = Left(rut, Len(rut) - 1)
   dve = Right(rut, 1)
   dvs = rutDigito(CLng(numeroRut))

   If (dve = dvs) Then 
      verificaRut = True
   Else
      verificaRut = False
   End If

End Function

Function formatearRut(rut As String) As String

   Dim numeroRut As String
   Dim dv As String

   numeroRut = Left(rut, Len(rut) - 1)
   dv = Right(rut, 1)

   numeroRut = Format(numeroRut, "##,###,###")
   numeroRut = Replace(numeroRut, ",", ".")
   
   formatearRut = numeroRut & "-" & dv

End Function
