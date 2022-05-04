Attribute VB_Name = "ModuloRut"
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

               If (dv > -1 And dv < 10) Or (dv = "K") Then
                  esRut = True
               End If

            End If
         End If
      End If
   End If 

   esRut = False

End Function

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

   Dim nuevoRut As String
   Dim dve As String
   Dim dvs As String

   nuevoRut = UCase(rut)
   nuevoRut = Replace(nuevoRut, ".", "")
   nuevoRut = Replace(nuevoRut, "-", "")

   dve = Right(nuevoRut, 1)
   nuevoRut = Left(nuevoRut, Len(nuevoRut) - 1)

   dvs = rutDigito(CLng(nuevoRut))

   If (dve = dvs) Then 
      verificaRut = True
   Else
      verificaRut = False
   End If

End Function

Function formatearRut(rut As String) As String

   Dim nuevoRut As String
   Dim dv As String

   nuevoRut = UCase(rut)

   If ((InStr(nuevoRut, ".") > 0) And (InStr(nuevoRut, "-") > 0)) Then
      formatearRut = rut
   Else

      dv = Right(nuevoRut, 1)
      nuevoRut = Left(nuevoRut, Len(nuevoRut) - 1)

      nuevoRut = Format(nuevoRut, "##,###,###")
      nuevoRut = Replace(nuevoRut, ",", ".")
      nuevoRut = nuevoRut & "-" & dv
      
      formatearRut = nuevoRut
   End If

End Function
