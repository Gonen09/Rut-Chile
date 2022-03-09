Attribute VB_Name = "ModuloRut"
Function limpiarRut(rut As String) As String

   Dim nuevoRut As String
   
   nuevoRut = Trim(rut)
   nuevoRut = Replace(nuevoRut, " ", "")
   nuevoRut = Replace(nuevoRut, vbTab, "")

   limpiarRut = nuevoRut

End Function

Function esRut(rut As String) As Boolean

   Dim nuevoRut As String
   Dim dv As String
   Dim punto As Integer
   Dim guion As Integer
   Dim lista() As String
   
   nuevoRut = UCase(rut)

   If (nuevoRut <> "") Then
   
      'RUT con formato
      If ((InStr(nuevoRut, ".") > 0) And (InStr(nuevoRut, "-") > 0)) Then

         punto = Len(nuevoRut) - Len(Replace(nuevoRut, ".", ""))
         guion = Len(nuevoRut) - Len(Replace(nuevoRut, "-", ""))

         If ((punto = 2) And (guion = 1)) Then

            lista = Split(nuevoRut, "-")
            dv = lista(1)

            If (Len(dv) <> 1) Then
               esRut = False
               Exit Function
            Else
               esRut = False
               Exit Function
            End If
         Else
            esRut = False
            Exit Function
         End If
      End If

      nuevoRut = Replace(nuevoRut, ".", "")
      nuevoRut = Replace(nuevoRut, "-", "")

      'RUT sin formato
      If (Len(nuevoRut) > 7 And Len(nuevoRut) < 10) Then

         dv = Right(nuevoRut, 1)
         
         If (dv <> "" And dv <> " " And dv <> vbTab) Then

            If (((dv > -1) And (dv < 10)) Or (dv = "K")) Then

               nuevoRut = Left(nuevoRut, Len(nuevoRut) - 1)
      
               If (IsNumeric(nuevoRut)) Then
                  esRut = True
               Else
                  esRut = False
               End If
            Else
               esRut = False
            End If
         Else
            esRut = False
         End If
      Else
         esRut = False
      End If
   Else
      esRut = False
   End If

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

   dvs = rutDigito(CLong(nuevoRut))

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
