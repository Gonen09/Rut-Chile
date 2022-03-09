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
