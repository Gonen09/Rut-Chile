VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "RUT CHILE: Validar y Formatear"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5145
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function limpiarRut(rut As String) As String

   Dim nuevoRut As String

   nuevoRut = Trim(rut)
   nuevoRut = Replace(nuevoRut, " ", "")
   nuevoRut = Replace(nuevoRut, vbTab, "")

   limpiarRut = nuevoRut

End Function

Function esRut (rut as String) as Boolean

   Dim nuevoRut As String
   Dim dv As String
   Dim punto, guion As Integer
   Dim lista() As String
   
   nuevoRut = UCase(rut)

   If (nuevoRut <> "") Then
   
      'RUT con formato
      If ((Instr(nuevoRut, ".") > 0) and (Instr(nuevoRut, "-") > 0)) Then 

         punto = Len(nuevoRut) - Len(Replace(nuevoRut, ".", ""))
         guion = Len(nuevoRut) - Len(Replace(nuevoRut, "-", ""))

         If ((punto = 2) and (guion = 1)) Then 

            lista = split(nuevoRut,"-")
            dv = lista(1)

            If (len(dv) <> 1) Then 
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

         dv = right(nuevoRut, 1)
         
         If (dv <> "" And dv <> " " And dv <> vbTab) Then

            If (((dv > -1) And (dv <10)) Or (dv = "K")) Then 

               nuevoRut = LEFT(nuevoRut, Len(nuevoRut) - 1)
      
               If (isnumeric(nuevoRut)) Then
                  esRut = True
               Else
                  esRut = False
               End If
            else
               esRut = False
            end if 
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