VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "RUT CHILE: Validar y Formatear"
   ClientHeight    =   4890
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

Private Sub btnSalir_Click()
    
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    Application.Quit
End Sub

Private Sub btnVerificar_Click()

   Dim rut As String
   
   rut = limpiarRut(txtRut.Text)
   txtSalida.ForeColor = vbGrayText
   
   If esRut(rut) Then
   
         If (verificaRut(rut)) Then
            txtRespuesta.Caption = "RUT ingresado es válido"
            txtSalida.Caption = formatearRut(rut)
            txtSalida.ForeColor = vbGreen
         Else
            txtRespuesta.Caption = "RUT ingresado no es válido"
            txtSalida.Caption = rut
            txtSalida.ForeColor = vbRed
         End If
   Else
         txtRespuesta.Caption = "Formato de RUT no válido"
         txtSalida.Caption = rut
         txtSalida.ForeColor = vbBlue
   End If

End Sub
