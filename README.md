# Rut Chile
Proyecto para validar y formatear RUT en Excel VBA 2019.

# Modo de uso

Puedes usar el archivo Excel "Rut Chile" directamente, ingresando el RUT
en el formulario (ventana) o ingresando el RUT en la celda "Ingresar RUT:".

# Instalación

Se debe crear un archivo Excel habilitado para macros (.xlsm). Abrir Excel,
Archivo, Guardar como y en tipo seleccionar "Libro de Excel habilitado para
macros".

Para utilizar este proyecto e integrarlo a un archivo Excel existente se
debe importar el archivo "ModuloRut.bas".

Para importar el archivo debemos abrir la ventana "Microsoft Visual Basic 
para Aplicaciones".

Debemos presionar la combinación de teclas "ALT" y "F11". En caso de no funcionar
debemos dirigirnos a la pestaña "Programador/Desarrollador" y en "Visual Basic".

En caso de no aparecer la pestaña "Programador/Desarrollador" debemos dirigirnos a
Archivo, Opciones, Personalizar cinta de opciones, Pestañas principales y activar
la opción "Programador/Desarrollador" según corresponda y aceptar.

A continuación, nos aparecerá una ventana con la estructura del proyecto, luego
seleccionamos en Archivo y en "importar archivo..." y finalmente seleccionamos 
el archivo "ModuloRut.bas" creándose una carpeta "Módulos" que contenga el archivo.

# Integrar a Formulas Excel

Para validar el RUT debes copiar la fórmula de la celda asignada a "Respuesta" y
reemplazar el valor de "B4" según la celda que utilizas en tú proyecto para obtener 
el RUT.

Para formatear el RUT debes copiar la fórmula de la celda asignada a "Valor" que lee
el contenido de "Respuesta". Sino solo debes llamar a la "formatearRut", esta recibe un
RUT limpio (sin espacios), sin formato y valido, funciones incluidas en el sistema.

El sistema de reglas de colores de "Valor" se encuentra en Inicio, Formato condicional,
Administrar reglas (se debe seleccionar la celda previamente). Se puede eliminar las
reglas o editar los colores seleccionando la regla, editar, formato, el color del texto
en Fuente, Color y el Fondo en Relleno y aceptar.

# Integrar en Macros VBA

Una vez importado el archivo "ModuloRut.bas". para integrar la funcionabilidad a tú 
propio proyecto puedes guiarte en la estructura de las llamadas de las funciones
en el archivo "UserForm.frm" en el método "btnVerificar_Click".

Si deseas integrar el formulario (ventana) completo debes importar los archivos,
"Módulo.bas" y "UserForm.frm". 

Luego ir a la pestaña "Programador/Desarrollador", Insertar, Controles de formulario e
insertar un botón y seleccionar la macro "Abrir Formulario".

Si deseas que el formulario se abra automáticamente  al iniciar Excel, debes abrir 
la ventana "Microsoft Visual Basic para Aplicaciones" ("ALT" y "F11") y buscar en el 
proyecto "ThisWorkbook", abrirlo y pegar el siguiente código:

        Private Sub Workbook_Open()
            UserForm.Show
        End Sub


# Licencia 
Eres libre de usar estos archivos según estimes conveniente, si estos archivos son útiles para ti y/o te ayudan a generar tu propio contenido y distribuirlo, ya sea de manera gratuita o pagada, si te es posible, te agradecería realizar una donación, de antemano muchas gracias.