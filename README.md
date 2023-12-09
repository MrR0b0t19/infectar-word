# Word Shellcode Example

Este repositorio contiene un ejemplo simple de cómo insertar un shellcode en un documento de Word utilizando C# y la biblioteca Microsoft.Office.Interop.Word.
Requisitos

    Visual Studio 2022 (u otra versión compatible)
    Microsoft Word (asegúrate de tener instalado Microsoft Word en tu máquina)

Instrucciones

    Clona o descarga este repositorio en tu máquina local.

    Abre el proyecto en Visual Studio.

    Asegúrate de agregar la referencia adecuada al ensamblado Microsoft.Office.Interop.Word. Puedes hacer esto siguiendo estos pasos:
        Haz clic derecho en el proyecto en el Explorador de Soluciones.
        Selecciona "Agregar" -> "Referencia...".
        Busca "Microsoft Word" o "Microsoft.Office.Interop.Word" y selecciona la versión correspondiente.
        Haz clic en "Aceptar".

    Ejecuta la aplicación. Esto abrirá Microsoft Word, creará un nuevo documento y le insertará un shellcode.

    Verifica el archivo documento_con_shellcode.docx en el directorio del proyecto. Este archivo contiene el documento de Word con el shellcode incrustado. de las versiones anteriores para escuchar el exploit es con metasploit saludos

Nota Importante
Ten en cuenta las implicaciones éticas y legales al utilizar técnicas de este tipo. ¡Usa este conocimiento sabiamente!
