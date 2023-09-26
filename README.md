# Excel Row Copy Tool 📊

El Excel Row Copy Tool es una herramienta de línea de comandos desarrollada en Python que te permite copiar filas específicas de un archivo Excel (xlsx) y pegarlas en un nuevo archivo Excel. Esta herramienta es útil cuando tienes un archivo Excel grande y solo necesitas ciertas filas para trabajar.

# Características 🚀

✂️ Copia filas específicas de un archivo Excel original a un nuevo archivo Excel.
🖋️ Personaliza el nombre del archivo de salida y el título de la hoja.
📋 Puedes seleccionar múltiples filas para copiar, separándolas por comas.
📝 Puedes asignar nombres personalizados a las columnas en el nuevo archivo.

# Requisitos Previos:

Asegúrate de tener Python instalado en tu sistema. Si no lo tienes instalado, puedes descargarlo desde python.org.

Descarga el Código Fuente:

Descarga el código fuente de esta herramienta desde GitHub.

# Ejecución del Programa:

Abre una terminal o línea de comandos en la ubicación donde se encuentra el archivo excel_row_copy.py y ejecuta el siguiente comando:

python excel_row_copy.py


# Instrucciones de Uso:

📄 Ingresa un nombre para el archivo de salida (sin la extensión .xlsx).
📌 Proporciona un título para la hoja del nuevo archivo Excel.
💡 Especifica las filas que deseas copiar. Puedes ingresar múltiples filas separadas por comas sin espacios. Por ejemplo, si deseas copiar las filas 1, 3 y 5, ingresa: 1,3,5.
📇 Para cada fila que selecciones, se te pedirá que ingreses un nombre para esa fila. Esto permitirá personalizar el nombre de las columnas en el nuevo archivo.
Resultado:

La herramienta creará un nuevo archivo Excel con el nombre que proporcionaste y pegará las filas seleccionadas en él.

# Ejemplo de Uso 📋
Supongamos que tienes un archivo Excel llamado "Lista-de-clientes-con-nombre-y-direccion.xlsx" que contiene datos de clientes y deseas copiar las filas 1, 3 y 5 a un nuevo archivo con el nombre "Clientes-Seleccionados". Puedes seguir estos pasos:

Ejecuta el programa y sigue las instrucciones para proporcionar los detalles requeridos.

Ingresa "Clientes-Seleccionados" como nombre del archivo de salida.

Proporciona un título, por ejemplo, "Clientes Seleccionados".

Ingresa 1,3,5 como las filas que deseas copiar.

Se te pedirá que ingreses un nombre para cada fila. Por ejemplo, puedes ingresar "Nombre", "Dirección" y "Correo Electrónico" para las filas 1, 3 y 5 respectivamente.

La herramienta creará un nuevo archivo Excel llamado "Clientes-Seleccionados.xlsx" con las filas seleccionadas y los nombres de columna personalizados.

# Manual de Uso 📖
Nombre del Archivo (sin extensión): Ingresa un nombre para el archivo de salida sin incluir la extensión .xlsx.

Título de la Hoja: Proporciona un título para la hoja del nuevo archivo Excel.

Número de Fila a Copiar: Ingresa las filas que deseas copiar separadas por comas sin espacios. Por ejemplo, para copiar las filas 1, 3 y 5, ingresa: 1,3,5.

Nombres Personalizados para Filas: Para cada fila que selecciones, se te pedirá que ingreses un nombre para esa fila. Esto permitirá personalizar el nombre de las columnas en el nuevo archivo.

Archivo Existente: Si el archivo de salida ya existe, se te informará y se te pedirá que ingreses otro nombre.

Proceso de Creación: Una vez que hayas proporcionado todos los detalles, la herramienta copiará las filas seleccionadas del archivo original al nuevo archivo Excel, asignando los nombres de columna personalizados si los has proporcionado.

Éxito: Una vez que se haya creado el archivo de salida, se mostrará un mensaje de éxito.

Esperamos que esta herramienta te ayude a trabajar de manera más eficiente con archivos Excel al permitirte seleccionar y copiar solo las filas que necesitas. ¡Disfruta usándola!

# Notas 📌
Este proyecto está disponible en GitHub. Si tienes sugerencias de mejoras o encuentras algún problema, no dudes en crear un problema o enviar una solicitud de extracción.

Asegúrate de tener instalada la biblioteca openpyxl. Si no la tienes instalada, puedes hacerlo ejecutando el siguiente comando: