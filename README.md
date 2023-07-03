# YOUTUBE-CSharpImportExcelToDataGrid

## Librerias a utilizar 
* Definio que se utilizara [EPPLUS](https://www.nuget.org/packages/EPPlus/) para trabajar con excel.
* Para mostrar excel se utilizar la libreria **DevExpress Spreadsheet** o **Syncfusion Essential XlsIO**.

## Como hacer que el archivo de excel siempre se copie en las carpetas de bin

* Archivo de excel siempre se copie a bin/Debug o bin/Release para utilizar la ruta relativa sin problema
    - En el Explorador de soluciones de Visual Studio, haz clic con el botón derecho en el archivo de Excel y selecciona "Propiedades".

    - En la ventana de propiedades, busca la opción "Copiar al directorio de salida".

    - Cambia esta opción a "Copiar siempre" o "Copiar si es más reciente". "Copiar siempre" copiará el archivo al directorio de salida cada vez que construyas el proyecto, incluso si el archivo no ha cambiado. "Copiar si es más reciente" solo copiará el archivo si ha cambiado desde la última vez que se construyó el proyecto.

    - Guarda los cambios y construye el proyecto. El archivo de Excel debería ahora copiarse en el directorio de salida.

# Nota de paquete a usar para manipular archivos de EXCEL
    La elección del paquete para manipular archivos Excel en .NET depende en gran medida de tus necesidades específicas y de las características de tu proyecto. Aquí te proporciono una breve descripción de cada uno de los paquetes que mencionaste:

- **EPPlus**: EPPlus es una biblioteca muy popular y potente que te permite leer, escribir y manipular archivos Excel en formato .xlsx. Soporta una amplia gama de características de Excel, incluyendo fórmulas, gráficos, imágenes, tablas, etc. EPPlus también permite crear archivos Excel a partir de cero. Sin embargo, desde la versión 5.0, EPPlus se ha convertido en una biblioteca de pago para uso comercial.

- **SpreadsheetLight**: SpreadsheetLight es otra biblioteca para manipular archivos Excel. Se enfoca en la simplicidad y la facilidad de uso, y podría ser una opción más ligera si solo necesitas funcionalidades básicas. Sin embargo, no soporta tantas características como EPPlus.

- **ExcelDataReader**: ExcelDataReader es una biblioteca para leer archivos Excel (.xls y .xlsx). Es muy útil si solo necesitas leer datos de archivos Excel, pero no ofrece funcionalidades para escribir o manipular archivos.

Entonces, la elección depende de tus necesidades:

- Si necesitas una biblioteca poderosa que soporte una amplia gama de características de Excel y no te importa pagar por su uso comercial, EPPlus puede ser una buena opción.

- Si solo necesitas funcionalidades básicas y quieres algo más ligero y fácil de usar, SpreadsheetLight puede ser suficiente.

- Si solo necesitas leer datos de archivos Excel y no necesitas escribir ni manipular archivos, ExcelDataReader puede ser la mejor opción.

Recuerda que hay muchas otras bibliotecas disponibles para trabajar con archivos Excel en .NET, y podrías encontrar una que se adapte mejor a tus necesidades. Por ejemplo, ClosedXML es otra biblioteca popular que podría ser de interés, y NPOI es una opción si necesitas soporte para el formato más antiguo .xls además de .xlsx.