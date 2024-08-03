# READ_EXCEL_Fsharp

`READ_EXCEL_Fsharp` es un proyecto desarrollado en F# que se conecta a un archivo Excel para generar un plan de cuotas. Utiliza un archivo Excel llamado "Calculadora de préstamos simple y tabla de amortización", disponible en [este enlace](https://create.microsoft.com/es-es/template/calculadora-de-pr%25C3%25A9stamos-simple-y-tabla-de-amortizaci%25C3%25B3n-923c86b5-63f8-42d1-99cb-c6ae4f4b679e).

## Descripción

Este proyecto permite ingresar parámetros mediante consola para interactuar con un archivo Excel, ejecutar cálculos y generar un plan de cuotas basado en los datos proporcionados. 

## Características

- **Desarrollado en F#**: Todo el proyecto está desarrollado usando el lenguaje de programación F#.
- **Interacción con Excel**: Utiliza late binding para interactuar con el archivo Excel.
- **Entrada por consola**: Permite al usuario ingresar parámetros directamente desde la consola.
- **Generación de Plan de Cuotas**: Calcula y muestra un plan de cuotas basado en los datos del archivo Excel.

## Requisitos

- [Microsoft Excel](https://www.microsoft.com/excel): Necesario para abrir y manipular el archivo Excel.
- [F#](https://fsharp.org): El código está escrito en F# y se necesita el compilador o entorno adecuado para ejecutarlo.

## Instalación

1. **Clonar el repositorio:**

    ```bash
    git clone https://github.com/tu_usuario/READ_EXCEL_Fsharp.git
    cd READ_EXCEL_Fsharp
    ```

2. **Instalar dependencias:**

    Asegúrate de tener todas las dependencias necesarias. Puedes instalar las dependencias usando el gestor de paquetes adecuado si es necesario.

## Uso

1. **Abrir el archivo Excel**: Asegúrate de que el archivo Excel "Calculadora de préstamos simple y tabla de amortización" esté disponible en la ruta especificada en el código.

2. **Ejecutar el programa**:

    Ejecuta el programa desde la consola. Se te pedirá que ingreses los parámetros necesarios para el cálculo.

    ```bash
    dotnet run
    ```

3. **Ingresar parámetros**: Proporciona los parámetros solicitados cuando se te pida. El programa actualizará el archivo Excel y generará el plan de cuotas.

