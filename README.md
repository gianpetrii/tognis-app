# Procesador de Excel para Mobiliario

Herramienta para procesar archivos Excel de listas de cortes de mobiliario, extrayendo y consolidando información de materiales como **resina de melamina** y **okuome**.

## 🚀 Características

- **Interfaz gráfica**: Selección fácil de carpetas y archivos
- **Procesamiento automático**: Detecta todos los archivos Excel en una carpeta
- **Cálculos inteligentes**: Considera ensamblajes y subensamblajes
- **Múltiples reportes**: Detalle, resúmenes y queries agrupados
- **Materiales soportados**: Resina de melamina y Okuome

## 📁 Estructura del proyecto

```
├── codigo.py          # Script principal
├── README.md          # Este archivo
├── .gitignore         # Archivos a ignorar
└── datos/             # Carpeta de ejemplo (opcional)
```

## 🔧 Instalación y uso

### Opción 1: Usar el ejecutable (Recomendado)
1. Descarga `ProcesadorExcel.exe` desde [Releases](../../releases)
2. Ejecuta el archivo (no requiere Python)
3. Selecciona la carpeta con tus archivos Excel
4. Elige dónde guardar el resultado

### Opción 2: Ejecutar desde código fuente
```bash
# Instalar dependencias
pip install pandas openpyxl tkinter

# Ejecutar
python codigo.py
```

## 🏗️ Generar tu propio ejecutable

```bash
# Instalar PyInstaller
pip install pyinstaller

# Generar ejecutable
python -m PyInstaller --onefile --windowed --name="ProcesadorExcel" codigo.py
```

El ejecutable se creará en la carpeta `dist/`

## 📊 Funcionamiento

El script:
1. **Lee archivos Excel** de la carpeta seleccionada
2. **Filtra materiales** (resina de melamina y okuome)  
3. **Calcula cantidades** considerando jerarquías de ensamblaje
4. **Genera reportes** en formato Excel con múltiples hojas

### Cálculo de cantidades
```
Cantidad Total = Cantidad del Elemento × Cantidad de Ensamblaje × Cantidad de Subensamblaje
```

## 📋 Reportes generados

Para cada material se crean 4 hojas:
- **Detalle**: Listado completo con cantidades calculadas
- **Resumen por Nombre**: Agrupado por nombre específico  
- **Resumen por Archivo**: Totalizado por archivo origen
- **Query**: Agrupado por dimensiones (ancho, alto, espesor, canto)

## 🛠️ Tecnologías utilizadas

- **Python 3.9+**
- **pandas**: Manipulación de datos
- **openpyxl**: Lectura/escritura de Excel
- **tkinter**: Interfaz gráfica
- **PyInstaller**: Generación de ejecutables

## 📝 Licencia

Este proyecto es de uso interno. 