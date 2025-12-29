# Sistema de Medición de Frecuencias CNT-91

## Descripción
Sistema de adquisición de datos para el frecuencímetro Timer/Counter/Analyzer CNT-91, desarrollado como parte del TFG de Carlos Darvoy Espigulé.

## Características
- **Medición continua de frecuencias** en tiempo real
- **Análisis estadístico** con desviación de Allan
- **Exportación de datos** a Excel
- **Visualización de resultados** con gráficas
- **Comunicación VISA** con instrumentos GPIB/USB

## Instalación

### Requisitos Previos
- Python 3.10 o superior (recomendado Python 3.11 para mejor compatibilidad)
- Conexión GPIB o USB al instrumento CNT-91

### Instalación de Dependencias
```bash
pip install -r requirements.txt
```

## Estructura del Proyecto

### Archivos Principales
- **`CNT_9X_pendulum.py`** - Clase principal para comunicación con el CNT-91
- **`Datalogger.py`** - Script principal para adquisición de datos
- **`ESTADÍSTICO CON EL CNT.py`** - Análisis estadístico y gráficas
- **`Frontend.py`** - Interfaz gráfica (en desarrollo)

### Archivos de Configuración
- **`requirements.txt`** - Dependencias del proyecto
- **`Datalogger_backup.py`** - Versión de respaldo del datalogger

## Uso

### Configuración Básica
1. Conectar el instrumento CNT-91 vía GPIB o USB
2. Ejecutar el script principal:
   ```bash
   python Datalogger.py
   ```

### Configuración Avanzada
- Modificar parámetros en `Datalogger.py`:
  - `Address_dispositivo` - Dirección del instrumento
  - `canal` - Canal de medición (A/B o 1/2)
  - `intervalo_s` - Intervalo entre mediciones
  - `acoplamiento` - Tipo de acoplamiento (DC/AC)
  - `impedancia` - Impedancia de entrada

### Análisis Estadístico
```bash
python "ESTADÍSTICO CON EL CNT.py"
```

## Dependencias Principales

### Comunicación con Instrumentos
- `pyvisa>=1.15.0` - Interfaz VISA para comunicación con instrumentos
- `pyvisa-py>=0.8.0` - Backend de PyVISA en Python puro
- `zeroconf>=0.38.0` - Descubrimiento automático de dispositivos

### Análisis de Datos
- `numpy>=1.21.0` - Cálculos numéricos
- `matplotlib>=3.5.0` - Visualización de datos
- `scipy>=1.7.0` - Análisis estadístico avanzado
- `pandas>=1.3.0` - Manipulación de datos tabulares

### Exportación
- `openpyxl>=3.1.0` - Exportación a Excel

## Funcionalidades

### Medición Continua
- Adquisición de datos en tiempo real
- Configuración flexible de parámetros
- Almacenamiento automático en Excel

### Análisis de Estabilidad
- Cálculo de desviación de Allan
- Gráficas log-log de estabilidad vs τ
- Análisis estadístico interno del instrumento

### Configuración del Instrumento
- Configuración de canales (A/B)
- Ajuste de acoplamiento (DC/AC/HF/LF)
- Configuración de impedancia (50Ω/1MΩ)
- Ajuste de atenuación
- Configuración de triggers
- Filtros digitales y analógicos

## Solución de Problemas

### Error: "No se encontraron recursos VISA"
- Verificar conexión física del instrumento
- Comprobar drivers GPIB/USB instalados
- Verificar dirección del instrumento

### Error de Compatibilidad con Python 3.13
- Usar Python 3.11 para mejor compatibilidad
- Actualizar dependencias: `pip install --upgrade pyvisa pyvisa-py`

## Autor
Carlos Darvoy Espigulé - TFG

## Licencia
Proyecto académico - Universidad de las Islas Baleares 