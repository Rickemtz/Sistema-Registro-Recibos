Claro, aquí tienes un ejemplo de README para tu proyecto en GitHub:

---

# Sistema de Registro y Generación de Recibos

Este proyecto es una aplicación de escritorio desarrollada en Python que permite registrar aportaciones de socios, guardar los registros en un archivo Excel y generar recibos impresos. La interfaz gráfica está construida utilizando Tkinter.

## Características

- **Registro de Socios**: Permite ingresar el número de socio, nombre, monto de ingreso y fecha.
- **Actualización de Registros**: Actualiza la información de un socio existente o agrega un nuevo registro.
- **Almacenamiento en Excel**: Guarda los registros en un archivo Excel (`Registros.xlsx`).
- **Generación de Recibos**: Genera e imprime recibos utilizando una impresora USB compatible con comandos ESC/POS.
- **Interfaz Gráfica**: Interfaz de usuario intuitiva construida con Tkinter.

## Requisitos

- Python 3.x
- Pandas
- Tkinter
- tkcalendar
- pyusb
- openpyxl

Puedes instalar los requisitos ejecutando:

```bash
pip install pandas tkcalendar pyusb openpyxl
```

## Uso

1. **Clona el repositorio**:

   ```bash
   git clone https://github.com/Rickemtz/Sistema-Registro-Recibos.git
   cd Sistema-Registro-Recibos
   ```

2. **Ejecuta la aplicación**:

   ```bash
   python app.py
   ```

3. **Interfaz de usuario**:
   - Ingresa el número de socio, nombre, monto de ingreso y selecciona la fecha.
   - Haz clic en "Registrar" para guardar el registro y generar un recibo.
   - Usa "Buscar Socio" para cargar información de un socio existente.
   - Usa "Imprimir Último Registro" para imprimir el recibo del último registro ingresado.
