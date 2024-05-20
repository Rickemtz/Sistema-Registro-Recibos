import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
import usb.core
import usb.util

# Inicializar el DataFrame para almacenar los registros
try:
    df = pd.read_excel('Registros.xlsx')
except FileNotFoundError:
    df = pd.DataFrame(columns=['Número de Socio', 'Nombre'])

# Función para actualizar el DataFrame con un nuevo ingreso o actualización
def actualizar_dataframe(numero_socio, nombre, monto, fecha):
    global df
    
    if 'Número de Socio' in df.columns and numero_socio in df['Número de Socio'].values:
        idx = df[df['Número de Socio'] == numero_socio].index[0]
        df.at[idx, 'Nombre'] = nombre
        df.at[idx, fecha] = monto
    else:
        nuevo_registro = pd.DataFrame({'Número de Socio': [numero_socio], 'Nombre': [nombre], fecha: [monto]})
        df = pd.concat([df, nuevo_registro], ignore_index=True)

# Función para guardar el DataFrame actualizado en un archivo Excel
def guardar_en_excel():
    try:
        df.to_excel('Registros.xlsx', index=False)
    except PermissionError:
        messagebox.showerror("Error de permisos", "No tienes permisos para guardar en el archivo Excel. Verifica los permisos de escritura.")
    except Exception as e:
        messagebox.showerror("Error al guardar", f"Ocurrió un error al guardar en el archivo Excel: {e}")

# Función para enviar datos a la impresora
def enviar_datos(texto):
    try:
        # Encuentra el dispositivo USB
        dev = usb.core.find(idVendor=0x0483, idProduct=0x5743)
        if dev is None:
            messagebox.showerror("Error de impresión", "Impresora no encontrada.")
            return
        
        # Activa la interfaz de la impresora
        interface = 0
        if dev.is_kernel_driver_active(interface):
            dev.detach_kernel_driver(interface)
        usb.util.claim_interface(dev, interface)
        
        # Comandos ESC/POS para imprimir el recibo
        esc_pos_commands = (
            b'\x1b\x40',                 # Inicializar impresora
            b'\x1b\x61\x01',             # Centrando texto
            texto.encode(),              # Texto del recibo
            b'\n\n\n\n\n',               # Espacio extra al final
            b'\x1d\x56\x42\x00'          # Corte del papel
        )
        
        for cmd in esc_pos_commands:
            dev.write(1, cmd, 1000)
        
        usb.util.release_interface(dev, interface)
        usb.util.dispose_resources(dev)
        
        messagebox.showinfo("Impresión", "¡Recibo generado con éxito!")
    except usb.core.USBError as e:
        messagebox.showerror("Error de impresión", f"Error USB: {e}")
    except Exception as e:
        messagebox.showerror("Error de impresión", f"Error al generar el recibo: {e}")

# Prueba de la función
enviar_datos("Texto de prueba")

# Función para imprimir un recibo
def imprimir_registro(numero_socio, nombre, monto, fecha):
    texto_recibo = (
        f"                RECIBO               \n"
        f"*************************************\n"        
        f"        Recibos por Aportación       \n"     
        f"                                     \n"
        f"Socio: {numero_socio} {nombre}       \n"
        f"       $    {monto:.2f}      {fecha} \n"
        f"*************************************\n"
        f"     GRACIAS POR SU COLABORACIÓN     \n"
    )
    
    recibos_text.delete(1.0, tk.END)
    recibos_text.insert(tk.END, texto_recibo)
    
    enviar_datos(texto_recibo)

# Función para registrar un nuevo ingreso
def registrar():
    try:
        numero_socio = int(numero_socio_entry.get())
        nombre = nombre_entry.get().strip()
        monto = float(monto_entry.get())
        fecha_seleccionada_value = fecha_seleccionada.get_date().strftime("%d/%m/%Y")

        actualizar_dataframe(numero_socio, nombre, monto, fecha_seleccionada_value)
        guardar_en_excel()
        imprimir_registro(numero_socio, nombre, monto, fecha_seleccionada_value)
        
        numero_socio_entry.delete(0, tk.END)
        nombre_entry.delete(0, tk.END)
        monto_entry.delete(0, tk.END)
        messagebox.showinfo("Registro", "¡Registro exitoso!")
    except ValueError as e:
        messagebox.showerror("Error", f"Por favor, ingrese valores válidos: {e}")

# Función para buscar y cargar información de un socio
def buscar_socio():
    try:
        numero_socio = int(numero_socio_entry.get())
        if 'Número de Socio' in df.columns and numero_socio in df['Número de Socio'].values:
            socio_info = df[df['Número de Socio'] == numero_socio].iloc[0]
            nombre_entry.delete(0, tk.END)
            nombre_entry.insert(0, socio_info['Nombre'])
        else:
            messagebox.showinfo("Información", f"No se encontró información para el Número de Socio {numero_socio}.")
    except ValueError:
        messagebox.showerror("Error", "Ingrese un número de socio válido.")

# Función para imprimir el recibo del último registro
def imprimir_ultimo_registro():
    if not df.empty:
        ultimo_registro = df.iloc[-1]
        numero_socio = ultimo_registro['Número de Socio']
        nombre = ultimo_registro['Nombre']
        fecha_seleccionada_value = fecha_seleccionada.get_date().strftime("%d/%m/%Y")
        monto = ultimo_registro.get(fecha_seleccionada_value, 0.0)

        imprimir_registro(numero_socio, nombre, monto, fecha_seleccionada_value)
    else:
        messagebox.showerror("Error", "No hay registros para imprimir.")

# Configuración de la interfaz gráfica
root = tk.Tk()
root.title("Sistema de Registro y Generación de Recibos")

tk.Label(root, text="Número de Socio:").grid(row=0, column=0)
numero_socio_entry = tk.Entry(root)
numero_socio_entry.grid(row=0, column=1)

tk.Label(root, text="Nombre:").grid(row=1, column=0)
nombre_entry = tk.Entry(root)
nombre_entry.grid(row=1, column=1)

tk.Label(root, text="Monto de ingreso:").grid(row=2, column=0)
monto_entry = tk.Entry(root)
monto_entry.grid(row=2, column=1)

tk.Label(root, text="Fecha:").grid(row=3, column=0)
fecha_seleccionada = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
fecha_seleccionada.grid(row=3, column=1)

registrar_button = tk.Button(root, text="Registrar", command=registrar)
registrar_button.grid(row=4, column=0, columnspan=2, pady=10)

recibos_text = tk.Text(root, height=20, width=50)
recibos_text.grid(row=5, column=0, columnspan=2)

buscar_button = tk.Button(root, text="Buscar Socio", command=buscar_socio)
buscar_button.grid(row=0, column=2, pady=10)

imprimir_button = tk.Button(root, text="Imprimir Último Registro", command=imprimir_ultimo_registro)
imprimir_button.grid(row=4, column=2, pady=10)

root.mainloop()

    