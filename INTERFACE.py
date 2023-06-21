import tkinter as tk
from tkinter import messagebox
import subprocess
# NOTAS #NOTAS # NOTAS # C:\Users\lamel\OneDrive\Documentos\PROGRAMAS PYTHON\AT FORMULARIOS IDIGER\Vias.py


def ejecutar_codigo():
    # NOTAS #NOTAS # NOTAS # Obtener el número de repeticiones y la ruta del archivo ingresados
    repeticiones = int(repeticiones_entry.get())
    ruta_archivo = ruta_entry.get()

    try:
        # NOTAS #NOTAS # NOTAS # Ejecutar el código la cantidad de repeticiones especificada
        contador = 0  # NOTAS #NOTAS # NOTAS # Inicializar el contador
        for _ in range(repeticiones):
            subprocess.call(["python", ruta_archivo])
            contador += 1  # NOTAS #NOTAS # NOTAS # Incrementar el contador en cada ejecución

            # NOTAS #NOTAS # NOTAS # Actualizar el valor del contador en la etiqueta
            contador_label.config(text=f"Contador: {contador}")
            ventana.update()  # NOTAS #NOTAS # NOTAS # Actualizar la ventana

        messagebox.showinfo(
            "Finalizado", "El código se ha ejecutado correctamente.")
    except Exception as e:
        messagebox.showerror(
            "Error", f"Ocurrió un error al ejecutar el código: {str(e)}")


# NOTAS #NOTAS # NOTAS # Crear la ventana principal
ventana = tk.Tk()
ventana.title("Ejecución de Código")

# NOTAS #NOTAS # NOTAS # Etiqueta y campo de entrada para el número de repeticiones
repeticiones_label = tk.Label(ventana, text="Número de repeticiones:")
repeticiones_label.pack()
repeticiones_entry = tk.Entry(ventana)
repeticiones_entry.pack()

# NOTAS #NOTAS # NOTAS # Etiqueta y campo de entrada para la ruta del archivo
ruta_label = tk.Label(ventana, text="Ruta del archivo:")
ruta_label.pack()
ruta_entry = tk.Entry(ventana)
ruta_entry.pack()

# NOTAS #NOTAS # NOTAS # Botón para ejecutar el código
ejecutar_boton = tk.Button(ventana, text="Ejecutar", command=ejecutar_codigo)
ejecutar_boton.pack()

# NOTAS #NOTAS # NOTAS # Etiqueta para mostrar el contador
contador_label = tk.Label(ventana, text="Contador: 0")
contador_label.pack()

# NOTAS #NOTAS # NOTAS # Iniciar el bucle de eventos
ventana.mainloop()
