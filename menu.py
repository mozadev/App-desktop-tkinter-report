from tkinter import *
from tkinter import ttk
import funciones
import reporte

ruta_averiaR = ""
ruta_bdR = ""
ruta_disponibilidadS = ""
ruta_tiempoR = ""
ruta_bdT = ""

def ventana_comp1():
    global ruta_averiaR, ruta_bdR, ruta_disponibilidadS, ruta_tiempoR
    if not hasattr(ventana_comp1, "boton1_enabled"):
    # Asegurarse de que el atributo existe
        ventana_comp1.boton1_enabled = True
    if ventana_comp1.boton1_enabled:
        v_componente1 = Toplevel(ventana)
        v_componente1.title("COMPONENTE 1")
        v_componente1.geometry("500x500")
        
        # Desactivar el botón después de abrir la ventana secundaria
        boton1["state"] = "disabled"
        
        # Variables para almacenar las rutas de los archivos
        global ruta_averiaR, ruta_bdR, ruta_disponibilidadS, ruta_tiempoR
        ruta_averiaR = StringVar()
        ruta_bdR = StringVar()
        ruta_disponibilidadS = StringVar()
        ruta_tiempoR = StringVar()
        
        # Cargar una imagen para el botón de subir archivo
        imagen_subir_archivo = PhotoImage(file="file.png")  # Reemplaza con la ruta de tu imagen

        # Crear un Frame para agrupar cada par de etiquetas (imagen y texto)

        frame1 = Frame(v_componente1)
        frame1.pack(pady=10)

        # Crear un Label para mostrar el nombre del archivo seleccionado
        title1 = Label(frame1, text="Seleciona el archivo Averias Reportadas:", font=('Helvetica', 10))
        title1.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        label_boton_imagen1 = Label(frame1, image=imagen_subir_archivo, cursor="hand2")
        label_boton_imagen1.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo1, ruta_averiaR))
        label_archivo1 = Label(frame1, text="Selecione el archivo:", font=('Helvetica', 10))
        
        # Repetir el proceso para el segundo par de etiquetas
        frame2 = Frame(v_componente1)
        frame2.pack(pady=10)

        # Crear un Label para mostrar el nombre del archivo seleccionado
        title2 = Label(frame2, text="Seleciona el archivo BD Reporte W-Minpu:", font=('Helvetica', 10))
        title2.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        label_boton_imagen2 = Label(frame2, image=imagen_subir_archivo, cursor="hand2")
        
        label_boton_imagen2.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo2, ruta_bdR))

        label_archivo2 = Label(frame2, text="Selecione el archivo:", font=('Helvetica', 10))
        # Repetir el proceso para el tercer par de etiquetas
        frame3 = Frame(v_componente1)
        frame3.pack(pady=10)

        # Crear un Label para mostrar el nombre del archivo seleccionado
        title3 = Label(frame3, text="Seleciona archivo Tiempo de Reparacion :", font=('Helvetica', 10))
        title3.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        label_boton_imagen3 = Label(frame3, image=imagen_subir_archivo, cursor="hand2")
        label_boton_imagen3.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo3, ruta_tiempoR))

        label_archivo3 = Label(frame3, text="Selecione el archivo:", font=('Helvetica', 10))
        
        # Repetir el proceso para el tercer par de etiquetas
        frame4 = Frame(v_componente1)
        frame4.pack(pady=10)

        # Crear un Label para mostrar el nombre del archivo seleccionado
        title4 = Label(frame4, text="Seleciona archivo Disponibilidad Servicio:", font=('Helvetica', 10))
        title4.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        label_boton_imagen4 = Label(frame4, image=imagen_subir_archivo, cursor="hand2")
        label_boton_imagen4.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo4, ruta_disponibilidadS ))

        label_archivo4 = Label(frame4, text="Selecione el archivo:", font=('Helvetica', 10))

        # Configurar la imagen para que no sea eliminada por el recolector de basura
        label_boton_imagen1.imagen = imagen_subir_archivo
        label_boton_imagen2.imagen = imagen_subir_archivo
        label_boton_imagen3.imagen = imagen_subir_archivo
        label_boton_imagen4.imagen = imagen_subir_archivo
        
        # Organizar los botones en dos columnas usando grid
        label_boton_imagen1.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")  # sticky para expandir tanto horizontal como verticalmente
        label_boton_imagen2.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        label_boton_imagen3.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
        label_boton_imagen4.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")

        label_archivo1.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")  # sticky para expandir tanto horizontal como verticalmente
        label_archivo2.grid(row=2, column=1, padx=10, pady=10, sticky="nsew")
        label_archivo3.grid(row=3, column=1, padx=10, pady=10, sticky="nsew")
        label_archivo4.grid(row=4, column=1, padx=10, pady=10, sticky="nsew")
        # Crear un objeto Style
        # Repetir el proceso para el tercer par de etiquetas
        frame5 = Frame(v_componente1)
        frame5.pack(pady=10)
        generar_reporte2 = Button(frame5, text="Componente 1 Datos", command=lambda: reporte.re_componente_1(ruta_averiaR.get(), ruta_bdR.get(), ruta_disponibilidadS.get(), ruta_tiempoR.get()), padx=10, pady=5, fg="white", bg="black")
        generar_reporte2.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="nsew") 
        
        v_componente1.protocol("WM_DELETE_WINDOW", lambda: on_close_comp1(v_componente1))

def on_close_comp1(sub_ventana):
    global boton1
    # Volver a activar el botón al cerrar la ventana secundaria
    boton1["state"] = "normal"
    sub_ventana.destroy()
    
def ventana_comp2():
    global ruta_averiaR, ruta_bdR, ruta_disponibilidadS, ruta_tiempoR
    if not hasattr(ventana_comp2, "boton2_enabled"):
    # Asegurarse de que el atributo existe
        ventana_comp2.boton2_enabled = True
    if ventana_comp2.boton2_enabled:
        v_componente2 = Toplevel(ventana)
        v_componente2.title("COMPONENTE 2")
        v_componente2.geometry("500x500")
        
        # Desactivar el botón después de abrir la ventana secundaria
        boton2["state"] = "disabled"
        
        # Variables para almacenar las rutas de los archivos
        global ruta_averiaR, ruta_bdR, ruta_disponibilidadS, ruta_tiempoR
        ruta_averiaR = StringVar()
        ruta_bdR = StringVar()
        ruta_disponibilidadS = StringVar()
        ruta_tiempoR = StringVar()
        
        # Cargar una imagen para el botón de subir archivo
        imagen_subir_archivo = PhotoImage(file="file.png")  # Reemplaza con la ruta de tu imagen

        # Crear un Frame para agrupar cada par de etiquetas (imagen y texto)

        frame1 = Frame(v_componente2)
        frame1.pack(pady=10)

        # Crear un Label para mostrar el nombre del archivo seleccionado
        title1 = Label(frame1, text="Seleciona el archivo Averias Reportadas:", font=('Helvetica', 10))
        title1.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        label_boton_imagen1 = Label(frame1, image=imagen_subir_archivo, cursor="hand2")
        label_boton_imagen1.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo1, ruta_averiaR))
        label_archivo1 = Label(frame1, text="Selecione el archivo:", font=('Helvetica', 10))
        
        # Repetir el proceso para el segundo par de etiquetas
        frame2 = Frame(v_componente2)
        frame2.pack(pady=10)

        # Crear un Label para mostrar el nombre del archivo seleccionado
        title2 = Label(frame2, text="Seleciona el archivo BD Reporte W-Minpu:", font=('Helvetica', 10))
        title2.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        label_boton_imagen2 = Label(frame2, image=imagen_subir_archivo, cursor="hand2")
        
        label_boton_imagen2.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo2, ruta_bdR))

        label_archivo2 = Label(frame2, text="Selecione el archivo:", font=('Helvetica', 10))
        # Repetir el proceso para el tercer par de etiquetas
        frame3 = Frame(v_componente2)
        frame3.pack(pady=10)

        # Crear un Label para mostrar el nombre del archivo seleccionado
        title3 = Label(frame3, text="Seleciona archivo Tiempo de Reparacion :", font=('Helvetica', 10))
        title3.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        label_boton_imagen3 = Label(frame3, image=imagen_subir_archivo, cursor="hand2")
        label_boton_imagen3.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo3, ruta_tiempoR))

        label_archivo3 = Label(frame3, text="Selecione el archivo:", font=('Helvetica', 10))
        
        # Repetir el proceso para el tercer par de etiquetas
        frame4 = Frame(v_componente2)
        frame4.pack(pady=10)

        # Crear un Label para mostrar el nombre del archivo seleccionado
        title4 = Label(frame4, text="Seleciona archivo Disponibilidad Servicio:", font=('Helvetica', 10))
        title4.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        label_boton_imagen4 = Label(frame4, image=imagen_subir_archivo, cursor="hand2")
        label_boton_imagen4.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo4, ruta_disponibilidadS ))

        label_archivo4 = Label(frame4, text="Selecione el archivo:", font=('Helvetica', 10))

        # Configurar la imagen para que no sea eliminada por el recolector de basura
        label_boton_imagen1.imagen = imagen_subir_archivo
        label_boton_imagen2.imagen = imagen_subir_archivo
        label_boton_imagen3.imagen = imagen_subir_archivo
        label_boton_imagen4.imagen = imagen_subir_archivo
        
        # Organizar los botones en dos columnas usando grid
        label_boton_imagen1.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")  # sticky para expandir tanto horizontal como verticalmente
        label_boton_imagen2.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        label_boton_imagen3.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
        label_boton_imagen4.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")

        label_archivo1.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")  # sticky para expandir tanto horizontal como verticalmente
        label_archivo2.grid(row=2, column=1, padx=10, pady=10, sticky="nsew")
        label_archivo3.grid(row=3, column=1, padx=10, pady=10, sticky="nsew")
        label_archivo4.grid(row=4, column=1, padx=10, pady=10, sticky="nsew")
        # Crear un objeto Style
        # Repetir el proceso para el tercer par de etiquetas
        frame5 = Frame(v_componente2)
        frame5.pack(pady=10)
        generar_reporte2 = Button(frame5, text="Componente 2 Datos", command=lambda: reporte.re_componente_2(ruta_averiaR.get(), ruta_bdR.get(), ruta_disponibilidadS.get(), ruta_tiempoR.get()), padx=10, pady=5, fg="white", bg="black")
        generar_reporte2.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="nsew") 
        
        v_componente2.protocol("WM_DELETE_WINDOW", lambda: on_close_comp2(v_componente2))

def on_close_comp2(sub_ventana):
    global boton2
    # Volver a activar el botón al cerrar la ventana secundaria
    boton2["state"] = "normal"
    sub_ventana.destroy()
    
def on_close_comp4(sub_ventana):
    global boton4
    # Volver a activar el botón al cerrar la ventana secundaria
    boton4["state"] = "normal"
    sub_ventana.destroy()


def ventana_comp4():
    global ruta_bdR, ruta_averiaR, ruta_tiempoR, ruta_bdSoli
    if not hasattr(ventana_comp4, "boton4_enabled"):
    # Asegurarse de que el atributo existe
        ventana_comp4.boton4_enabled = True
    if ventana_comp4.boton4_enabled:
        v_componente4 = Toplevel(ventana)
        v_componente4.title("COMPONENTE 4")
        v_componente4.geometry("500x500")
        
        # Desactivar el botón después de abrir la ventana secundaria
        boton4["state"] = "disabled"
        
        # Variables para almacenar las rutas de los archivos
        global ruta_bdR, ruta_averiaR, ruta_tiempoR, ruta_bdSoli
        ruta_bdR = StringVar()
        ruta_bdSoli = StringVar()
        ruta_averiaR = StringVar()
        # ruta_disponibilidadS = StringVar()
        ruta_tiempoR = StringVar()

        # Cargar una imagen para el botón de subir archivo
        imagen_subir_archivo = PhotoImage(file="file.png")  # Reemplaza con la ruta de tu imagen

        # Crear un Frame para agrupar cada par de etiquetas (imagen y texto)

        frame1 = Frame(v_componente4)
        frame1.pack(pady=10)

        # Crear un Label para mostrar el nombre del archivo seleccionado
        title1 = Label(frame1, text="Seleciona el archivo BD reporte Word :", font=('Helvetica', 10))
        title1.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        label_boton_imagen1 = Label(frame1, image=imagen_subir_archivo, cursor="hand2")
        label_boton_imagen1.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo1, ruta_bdR))
        label_archivo1 = Label(frame1, text="Selecione el archivo:", font=('Helvetica', 10))
        
        # Repetir el proceso para el segundo par de etiquetas
        frame2 = Frame(v_componente4)
        frame2.pack(pady=10)

        # Crear un Label para mostrar el nombre del archivo seleccionado
        title2 = Label(frame2, text="Seleciona el archivo BD tabla solicitudes:", font=('Helvetica', 10))
        title2.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        label_boton_imagen2 = Label(frame2, image=imagen_subir_archivo, cursor="hand2")
        
        label_boton_imagen2.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo2, ruta_bdSoli))

        label_archivo2 = Label(frame2, text="Selecione el archivo:", font=('Helvetica', 10))
        
        # Repetir el proceso para el tercer par de etiquetas
        frame3 = Frame(v_componente4)
        frame3.pack(pady=10)

        # Crear un Label para mostrar el nombre del archivo seleccionado
        title3 = Label(frame3, text="Seleciona archivo Tiempo de Reparacion :", font=('Helvetica', 10))
        title3.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        label_boton_imagen3 = Label(frame3, image=imagen_subir_archivo, cursor="hand2")
        label_boton_imagen3.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo3, ruta_tiempoR))

        label_archivo3 = Label(frame3, text="Selecione el archivo:", font=('Helvetica', 10))

        # # Repetir el proceso para el tercer par de etiquetas
        # frame4 = Frame(v_componente4)
        # frame4.pack(pady=10)

        # # Crear un Label para mostrar el nombre del archivo seleccionado
        # title4 = Label(frame4, text="Seleciona archivo Disponibilidad Servicio:", font=('Helvetica', 10))
        # title4.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        # label_boton_imagen4 = Label(frame4, image=imagen_subir_archivo, cursor="hand2")
        # label_boton_imagen4.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo4, ruta_disponibilidadS ))

        # label_archivo4 = Label(frame4, text="Selecione el archivo:", font=('Helvetica', 10))
        
        # Repetir el proceso para el tercer par de etiquetas
        frame4 = Frame(v_componente4)
        frame4.pack(pady=10)

        # Crear un Label para mostrar el nombre del archivo seleccionado
        title4 = Label(frame4, text="Seleciona el archivo Averias Reportadas:", font=('Helvetica', 10))
        title4.grid(row=0, column=0, columnspan=2)  # Ocupa dos columnas
        
        label_boton_imagen4 = Label(frame4, image=imagen_subir_archivo, cursor="hand2")
        label_boton_imagen4.bind("<Button-1>", lambda event: funciones.seleccionar_archivo(label_archivo4, ruta_averiaR ))

        label_archivo4 = Label(frame4, text="Selecione el archivo:", font=('Helvetica', 10))

        # Configurar la imagen para que no sea eliminada por el recolector de basura
        label_boton_imagen1.imagen = imagen_subir_archivo
        label_boton_imagen2.imagen = imagen_subir_archivo
        label_boton_imagen3.imagen = imagen_subir_archivo
        # label_boton_imagen4.imagen = imagen_subir_archivo
        label_boton_imagen4.imagen = imagen_subir_archivo
        
        # Organizar los botones en dos columnas usando grid
        label_boton_imagen1.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")  # sticky para expandir tanto horizontal como verticalmente
        label_boton_imagen2.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        label_boton_imagen3.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
        # label_boton_imagen4.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")
        label_boton_imagen4.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")

        label_archivo1.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")  # sticky para expandir tanto horizontal como verticalmente
        label_archivo2.grid(row=2, column=1, padx=10, pady=10, sticky="nsew")
        label_archivo3.grid(row=3, column=1, padx=10, pady=10, sticky="nsew")
        # label_archivo4.grid(row=4, column=1, padx=10, pady=10, sticky="nsew")
        label_archivo4.grid(row=4, column=1, padx=10, pady=10, sticky="nsew")
        # Crear un objeto Style
        # Repetir el proceso para el tercer par de etiquetas
        frame5 = Frame(v_componente4)
        frame5.pack(pady=10)
        generar_reporte2 = Button(frame5, text="Componente 4 Telfonos", command=lambda: reporte.re_componente_4( ruta_bdR.get(), ruta_averiaR.get(), ruta_tiempoR.get(), ruta_bdSoli.get()), padx=10, pady=5, fg="white", bg="black")
        generar_reporte2.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        v_componente4.protocol("WM_DELETE_WINDOW", lambda: on_close_comp4(v_componente4))
    

#------------------------------------------------------------------------------------------#

#INTERFAZ PRINCIPAL
# Crear la ventana principal
ventana = Tk()
# Establecer el tamaño de la ventana a 800x400 píxeles
ventana.geometry("500x400")
ventana.title("GENERADOR DE REPORTE MINPU")

# Crear un Label para acompañar los botones
etiqueta = Label(ventana, text="Escoge el reporte que deseas generar:", font=('Helvetica', 12))
etiqueta.grid(row=0, column=0, columnspan=2, pady=10)  # Ocupa dos columnas

# Crear un objeto Style
style = ttk.Style()
# Definir un estilo para los botones
style.configure("TButton", padding=10, font=('Helvetica', 12), padx=10, pady=5, fg="white")

# Crear los botones y asignar funciones
boton1 = ttk.Button(ventana, text="Componente 1 Internet", command=ventana_comp1)
boton2 = ttk.Button(ventana, text="Componente 2 Datos", command=ventana_comp2)
boton3 = ttk.Button(ventana, text="Componente 3")
boton4 = ttk.Button(ventana, text="Componente 4", command=ventana_comp4)
boton5 = ttk.Button(ventana, text="Componente 5")

# Organizar los botones en dos columnas usando grid
boton1.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")  # sticky para expandir tanto horizontal como verticalmente
boton2.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")
boton3.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
boton4.grid(row=2, column=1, padx=10, pady=10, sticky="nsew")
boton5.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")  # Para que el quinto botón ocupe dos columnas

# Configurar las columnas y filas para expandirse
ventana.grid_columnconfigure(0, weight=1)
ventana.grid_columnconfigure(1, weight=1)
ventana.grid_rowconfigure(1, weight=1)
ventana.grid_rowconfigure(2, weight=1)
ventana.grid_rowconfigure(3, weight=1)


# Iniciar el bucle principal de la interfaz gráfica
ventana.mainloop()