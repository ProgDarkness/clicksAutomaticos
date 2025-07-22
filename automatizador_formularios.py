import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pyautogui
import time
import json
import os
import threading
import keyboard
import pytesseract
from PIL import Image, ImageOps, ImageTk  
import sys

# Configuración de pytesseract
try:
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
except Exception as e:
    print(f"Error configurando Tesseract: {e}")

# Configuración básica
pyautogui.PAUSE = 0.5
pyautogui.FAILSAFE = True
CONFIG_FILE = "automatizacion_config.json"

class AutomatizadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatizador Avanzado v2.0")
        self.root.geometry("1000x700")
        self.ejecucion_automatica = False
        self.intervalo = 30
        self.acciones = []
        self.capturando_posicion = False
        self.capturando_region = False
        self.posicion_actual = None
        self.region_actual = None
        
        # Estilo
        self.setup_estilos()
        
        # Marco principal
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Paneles
        self.setup_panel_acciones()
        self.setup_panel_ejecucion()
        self.setup_panel_registro()
        
        # Cargar configuración inicial
        self.cargar_configuracion()
        
        # Bindear tecla SUPR
        self.root.bind('<Delete>', lambda e: self.detener_ejecucion())

    def setup_estilos(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('.', background='#f0f0f0')
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TButton', padding=6, font=('Segoe UI', 9))
        style.configure('TLabel', background='#f0f0f0', font=('Segoe UI', 9))
        style.configure('Header.TLabel', font=('Segoe UI', 11, 'bold'))
        style.configure('Treeview', rowheight=25)
        style.map('TButton', background=[('active', '#e0e0e0')])

    def setup_panel_acciones(self):
        frame = ttk.LabelFrame(self.main_frame, text="Administrar Acciones", padding=10)
        frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        # Botones de acciones
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=0, column=0, pady=5, sticky="ew")
        
        ttk.Button(btn_frame, text="Agregar Simple", command=self.agregar_accion_simple, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Agregar Condicional", command=self.agregar_accion_condicional, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Editar", command=self.editar_accion, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Eliminar", command=self.eliminar_accion, width=15).pack(side=tk.LEFT, padx=2)
        
        # Lista de acciones
        self.lista_acciones = ttk.Treeview(frame, columns=('tipo', 'detalles'), show='headings', height=12)
        self.lista_acciones.heading('#0', text='Nombre')
        self.lista_acciones.heading('tipo', text='Tipo')
        self.lista_acciones.heading('detalles', text='Detalles')
        self.lista_acciones.column('#0', width=200)
        self.lista_acciones.column('tipo', width=120)
        self.lista_acciones.column('detalles', width=300)
        self.lista_acciones.grid(row=1, column=0, pady=10, sticky="nsew")
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.lista_acciones.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.lista_acciones.configure(yscrollcommand=scrollbar.set)
        
        # Configurar grid para expansión
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

    def setup_panel_ejecucion(self):
        frame = ttk.LabelFrame(self.main_frame, text="Control de Ejecución", padding=10)
        frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        
        # Estado
        status_frame = ttk.Frame(frame)
        status_frame.grid(row=0, column=0, columnspan=2, pady=5, sticky="ew")
        ttk.Label(status_frame, text="Estado:", style='Header.TLabel').pack(side=tk.LEFT)
        self.estado_label = ttk.Label(status_frame, text="Listo", foreground="green")
        self.estado_label.pack(side=tk.LEFT, padx=10)
        
        # Botones de control
        ttk.Button(frame, text="Ejecutar Una Vez", command=self.ejecutar_automatizacion).grid(row=1, column=0, columnspan=2, pady=5, sticky="ew")
        ttk.Button(frame, text="Iniciar Automático", command=self.iniciar_automatico).grid(row=2, column=0, columnspan=2, pady=5, sticky="ew")
        ttk.Button(frame, text="Detener", command=self.detener_ejecucion).grid(row=3, column=0, columnspan=2, pady=5, sticky="ew")
        
        # Configuración
        config_frame = ttk.Frame(frame)
        config_frame.grid(row=4, column=0, columnspan=2, pady=10, sticky="ew")
        
        ttk.Label(config_frame, text="Intervalo (seg):").pack(side=tk.LEFT)
        self.intervalo_entry = ttk.Entry(config_frame, width=10)
        self.intervalo_entry.pack(side=tk.LEFT, padx=5)
        self.intervalo_entry.insert(0, "30")
        
        # Captura de posición/región
        ttk.Button(frame, text="Capturar Posición", command=self.iniciar_captura_posicion).grid(row=5, column=0, pady=5, sticky="ew")
        ttk.Button(frame, text="Capturar Región", command=self.iniciar_captura_region).grid(row=5, column=1, pady=5, sticky="ew")
        
        self.posicion_label = ttk.Label(frame, text="Posición: No capturada")
        self.posicion_label.grid(row=6, column=0, columnspan=2, sticky="w")
        
        # Configurar grid para expansión
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

    def setup_panel_registro(self):
        frame = ttk.LabelFrame(self.main_frame, text="Registro de Actividad", padding=10)
        frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        
        self.registro_text = tk.Text(frame, height=12, wrap=tk.WORD, font=('Consolas', 9))
        self.registro_text.grid(row=0, column=0, sticky="nsew")
        
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.registro_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.registro_text.configure(yscrollcommand=scrollbar.set)
        
        # Configurar grid para expansión
        self.main_frame.rowconfigure(1, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def log(self, mensaje):
        timestamp = time.strftime("%H:%M:%S")
        self.registro_text.insert(tk.END, f"[{timestamp}] {mensaje}\n")
        self.registro_text.see(tk.END)
        self.root.update()

    def cargar_configuracion(self):
        """Carga la configuración desde el archivo JSON"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    self.acciones = json.load(f)
            else:
                self.acciones = []
                # Crear archivo vacío si no existe
                with open(CONFIG_FILE, 'w') as f:
                    json.dump([], f)
            
            self.actualizar_lista_acciones()
            self.log("Configuración cargada correctamente")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar la configuración: {str(e)}")
            self.acciones = []

    def guardar_configuracion_automatica(self):
        """Guarda automáticamente la configuración en el archivo JSON"""
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(self.acciones, f, indent=4, ensure_ascii=False)
            self.log("Configuración guardada automáticamente")
        except Exception as e:
            self.log(f"Error al guardar automáticamente: {str(e)}")

    def actualizar_lista_acciones(self):
        self.lista_acciones.delete(*self.lista_acciones.get_children())
        for accion in self.acciones:
            detalles = self.obtener_detalles_accion(accion)
            self.lista_acciones.insert('', 'end', text=accion['nombre'], 
                                     values=(accion['tipo'], detalles))

    def obtener_detalles_accion(self, accion):
        if accion['tipo'] == "click":
            return f"X: {accion['posicion']['x']}, Y: {accion['posicion']['y']}"
        elif accion['tipo'] == "escribir":
            return f"Texto: '{accion['texto']}'"
        elif accion['tipo'] == "tecla":
            return f"Tecla: {accion['tecla']}"
        elif accion['tipo'] == "esperar_texto":
            region = f" en región {accion['region']}" if accion.get('region') else ""
            return f"Texto: '{accion['texto']}'{region}"
        elif accion['tipo'] == "condicional":
            return f"Si '{accion['texto_condicion']}'"
        elif accion['tipo'] == "terminar":
            return "Finaliza el proceso"
        return ""

    # Funciones para captura de posición/región
    def iniciar_captura_posicion(self):
        self.capturando_posicion = True
        self.posicion_actual = None
        self.estado_label.config(text="Capturando posición...", foreground="blue")
        self.log("Mueve el mouse a la posición deseada y presiona 's' para capturar")
        self.root.after(100, self.verificar_captura_posicion)

    def verificar_captura_posicion(self):
        if self.capturando_posicion:
            if keyboard.is_pressed('s'):
                self.posicion_actual = pyautogui.position()
                self.capturando_posicion = False
                self.estado_label.config(text="Listo", foreground="green")
                self.posicion_label.config(text=f"Posición: X={self.posicion_actual.x}, Y={self.posicion_actual.y}")
                self.log(f"Posición capturada: {self.posicion_actual}")
            else:
                self.root.after(100, self.verificar_captura_posicion)

    def iniciar_captura_region(self):
        """Inicia el proceso de captura de una región rectangular"""
        self.capturando_region = True
        self.region_actual = None
        self.esquina_superior = None
        self.esquina_inferior = None
    
        self.estado_label.config(text="Capturando región...", foreground="blue")
        self.log("Captura de región iniciada")
        self.log("Paso 1: Mueve el mouse a la esquina superior izquierda y presiona 's'")
    
        # Cambiar el cursor para indicar modo de captura
        self.root.config(cursor="crosshair")
    
        # Iniciar el proceso de captura
        self.root.after(100, self.capturar_esquina_superior)
        
    def capturar_esquina_superior(self):
        if not self.capturando_region:
            return
    
        if keyboard.is_pressed('s'):
            self.esquina_superior = pyautogui.position()
            self.log(f"Esquina superior izquierda capturada: {self.esquina_superior}")
            self.log("Paso 2: Mueve el mouse a la esquina inferior derecha y presiona 's'")
            self.root.after(100, self.capturar_esquina_inferior)
        elif keyboard.is_pressed('escape'):
            self.cancelar_captura_region()
        else:
            self.root.after(100, self.capturar_esquina_superior)
    
    def capturar_esquina_inferior(self):
        if not self.capturando_region:
            return
    
        if keyboard.is_pressed('s'):
            self.esquina_inferior = pyautogui.position()
        
            # Asegurarse de que las coordenadas sean válidas
            x1 = min(self.esquina_superior.x, self.esquina_inferior.x)
            y1 = min(self.esquina_superior.y, self.esquina_inferior.y)
            x2 = max(self.esquina_superior.x, self.esquina_inferior.x)
            y2 = max(self.esquina_superior.y, self.esquina_inferior.y)
        
            width = x2 - x1
            height = y2 - y1
        
            # Validar tamaño mínimo
            if width < 10 or height < 10:
                self.log("Región demasiado pequeña, debe ser al menos 10x10 píxeles")
                self.cancelar_captura_region()
                return
        
            self.region_actual = (x1, y1, width, height)
            self.finalizar_captura_region()
        elif keyboard.is_pressed('escape'):
            self.cancelar_captura_region()
        else:
            self.root.after(100, self.capturar_esquina_inferior)
    
    def finalizar_captura_region(self):
        """Finaliza el proceso de captura exitosamente"""
        self.capturando_region = False
        self.estado_label.config(text="Listo", foreground="green")
        self.posicion_label.config(text=f"Región capturada: {self.region_actual}")
        self.log(f"Región capturada exitosamente: {self.region_actual}")
        self.root.config(cursor="")
    
        # Dibujar un rectángulo temporal para visualización
        self.mostrar_preview_region()
        
    def cancelar_captura_region(self):
        """Cancela el proceso de captura"""
        self.capturando_region = False
        self.region_actual = None
        self.esquina_superior = None
        self.esquina_inferior = None
        self.estado_label.config(text="Listo", foreground="green")
        self.posicion_label.config(text="Región: No capturada")
        self.log("Captura de región cancelada")
        self.root.config(cursor="")
    
    def mostrar_preview_region(self):
        """Muestra un preview visual de la región capturada"""
        try:
            if not self.region_actual:
                return
        
            x, y, w, h = self.region_actual
            screenshot = pyautogui.screenshot(region=self.region_actual)
        
            # Crear ventana de preview
            preview = tk.Toplevel(self.root)
            preview.title("Vista Previa de Región")
        
            # Convertir imagen para Tkinter
            tk_image = ImageTk.PhotoImage(screenshot)
        
            # Mostrar imagen
            label = tk.Label(preview, image=tk_image)
            label.image = tk_image  # Mantener referencia
            label.pack()
        
            # Mostrar coordenadas
            coords_label = tk.Label(preview, text=f"X: {x}, Y: {y}, Ancho: {w}, Alto: {h}")
            coords_label.pack()
        
            # Botón para aceptar
            ttk.Button(preview, text="Aceptar", command=preview.destroy).pack(pady=5)
        
        except Exception as e:
            self.log(f"Error al mostrar preview: {str(e)}")
    
    def verificar_captura_region_paso1(self):
        if self.capturando_region:
            if keyboard.is_pressed('s'):
                self.esquina_superior = pyautogui.position()
                self.capturando_region = True
                self.log(f"Esquina superior izquierda: {self.esquina_superior}")
                self.log("Mueve el mouse a la esquina inferior derecha y presiona 's'")
                self.root.after(100, self.verificar_captura_region_paso2)
            else:
                self.root.after(100, self.verificar_captura_region_paso1)

    def verificar_captura_region_paso2(self):
        if self.capturando_region:
            if keyboard.is_pressed('s'):
                esquina_inferior = pyautogui.position()
                self.region_actual = (
                    self.esquina_superior.x, 
                    self.esquina_superior.y,
                    esquina_inferior.x - self.esquina_superior.x,
                    esquina_inferior.y - self.esquina_superior.y
                )
                self.capturando_region = False
                self.estado_label.config(text="Listo", foreground="green")
                self.posicion_label.config(text=f"Región: {self.region_actual}")
                self.log(f"Región capturada: {self.region_actual}")
            else:
                self.root.after(100, self.verificar_captura_region_paso2)

    # Funciones para manejar acciones
    def agregar_accion_simple(self):
        dialogo = tk.Toplevel(self.root)
        dialogo.title("Agregar Acción Simple")
        dialogo.geometry("400x300")
        
        ttk.Label(dialogo, text="Tipo de acción:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        tipo_var = tk.StringVar()
        tipos = ["click", "escribir", "tecla", "esperar_texto", "terminar"]
        ttk.Combobox(dialogo, textvariable=tipo_var, values=tipos, state="readonly").grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(dialogo, text="Nombre:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        nombre_entry = ttk.Entry(dialogo)
        nombre_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        # Campos específicos
        campos_frame = ttk.Frame(dialogo)
        campos_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        
        def actualizar_campos(*args):
            for widget in campos_frame.winfo_children():
                widget.destroy()
            
            tipo = tipo_var.get()
            
            if tipo == "click":
                ttk.Button(campos_frame, text="Capturar Posición", command=self.capturar_posicion_desde_dialogo).pack(pady=5)
            elif tipo == "escribir":
                ttk.Label(campos_frame, text="Texto:").pack(anchor="w")
                texto_entry = ttk.Entry(campos_frame)
                texto_entry.pack(fill="x", pady=5)
                ttk.Button(campos_frame, text="Capturar Posición", command=self.capturar_posicion_desde_dialogo).pack(pady=5)
            elif tipo == "tecla":
                ttk.Label(campos_frame, text="Tecla:").pack(anchor="w")
                ttk.Entry(campos_frame).pack(fill="x", pady=5)
            elif tipo == "esperar_texto":
                ttk.Label(campos_frame, text="Texto a esperar:").pack(anchor="w")
                ttk.Entry(campos_frame).pack(fill="x", pady=5)
                ttk.Label(campos_frame, text="Timeout (seg):").pack(anchor="w")
                ttk.Entry(campos_frame).pack(fill="x", pady=5)
                ttk.Label(campos_frame, text="Umbral (0-255):").pack(anchor="w")
                ttk.Entry(campos_frame).pack(fill="x", pady=5)
                ttk.Button(campos_frame, text="Capturar Región (opcional)", command=self.capturar_region_desde_dialogo).pack(pady=5)
        
        tipo_var.trace("w", actualizar_campos)
        tipo_var.set("click")  # Valor por defecto
        
        def guardar_accion():
            tipo = tipo_var.get()
            nombre = nombre_entry.get()
            
            if not nombre:
                messagebox.showerror("Error", "Debe ingresar un nombre para la acción")
                return
            
            accion = {"tipo": tipo, "nombre": nombre}
            
            if tipo == "click":
                if not self.posicion_actual:
                    messagebox.showerror("Error", "Debe capturar una posición primero")
                    return
                accion["posicion"] = {"x": self.posicion_actual.x, "y": self.posicion_actual.y}
            elif tipo == "escribir":
                texto = campos_frame.winfo_children()[1].get()
                if not texto:
                    messagebox.showerror("Error", "Debe ingresar un texto")
                    return
                if not self.posicion_actual:
                    messagebox.showerror("Error", "Debe capturar una posición primero")
                    return
                accion["texto"] = texto
                accion["posicion"] = {"x": self.posicion_actual.x, "y": self.posicion_actual.y}
            elif tipo == "tecla":
                tecla = campos_frame.winfo_children()[1].get()
                if not tecla:
                    messagebox.showerror("Error", "Debe ingresar una tecla")
                    return
                accion["tecla"] = tecla
            elif tipo == "esperar_texto":
                texto = campos_frame.winfo_children()[1].get()
                timeout = campos_frame.winfo_children()[3].get()
                umbral = campos_frame.winfo_children()[5].get()
                
                if not texto:
                    messagebox.showerror("Error", "Debe ingresar un texto a esperar")
                    return
                
                accion["texto"] = texto
                accion["timeout"] = int(timeout) if timeout.isdigit() else 30
                accion["umbral"] = int(umbral) if umbral.isdigit() else 160
                
                if hasattr(self, 'region_dialogo_actual'):
                    accion["region"] = self.region_dialogo_actual
            elif tipo == "terminar":
                pass  # No necesita parámetros adicionales
            
            self.acciones.append(accion)
            self.actualizar_lista_acciones()
            self.guardar_configuracion_automatica()
            dialogo.destroy()
            self.log(f"Acción '{nombre}' agregada")
        
        ttk.Button(dialogo, text="Guardar", command=guardar_accion).grid(row=3, column=0, columnspan=2, pady=10)
        
        dialogo.transient(self.root)
        dialogo.grab_set()
        self.root.wait_window(dialogo)

    def capturar_posicion_desde_dialogo(self):
        self.iniciar_captura_posicion()
        messagebox.showinfo("Capturar Posición", "Mueve el mouse a la posición deseada y presiona 's' para capturar")

    def capturar_region_desde_dialogo(self):
        self.iniciar_captura_region()
        messagebox.showinfo("Capturar Región", "1. Mueve a la esquina superior izquierda y presiona 's'\n2. Mueve a la esquina inferior derecha y presiona 's'")

    def agregar_accion_condicional(self):
        dialogo = tk.Toplevel(self.root)
        dialogo.title("Agregar Acción Condicional")
        dialogo.geometry("260x460")
        
        # Configuración de la condición
        ttk.Label(dialogo, text="Nombre de la acción:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        nombre_entry = ttk.Entry(dialogo)
        nombre_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(dialogo, text="Texto a detectar:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        texto_entry = ttk.Entry(dialogo)
        texto_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(dialogo, text="Timeout (seg):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        timeout_entry = ttk.Entry(dialogo)
        timeout_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        timeout_entry.insert(0, "10")
        
        ttk.Button(dialogo, text="Capturar Región (opcional)", command=self.capturar_region_desde_dialogo).grid(row=3, column=0, columnspan=2, pady=5)
        
        # Pestañas para acciones SI/NO
        notebook = ttk.Notebook(dialogo)
        notebook.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        
        frame_si = ttk.Frame(notebook)
        frame_no = ttk.Frame(notebook)
        
        notebook.add(frame_si, text="Acciones SI")
        notebook.add(frame_no, text="Acciones NO")
        
        # Listas de acciones
        self.lista_acciones_si = ttk.Treeview(frame_si, columns=('tipo'), show='headings', height=5)
        self.lista_acciones_si.heading('#0', text='Nombre')
        self.lista_acciones_si.heading('tipo', text='Tipo')
        self.lista_acciones_si.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.lista_acciones_no = ttk.Treeview(frame_no, columns=('tipo'), show='headings', height=5)
        self.lista_acciones_no.heading('#0', text='Nombre')
        self.lista_acciones_no.heading('tipo', text='Tipo')
        self.lista_acciones_no.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Botones para agregar acciones
        ttk.Button(frame_si, text="Agregar Acción", command=lambda: self.agregar_accion_a_lista(self.lista_acciones_si)).pack(pady=5)
        ttk.Button(frame_no, text="Agregar Acción", command=lambda: self.agregar_accion_a_lista(self.lista_acciones_no)).pack(pady=5)
        
        def guardar_accion():
            nombre = nombre_entry.get()
            texto = texto_entry.get()
            
            if not nombre or not texto:
                messagebox.showerror("Error", "Debe completar todos los campos")
                return
            
            # Obtener acciones SI
            acciones_si = []
            for item in self.lista_acciones_si.get_children():
                nombre_accion = self.lista_acciones_si.item(item, 'text')
                for accion in self.acciones:
                    if accion['nombre'] == nombre_accion:
                        acciones_si.append(accion)
                        break
            
            # Obtener acciones NO
            acciones_no = []
            for item in self.lista_acciones_no.get_children():
                nombre_accion = self.lista_acciones_no.item(item, 'text')
                for accion in self.acciones:
                    if accion['nombre'] == nombre_accion:
                        acciones_no.append(accion)
                        break
            
            accion_condicional = {
                "tipo": "condicional",
                "nombre": nombre,
                "texto_condicion": texto,
                "timeout": int(timeout_entry.get()) if timeout_entry.get().isdigit() else 10,
                "acciones_si": acciones_si,
                "acciones_no": acciones_no
            }
            
            if hasattr(self, 'region_dialogo_actual'):
                accion_condicional["region"] = self.region_dialogo_actual
            
            self.acciones.append(accion_condicional)
            self.actualizar_lista_acciones()
            self.guardar_configuracion_automatica()
            dialogo.destroy()
            self.log(f"Acción condicional '{nombre}' agregada")
        
        ttk.Button(dialogo, text="Guardar", command=guardar_accion).grid(row=5, column=0, columnspan=2, pady=10)
        
        dialogo.transient(self.root)
        dialogo.grab_set()
        self.root.wait_window(dialogo)

    def agregar_accion_a_lista(self, lista):
        dialogo = tk.Toplevel(self.root)
        dialogo.title("Agregar Acción")
        dialogo.geometry("500x400")
        
        notebook = ttk.Notebook(dialogo)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Pestaña para seleccionar existentes
        tab_existentes = ttk.Frame(notebook)
        notebook.add(tab_existentes, text="Usar Existente")
        
        # Pestaña para crear nueva
        tab_nueva = ttk.Frame(notebook)
        notebook.add(tab_nueva, text="Crear Nueva")
        
        # Contenido pestaña existentes
        nombres_acciones = [accion['nombre'] for accion in self.acciones if accion['tipo'] != "condicional"]
        listbox = tk.Listbox(tab_existentes)
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        for nombre in nombres_acciones:
            listbox.insert(tk.END, nombre)
        
        # Contenido pestaña nueva
        self.setup_panel_crear_accion(tab_nueva)
        
        def seleccionar():
            # Determinar qué pestaña está activa
            if notebook.index(notebook.select()) == 0:  # Pestaña existentes
                seleccionados = listbox.curselection()
                if seleccionados:
                    nombre = listbox.get(seleccionados[0])
                    for accion in self.acciones:
                        if accion['nombre'] == nombre:
                            lista.insert('', 'end', text=accion['nombre'], 
                                    values=(accion['tipo'], self.obtener_detalles_accion(accion)))
                            break
            else:  # Pestaña nueva
                accion = self.obtener_accion_desde_panel(tab_nueva)
                if accion:
                    self.acciones.append(accion)
                    lista.insert('', 'end', text=accion['nombre'], 
                            values=(accion['tipo'], self.obtener_detalles_accion(accion)))
                    self.guardar_configuracion_automatica()  # <-- Nueva línea
            
            dialogo.destroy()
        
        ttk.Button(dialogo, text="Agregar", command=seleccionar).pack(pady=10)
        dialogo.transient(self.root)
        dialogo.grab_set()
        self.root.wait_window(dialogo)

    def setup_panel_crear_accion(self, parent):
        # Similar a tu método agregar_accion_simple pero dentro de este panel
        ttk.Label(parent, text="Tipo de acción:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.tipo_var = tk.StringVar()
        tipos = ["click", "escribir", "tecla", "esperar_texto", "terminar"]
        ttk.Combobox(parent, textvariable=self.tipo_var, values=tipos, state="readonly").grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(parent, text="Nombre:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.nombre_entry = ttk.Entry(parent)
        self.nombre_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        # Frame para campos dinámicos
        self.campos_frame = ttk.Frame(parent)
        self.campos_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        
        self.tipo_var.trace("w", self.actualizar_campos_accion)
        self.tipo_var.set("click")

    def actualizar_campos_accion(self, *args):
        for widget in self.campos_frame.winfo_children():
            widget.destroy()
        
        tipo = self.tipo_var.get()
        
        if tipo == "click":
            ttk.Button(self.campos_frame, text="Capturar Posición", command=self.capturar_posicion_desde_dialogo).pack(pady=5)
        elif tipo == "escribir":
            ttk.Label(self.campos_frame, text="Texto:").pack(anchor="w")
            self.texto_entry = ttk.Entry(self.campos_frame)
            self.texto_entry.pack(fill="x", pady=5)
            ttk.Button(self.campos_frame, text="Capturar Posición", command=self.capturar_posicion_desde_dialogo).pack(pady=5)
        elif tipo == "tecla":
            ttk.Label(self.campos_frame, text="Tecla:").pack(anchor="w")
            self.tecla_entry = ttk.Entry(self.campos_frame)
            self.tecla_entry.pack(fill="x", pady=5)
        elif tipo == "esperar_texto":
            ttk.Label(self.campos_frame, text="Texto a esperar:").pack(anchor="w")
            self.texto_esperar_entry = ttk.Entry(self.campos_frame)
            self.texto_esperar_entry.pack(fill="x", pady=5)
            ttk.Label(self.campos_frame, text="Timeout (seg):").pack(anchor="w")
            self.timeout_entry = ttk.Entry(self.campos_frame)
            self.timeout_entry.pack(fill="x", pady=5)
            ttk.Label(self.campos_frame, text="Umbral (0-255):").pack(anchor="w")
            self.umbral_entry = ttk.Entry(self.campos_frame)
            self.umbral_entry.pack(fill="x", pady=5)
            ttk.Button(self.campos_frame, text="Capturar Región (opcional)", command=self.capturar_region_desde_dialogo).pack(pady=5)

    def obtener_accion_desde_panel(self, parent):
        nombre = self.nombre_entry.get()
        tipo = self.tipo_var.get()
        
        if not nombre:
            messagebox.showerror("Error", "Debe ingresar un nombre para la acción")
            return None
        
        accion = {"tipo": tipo, "nombre": nombre}
        
        if tipo == "click":
            if not hasattr(self, 'posicion_actual') or not self.posicion_actual:
                messagebox.showerror("Error", "Debe capturar una posición primero")
                return None
            accion["posicion"] = {"x": self.posicion_actual.x, "y": self.posicion_actual.y}
        elif tipo == "escribir":
            texto = self.texto_entry.get()
            if not texto:
                messagebox.showerror("Error", "Debe ingresar un texto")
                return None
            if not hasattr(self, 'posicion_actual') or not self.posicion_actual:
                messagebox.showerror("Error", "Debe capturar una posición primero")
                return None
            accion["texto"] = texto
            accion["posicion"] = {"x": self.posicion_actual.x, "y": self.posicion_actual.y}
        elif tipo == "tecla":
            tecla = self.tecla_entry.get()
            if not tecla:
                messagebox.showerror("Error", "Debe ingresar una tecla")
                return None
            accion["tecla"] = tecla
        elif tipo == "esperar_texto":
            texto = self.texto_esperar_entry.get()
            timeout = self.timeout_entry.get()
            umbral = self.umbral_entry.get()
            
            if not texto:
                messagebox.showerror("Error", "Debe ingresar un texto a esperar")
                return None
            
            accion["texto"] = texto
            accion["timeout"] = int(timeout) if timeout.isdigit() else 30
            accion["umbral"] = int(umbral) if umbral.isdigit() else 160
            
            if hasattr(self, 'region_dialogo_actual'):
                accion["region"] = self.region_dialogo_actual
        elif tipo == "terminar":
            pass  # No necesita parámetros adicionales
        
        return accion

    def editar_accion(self):
        seleccion = self.lista_acciones.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione una acción para editar")
            return
        
        item = seleccion[0]
        nombre = self.lista_acciones.item(item, 'text')
        
        for i, accion in enumerate(self.acciones):
            if accion['nombre'] == nombre:
                if accion['tipo'] == "condicional":
                    messagebox.showinfo("Información", "Las acciones condicionales deben eliminarse y crearse de nuevo para editarlas")
                else:
                    self.editar_accion_simple(i, accion)
                break

    def editar_accion_simple(self, index, accion):
        dialogo = tk.Toplevel(self.root)
        dialogo.title("Editar Acción")
        dialogo.geometry("400x300")
        
        ttk.Label(dialogo, text="Nombre:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        nombre_entry = ttk.Entry(dialogo)
        nombre_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        nombre_entry.insert(0, accion['nombre'])
        
        # Mostrar campos según tipo de acción
        campos_frame = ttk.Frame(dialogo)
        campos_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        
        if accion['tipo'] == "click":
            ttk.Label(campos_frame, text=f"Posición actual: X={accion['posicion']['x']}, Y={accion['posicion']['y']}").pack(anchor="w")
            ttk.Button(campos_frame, text="Cambiar Posición", command=self.capturar_posicion_desde_dialogo).pack(pady=5)
        elif accion['tipo'] == "escribir":
            ttk.Label(campos_frame, text="Texto:").pack(anchor="w")
            texto_entry = ttk.Entry(campos_frame)
            texto_entry.pack(fill="x", pady=5)
            texto_entry.insert(0, accion['texto'])
            ttk.Label(campos_frame, text=f"Posición actual: X={accion['posicion']['x']}, Y={accion['posicion']['y']}").pack(anchor="w")
            ttk.Button(campos_frame, text="Cambiar Posición", command=self.capturar_posicion_desde_dialogo).pack(pady=5)
        elif accion['tipo'] == "tecla":
            ttk.Label(campos_frame, text="Tecla:").pack(anchor="w")
            tecla_entry = ttk.Entry(campos_frame)
            tecla_entry.pack(fill="x", pady=5)
            tecla_entry.insert(0, accion['tecla'])
        elif accion['tipo'] == "esperar_texto":
            ttk.Label(campos_frame, text="Texto a esperar:").pack(anchor="w")
            texto_entry = ttk.Entry(campos_frame)
            texto_entry.pack(fill="x", pady=5)
            texto_entry.insert(0, accion['texto'])
            
            ttk.Label(campos_frame, text="Timeout (seg):").pack(anchor="w")
            timeout_entry = ttk.Entry(campos_frame)
            timeout_entry.pack(fill="x", pady=5)
            timeout_entry.insert(0, str(accion.get('timeout', 30)))
            
            ttk.Label(campos_frame, text="Umbral (0-255):").pack(anchor="w")
            umbral_entry = ttk.Entry(campos_frame)
            umbral_entry.pack(fill="x", pady=5)
            umbral_entry.insert(0, str(accion.get('umbral', 160)))
            
            if 'region' in accion:
                ttk.Label(campos_frame, text=f"Región actual: {accion['region']}").pack(anchor="w")
            ttk.Button(campos_frame, text="Cambiar Región", command=self.capturar_region_desde_dialogo).pack(pady=5)
        
        def guardar_cambios():
            nuevo_nombre = nombre_entry.get()
            if not nuevo_nombre:
                messagebox.showerror("Error", "Debe ingresar un nombre")
                return
            
            # Actualizar acción
            accion['nombre'] = nuevo_nombre
            
            if accion['tipo'] == "click" and hasattr(self, 'posicion_actual'):
                accion['posicion'] = {"x": self.posicion_actual.x, "y": self.posicion_actual.y}
            elif accion['tipo'] == "escribir":
                accion['texto'] = texto_entry.get()
                if hasattr(self, 'posicion_actual'):
                    accion['posicion'] = {"x": self.posicion_actual.x, "y": self.posicion_actual.y}
            elif accion['tipo'] == "tecla":
                accion['tecla'] = tecla_entry.get()
            elif accion['tipo'] == "esperar_texto":
                accion['texto'] = texto_entry.get()
                accion['timeout'] = int(timeout_entry.get()) if timeout_entry.get().isdigit() else 30
                accion['umbral'] = int(umbral_entry.get()) if umbral_entry.get().isdigit() else 160
                if hasattr(self, 'region_dialogo_actual'):
                    accion['region'] = self.region_dialogo_actual
            
            self.actualizar_lista_acciones()
            self.guardar_configuracion_automatica()
            dialogo.destroy()
            self.log(f"Acción '{nuevo_nombre}' actualizada")
        
        ttk.Button(dialogo, text="Guardar Cambios", command=guardar_cambios).grid(row=2, column=0, columnspan=2, pady=10)
        
        dialogo.transient(self.root)
        dialogo.grab_set()
        self.root.wait_window(dialogo)

    def eliminar_accion(self):
        seleccion = self.lista_acciones.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione una acción para eliminar")
            return
        
        item = seleccion[0]
        nombre = self.lista_acciones.item(item, 'text')
        
        if messagebox.askyesno("Confirmar", f"¿Eliminar la acción '{nombre}'?"):
            for i, accion in enumerate(self.acciones):
                if accion['nombre'] == nombre:
                    del self.acciones[i]
                    break
            
            self.actualizar_lista_acciones()
            self.guardar_configuracion_automatica()
            self.log(f"Acción '{nombre}' eliminada")

    # Funciones de automatización
    def preprocesar_imagen(self, imagen, umbral=160):
        """Convierte la imagen a blanco y negro puro"""
        # Convertir a escala de grises
        imagen = imagen.convert('L')
        # Aplicar binarización
        imagen = imagen.point(lambda x: 0 if x < umbral else 255, '1')
        return imagen

    def esperar_texto(self, texto_esperado, timeout=30, region=None, umbral=160):
        start_time = time.time()
        texto_esperado = texto_esperado.lower()
        
        while time.time() - start_time < timeout:
            if not self.ejecucion_automatica:
                return False
                
            try:
                # Capturar pantalla o región
                screenshot = pyautogui.screenshot(region=region) if region else pyautogui.screenshot()
                
                # Preprocesamiento intensivo (blanco y negro puro)
                screenshot = self.preprocesar_imagen(screenshot, umbral)
                
                # Configuración de Tesseract
                config_tesseract = '--psm 6 --oem 3'
                
                # Usar OCR para extraer texto
                texto_en_pantalla = pytesseract.image_to_string(screenshot, config=config_tesseract).lower()
                
                if texto_esperado in texto_en_pantalla:
                    self.log(f"Texto encontrado: '{texto_esperado}'")
                    return True
                    
            except Exception as e:
                self.log(f"Error en OCR: {str(e)}")
                time.sleep(1)
                continue
            
            tiempo_restante = int(timeout - (time.time() - start_time))
            self.estado_label.config(text=f"Esperando texto... ({tiempo_restante}s)")
            time.sleep(1)
        
        self.log(f"Tiempo agotado sin encontrar: '{texto_esperado}'")
        return False

    def ejecutar_acciones(self, acciones):
        """Ejecuta una lista de acciones"""
        for accion in acciones:
            if not self.ejecucion_automatica:
                return False
                
            self.log(f"Ejecutando: {accion['nombre']}")
            self.estado_label.config(text=f"Ejecutando: {accion['nombre']}")
            
            if accion['tipo'] == "terminar":
                self.log("Acción 'Terminar proceso' detectada - finalizando ejecución")
                return False
            elif accion['tipo'] == "click":
                pyautogui.click(accion['posicion']['x'], accion['posicion']['y'])
            elif accion['tipo'] == "escribir":
                pyautogui.click(accion['posicion']['x'], accion['posicion']['y'])
                pyautogui.write(accion['texto'])
            elif accion['tipo'] == "tecla":
                pyautogui.press(accion['tecla'])
            elif accion['tipo'] == "esperar_texto":
                region = tuple(accion['region']) if accion.get('region') else None
                umbral = accion.get('umbral', 160)
                if not self.esperar_texto(accion['texto'], accion['timeout'], region, umbral):
                    self.log(f"No se encontró el texto '{accion['texto']}'")
                    return False
            
            time.sleep(0.3)
        return True

    def ejecutar_automatizacion(self):
        if not self.acciones:
            messagebox.showwarning("Advertencia", "No hay acciones configuradas para ejecutar")
            return
    
        self.log("\nIniciando ejecución manual...")
        self.estado_label.config(text="Ejecutando...", foreground="blue")
    
        # Guardar estado de ejecución automática y forzar modo manual
        estado_previo = self.ejecucion_automatica
        self.ejecucion_automatica = True  # Permitir que se ejecuten todas las acciones
    
        try:
            for i, accion in enumerate(self.acciones, 1):
                if not self.ejecucion_automatica:
                    break
                
                self.log(f"\nProcesando acción {i}/{len(self.acciones)}: {accion['nombre']}")
                self.estado_label.config(text=f"Procesando: {accion['nombre']}")
            
                if accion['tipo'] == "condicional":
                    # Ejecutar lógica condicional
                    region = tuple(accion['region']) if accion.get('region') else None
                    self.log(f"Verificando condición: '{accion['texto_condicion']}'")
                
                    texto_encontrado = self.esperar_texto(
                        accion['texto_condicion'],
                        accion['timeout'],
                        region
                    )
                
                    if texto_encontrado:
                        self.log("Condición CUMPLIDA - ejecutando acciones correspondientes")
                        if not self.ejecutar_acciones(accion['acciones_si']):
                            break
                    else:
                        self.log("Condición NO cumplida - ejecutando acciones alternativas")
                        if not self.ejecutar_acciones(accion['acciones_no']):
                            break
                else:
                    # Ejecutar acción normal
                    if accion['tipo'] == "click":
                        pyautogui.click(accion['posicion']['x'], accion['posicion']['y'])
                    elif accion['tipo'] == "escribir":
                        pyautogui.click(accion['posicion']['x'], accion['posicion']['y'])
                        pyautogui.write(accion['texto'])
                    elif accion['tipo'] == "tecla":
                        pyautogui.press(accion['tecla'])
                    elif accion['tipo'] == "esperar_texto":
                        region = tuple(accion['region']) if accion.get('region') else None
                        umbral = accion.get('umbral', 160)
                        if not self.esperar_texto(accion['texto'], accion['timeout'], region, umbral):
                            self.log(f"No se encontró el texto '{accion['texto']}'")
                            break
                    elif accion['tipo'] == "terminar":
                        self.log("Acción 'Terminar proceso' detectada - finalizando ejecución")
                        break
            
                time.sleep(0.3)
        
            self.log("\nEjecución completada")
            self.estado_label.config(text="Listo", foreground="green")
        except pyautogui.FailSafeException:
            self.log("\nEjecución cancelada por failsafe (mouse en esquina)")
            self.estado_label.config(text="Detenido", foreground="red")
        except Exception as e:
            self.log(f"\nError durante la ejecución: {str(e)}")
            self.estado_label.config(text="Error", foreground="red")
        finally:
            # Restaurar estado original
            self.ejecucion_automatica = estado_previo

    def iniciar_automatico(self):
        if not self.acciones:
            messagebox.showwarning("Advertencia", "No hay acciones configuradas para ejecutar")
            return
            
        self.ejecucion_automatica = True
        self.intervalo = int(self.intervalo_entry.get()) if self.intervalo_entry.get().isdigit() else 30
        self.log(f"\nIniciando ejecución automática cada {self.intervalo} segundos...")
        self.estado_label.config(text="Ejecución automática", foreground="blue")
        threading.Thread(target=self.ciclo_automatico, daemon=True).start()

    def ciclo_automatico(self):
        while self.ejecucion_automatica:
            self.ejecutar_automatizacion()
            if not self.ejecucion_automatica:
                break
                
            for i in range(self.intervalo, 0, -1):
                if not self.ejecucion_automatica:
                    break
                self.estado_label.config(text=f"Esperando... {i}s")
                time.sleep(1)
        
        self.estado_label.config(text="Listo", foreground="green")

    def detener_ejecucion(self):
        self.ejecucion_automatica = False
        self.log("Ejecución detenida por el usuario")
        self.estado_label.config(text="Detenido", foreground="red")

if __name__ == "__main__":
    root = tk.Tk()
    try:
        app = AutomatizadorApp(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Error crítico", f"Ha ocurrido un error: {str(e)}")
        root.destroy()