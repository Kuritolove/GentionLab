import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime, timedelta
import pandas as pd
from matplotlib import pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import webbrowser
import os

class SistemaGestionLaboratorio:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestión de Laboratorio de Redes")
        self.root.geometry("1200x800")
        self.root.state('zoomed')
        
        # Configuración de estilos
        self.configurar_estilos()
        
        # Barra de menú
        self.crear_barra_menu()
        
        # Conexión a la base de datos
        self.conexion_db()
        
        # Crear pestañas principales
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)
        
        # Frames para cada pestaña
        self.frame_reportes = ttk.Frame(self.notebook)
        self.frame_inventario = ttk.Frame(self.notebook)
        self.frame_gestion = ttk.Frame(self.notebook)
        self.frame_usuarios = ttk.Frame(self.notebook)
        
        self.notebook.add(self.frame_reportes, text="Reportes de Equipos")
        self.notebook.add(self.frame_inventario, text="Gestión de Inventario")
        self.notebook.add(self.frame_gestion, text="Gestión del Laboratorio")
        self.notebook.add(self.frame_usuarios, text="Usuarios y Accesos")
        
        # Inicializar componentes de cada pestaña
        self.inicializar_reportes()
        self.inicializar_inventario()
        self.inicializar_gestion()
        self.inicializar_usuarios()
        
        # Cargar datos iniciales
        self.cargar_datos_iniciales()
        
        # Configurar cierre seguro
        self.root.protocol("WM_DELETE_WINDOW", self.cerrar_aplicacion)
    
    def configurar_estilos(self):
        """Configura los estilos visuales de la aplicación"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configurar colores
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        style.configure('TButton', font=('Arial', 10), padding=5)
        style.configure('TEntry', font=('Arial', 10), padding=5)
        style.configure('TCombobox', font=('Arial', 10), padding=5)
        style.configure('Treeview', font=('Arial', 10), rowheight=25)
        style.configure('Treeview.Heading', font=('Arial', 10, 'bold'))
        style.map('TButton', foreground=[('active', '!disabled', 'black')], 
                  background=[('active', '#4CAF50')])
    
    def crear_barra_menu(self):
        """Crea la barra de menú principal"""
        menubar = tk.Menu(self.root)
        
        # Menú Archivo
        menu_archivo = tk.Menu(menubar, tearoff=0)
        menu_archivo.add_command(label="Copia de seguridad", command=self.crear_respaldo)
        menu_archivo.add_command(label="Restaurar", command=self.restaurar_respaldo)
        menu_archivo.add_separator()
        menu_archivo.add_command(label="Salir", command=self.cerrar_aplicacion)
        menubar.add_cascade(label="Archivo", menu=menu_archivo)
        
        # Menú Ayuda
        menu_ayuda = tk.Menu(menubar, tearoff=0)
        menu_ayuda.add_command(label="Documentación", command=self.mostrar_documentacion)
        menu_ayuda.add_command(label="Acerca de...", command=self.mostrar_acerca_de)
        menubar.add_cascade(label="Ayuda", menu=menu_ayuda)
        
        self.root.config(menu=menubar)
    
    def conexion_db(self):
        """Establece conexión con la base de datos SQLite"""
        try:
            self.conn = sqlite3.connect('laboratorio.db')
            self.c = self.conn.cursor()
            
            # Crear tablas si no existen
            self.crear_tablas()
                    
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo conectar a la base de datos: {e}")
            self.root.destroy()
    
    def crear_tablas(self):
        """Crea las tablas necesarias en la base de datos"""
        tablas = [
            """CREATE TABLE IF NOT EXISTS equipos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                tipo TEXT,
                modelo TEXT,
                serial TEXT UNIQUE,
                estado TEXT,
                ubicacion TEXT,
                fecha_adquisicion TEXT,
                ultimo_mantenimiento TEXT,
                observaciones TEXT
            )""",
            
            """CREATE TABLE IF NOT EXISTS inventario (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                componente TEXT NOT NULL,
                tipo TEXT,
                cantidad INTEGER,
                minimo INTEGER,
                proveedor TEXT,
                ubicacion TEXT,
                fecha_actualizacion TEXT,
                observaciones TEXT
            )""",
            
            """CREATE TABLE IF NOT EXISTS reportes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                equipo_id INTEGER,
                tipo TEXT,
                descripcion TEXT,
                fecha TEXT,
                estado TEXT,
                solucion TEXT,
                usuario TEXT,
                prioridad TEXT,
                FOREIGN KEY (equipo_id) REFERENCES equipos (id)
            )""",
            
            """CREATE TABLE IF NOT EXISTS reservas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                equipo_id INTEGER,
                usuario_id INTEGER,
                fecha_inicio TEXT,
                fecha_fin TEXT,
                proposito TEXT,
                estado TEXT,
                fecha_solicitud TEXT,
                FOREIGN KEY (equipo_id) REFERENCES equipos (id),
                FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
            )""",
            
            """CREATE TABLE IF NOT EXISTS mantenimientos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                equipo_id INTEGER,
                tipo TEXT,
                fecha_programada TEXT,
                fecha_realizado TEXT,
                descripcion TEXT,
                tecnico TEXT,
                estado TEXT,
                costo REAL,
                observaciones TEXT,
                FOREIGN KEY (equipo_id) REFERENCES equipos (id)
            )""",
            
            """CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                apellido TEXT,
                email TEXT UNIQUE,
                rol TEXT,
                usuario TEXT UNIQUE,
                contrasena TEXT,
                fecha_registro TEXT,
                estado TEXT
            )""",
            
            """CREATE TABLE IF NOT EXISTS accesos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario_id INTEGER,
                fecha_hora TEXT,
                accion TEXT,
                detalles TEXT,
                FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
            )"""
        ]
        
        for tabla in tablas:
            try:
                self.c.execute(tabla)
            except sqlite3.Error as e:
                messagebox.showerror("Error", f"No se pudo crear tabla: {e}")
        
        self.conn.commit()
    
    def cargar_datos_iniciales(self):
        """Carga datos iniciales en las tablas"""
        self.buscar_reportes()
        self.actualizar_inventario()
        self.actualizar_equipos()
        self.actualizar_reservas()
        self.actualizar_mantenimientos()
        self.actualizar_usuarios()
    
    # ------------------------- Pestaña de Reportes -------------------------
    def inicializar_reportes(self):
        """Configura los componentes de la pestaña de reportes"""
        # Frame superior para filtros
        frame_filtros = ttk.LabelFrame(self.frame_reportes, text="Filtros de Búsqueda", padding=10)
        frame_filtros.pack(fill='x', padx=10, pady=5)
        
        # Componentes de filtro
        ttk.Label(frame_filtros, text="Tipo de Reporte:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.combo_tipo_reporte = ttk.Combobox(frame_filtros, values=["Todos", "Hardware", "Software", "Redes", "Otros"], state='readonly')
        self.combo_tipo_reporte.grid(row=0, column=1, padx=5, pady=5, sticky='we')
        self.combo_tipo_reporte.set("Todos")
        
        ttk.Label(frame_filtros, text="Estado:").grid(row=0, column=2, padx=5, pady=5, sticky='e')
        self.combo_estado_reporte = ttk.Combobox(frame_filtros, values=["Todos", "Abierto", "En Progreso", "Resuelto"], state='readonly')
        self.combo_estado_reporte.grid(row=0, column=3, padx=5, pady=5, sticky='we')
        self.combo_estado_reporte.set("Todos")
        
        ttk.Label(frame_filtros, text="Prioridad:").grid(row=0, column=4, padx=5, pady=5, sticky='e')
        self.combo_prioridad_reporte = ttk.Combobox(frame_filtros, values=["Todos", "Alta", "Media", "Baja"], state='readonly')
        self.combo_prioridad_reporte.grid(row=0, column=5, padx=5, pady=5, sticky='we')
        self.combo_prioridad_reporte.set("Todos")
        
        ttk.Label(frame_filtros, text="Fecha Desde:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.entry_fecha_desde = ttk.Entry(frame_filtros)
        self.entry_fecha_desde.grid(row=1, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_filtros, text="Fecha Hasta:").grid(row=1, column=2, padx=5, pady=5, sticky='e')
        self.entry_fecha_hasta = ttk.Entry(frame_filtros)
        self.entry_fecha_hasta.grid(row=1, column=3, padx=5, pady=5, sticky='we')
        
        btn_buscar = ttk.Button(frame_filtros, text="Buscar", command=self.buscar_reportes)
        btn_buscar.grid(row=1, column=4, padx=5, pady=5)
        
        btn_limpiar = ttk.Button(frame_filtros, text="Limpiar", command=self.limpiar_filtros_reportes)
        btn_limpiar.grid(row=1, column=5, padx=5, pady=5)
        
        # Frame para la tabla de reportes
        frame_tabla = ttk.Frame(self.frame_reportes)
        frame_tabla.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Tabla de reportes
        columns = ("ID", "Equipo", "Tipo", "Descripción", "Fecha", "Estado", "Prioridad")
        self.tree_reportes = ttk.Treeview(frame_tabla, columns=columns, show='headings', selectmode='browse')
        
        for col in columns:
            self.tree_reportes.heading(col, text=col)
            self.tree_reportes.column(col, width=100, anchor='center')
        
        self.tree_reportes.column("Descripción", width=250)
        self.tree_reportes.pack(fill='both', expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree_reportes.yview)
        scrollbar.pack(side='right', fill='y')
        self.tree_reportes.configure(yscrollcommand=scrollbar.set)
        
        # Frame inferior para acciones
        frame_acciones = ttk.Frame(self.frame_reportes)
        frame_acciones.pack(fill='x', padx=10, pady=5)
        
        btn_nuevo_reporte = ttk.Button(frame_acciones, text="Nuevo Reporte", command=self.abrir_nuevo_reporte)
        btn_nuevo_reporte.pack(side='left', padx=5)
        
        btn_ver_detalle = ttk.Button(frame_acciones, text="Ver Detalle", command=self.ver_detalle_reporte)
        btn_ver_detalle.pack(side='left', padx=5)
        
        btn_exportar = ttk.Button(frame_acciones, text="Exportar a Excel", command=self.exportar_reportes)
        btn_exportar.pack(side='left', padx=5)
        
        btn_imprimir = ttk.Button(frame_acciones, text="Imprimir", command=self.imprimir_reporte)
        btn_imprimir.pack(side='left', padx=5)
    
    def limpiar_filtros_reportes(self):
        """Limpia los filtros de búsqueda de reportes"""
        self.combo_tipo_reporte.set("Todos")
        self.combo_estado_reporte.set("Todos")
        self.combo_prioridad_reporte.set("Todos")
        self.entry_fecha_desde.delete(0, 'end')
        self.entry_fecha_hasta.delete(0, 'end')
        self.buscar_reportes()
    
    def buscar_reportes(self):
        """Busca reportes según los filtros aplicados"""
        try:
            # Limpiar tabla
            for item in self.tree_reportes.get_children():
                self.tree_reportes.delete(item)
            
            # Construir consulta SQL
            query = """SELECT r.id, e.nombre, r.tipo, r.descripcion, r.fecha, r.estado, r.prioridad 
                       FROM reportes r LEFT JOIN equipos e ON r.equipo_id = e.id 
                       WHERE 1=1"""
            params = []
            
            tipo = self.combo_tipo_reporte.get()
            if tipo != "Todos":
                query += " AND r.tipo = ?"
                params.append(tipo)
                
            estado = self.combo_estado_reporte.get()
            if estado != "Todos":
                query += " AND r.estado = ?"
                params.append(estado)
                
            prioridad = self.combo_prioridad_reporte.get()
            if prioridad != "Todos":
                query += " AND r.prioridad = ?"
                params.append(prioridad)
                
            fecha_desde = self.entry_fecha_desde.get()
            if fecha_desde:
                query += " AND r.fecha >= ?"
                params.append(fecha_desde)
                
            fecha_hasta = self.entry_fecha_hasta.get()
            if fecha_hasta:
                query += " AND r.fecha <= ?"
                params.append(fecha_hasta)
            
            query += " ORDER BY r.fecha DESC"
            
            # Ejecutar consulta
            self.c.execute(query, params)
            reportes = self.c.fetchall()
            
            # Llenar tabla
            for reporte in reportes:
                self.tree_reportes.insert('', 'end', values=reporte)
                
            # Resaltar reportes según prioridad
            for item in self.tree_reportes.get_children():
                valores = self.tree_reportes.item(item, 'values')
                if valores[6] == "Alta":
                    self.tree_reportes.tag_configure('alta', foreground='red')
                    self.tree_reportes.item(item, tags=('alta',))
                elif valores[6] == "Media":
                    self.tree_reportes.tag_configure('media', foreground='orange')
                    self.tree_reportes.item(item, tags=('media',))
                
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"Error al buscar reportes: {e}")
    
    def abrir_nuevo_reporte(self):
        """Abre ventana para crear nuevo reporte"""
        ventana = tk.Toplevel(self.root)
        ventana.title("Nuevo Reporte")
        ventana.geometry("500x400")
        ventana.grab_set()
        
        # Variables
        self.var_equipo = tk.StringVar()
        self.var_tipo = tk.StringVar(value="Hardware")
        self.var_prioridad = tk.StringVar(value="Media")
        self.var_descripcion = tk.StringVar()
        
        # Campos del formulario
        frame_form = ttk.Frame(ventana, padding=10)
        frame_form.pack(fill='both', expand=True)
        
        # Equipo
        ttk.Label(frame_form, text="Equipo:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        combo_equipo = ttk.Combobox(frame_form, textvariable=self.var_equipo, state='readonly')
        combo_equipo.grid(row=0, column=1, padx=5, pady=5, sticky='we')
        
        # Cargar equipos disponibles
        self.c.execute("SELECT id, nombre FROM equipos ORDER BY nombre")
        equipos = self.c.fetchall()
        combo_equipo['values'] = [f"{e[1]} (ID: {e[0]})" for e in equipos]
        
        # Tipo de reporte
        ttk.Label(frame_form, text="Tipo de Reporte:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        combo_tipo = ttk.Combobox(frame_form, textvariable=self.var_tipo, 
                                values=["Hardware", "Software", "Redes", "Otros"], state='readonly')
        combo_tipo.grid(row=1, column=1, padx=5, pady=5, sticky='we')
        
        # Prioridad
        ttk.Label(frame_form, text="Prioridad:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        combo_prioridad = ttk.Combobox(frame_form, textvariable=self.var_prioridad, 
                                     values=["Alta", "Media", "Baja"], state='readonly')
        combo_prioridad.grid(row=2, column=1, padx=5, pady=5, sticky='we')
        
        # Descripción
        ttk.Label(frame_form, text="Descripción:").grid(row=3, column=0, padx=5, pady=5, sticky='ne')
        text_desc = tk.Text(frame_form, height=10, width=40, wrap='word')
        text_desc.grid(row=3, column=1, padx=5, pady=5, sticky='we')
        
        # Barra de desplazamiento para el texto
        scrollbar = ttk.Scrollbar(frame_form, orient='vertical', command=text_desc.yview)
        scrollbar.grid(row=3, column=2, sticky='ns')
        text_desc.config(yscrollcommand=scrollbar.set)
        
        # Botones
        frame_botones = ttk.Frame(ventana)
        frame_botones.pack(fill='x', padx=10, pady=5)
        
        btn_guardar = ttk.Button(frame_botones, text="Guardar", command=lambda: self.guardar_reporte(
            self.var_equipo.get().split("(ID: ")[1].rstrip(")") if self.var_equipo.get() else None,
            self.var_tipo.get(),
            text_desc.get("1.0", "end").strip(),
            self.var_prioridad.get(),
            ventana
        ))
        btn_guardar.pack(side='left', padx=5)
        
        btn_cancelar = ttk.Button(frame_botones, text="Cancelar", command=ventana.destroy)
        btn_cancelar.pack(side='left', padx=5)
    
    def guardar_reporte(self, equipo_id, tipo, descripcion, prioridad, ventana):
        """Guarda un nuevo reporte en la base de datos"""
        if not descripcion:
            messagebox.showwarning("Advertencia", "La descripción es obligatoria")
            return
            
        try:
            fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            usuario = "admin"  # En una aplicación real, obtendríamos el usuario actual
            
            self.c.execute("""INSERT INTO reportes 
                              (equipo_id, tipo, descripcion, fecha, estado, usuario, prioridad) 
                              VALUES (?, ?, ?, ?, ?, ?, ?)""",
                          (equipo_id, tipo, descripcion, fecha_actual, "Abierto", usuario, prioridad))
            self.conn.commit()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Creó reporte de {tipo} (Prioridad: {prioridad})")
            
            messagebox.showinfo("Éxito", "Reporte creado correctamente")
            ventana.destroy()
            self.buscar_reportes()
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo guardar el reporte: {e}")
    
    def ver_detalle_reporte(self):
        """Muestra los detalles del reporte seleccionado"""
        seleccion = self.tree_reportes.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione un reporte para ver detalles")
            return
            
        reporte_id = self.tree_reportes.item(seleccion[0], 'values')[0]
        
        try:
            self.c.execute("""SELECT r.id, e.nombre, r.tipo, r.descripcion, r.fecha, r.estado, 
                              r.solucion, r.usuario, r.prioridad 
                              FROM reportes r LEFT JOIN equipos e ON r.equipo_id = e.id 
                              WHERE r.id = ?""", (reporte_id,))
            reporte = self.c.fetchone()
            
            ventana = tk.Toplevel(self.root)
            ventana.title(f"Detalle del Reporte #{reporte[0]}")
            ventana.geometry("600x500")
            
            # Frame principal
            frame_principal = ttk.Frame(ventana, padding=10)
            frame_principal.pack(fill='both', expand=True)
            
            # Mostrar información del reporte
            ttk.Label(frame_principal, text=f"ID: {reporte[0]}", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', pady=5)
            
            ttk.Label(frame_principal, text="Equipo:").grid(row=1, column=0, sticky='w', pady=2)
            ttk.Label(frame_principal, text=reporte[1]).grid(row=1, column=1, sticky='w', pady=2)
            
            ttk.Label(frame_principal, text="Tipo:").grid(row=2, column=0, sticky='w', pady=2)
            ttk.Label(frame_principal, text=reporte[2]).grid(row=2, column=1, sticky='w', pady=2)
            
            ttk.Label(frame_principal, text="Prioridad:").grid(row=3, column=0, sticky='w', pady=2)
            ttk.Label(frame_principal, text=reporte[8], 
                     foreground='red' if reporte[8] == "Alta" else 'orange' if reporte[8] == "Media" else 'black').grid(row=3, column=1, sticky='w', pady=2)
            
            ttk.Label(frame_principal, text="Fecha:").grid(row=4, column=0, sticky='w', pady=2)
            ttk.Label(frame_principal, text=reporte[4]).grid(row=4, column=1, sticky='w', pady=2)
            
            ttk.Label(frame_principal, text="Estado:").grid(row=5, column=0, sticky='w', pady=2)
            ttk.Label(frame_principal, text=reporte[5], 
                     foreground='green' if reporte[5] == "Resuelto" else 'blue' if reporte[5] == "En Progreso" else 'black').grid(row=5, column=1, sticky='w', pady=2)
            
            ttk.Label(frame_principal, text="Reportado por:").grid(row=6, column=0, sticky='w', pady=2)
            ttk.Label(frame_principal, text=reporte[7]).grid(row=6, column=1, sticky='w', pady=2)
            
            # Descripción
            ttk.Label(frame_principal, text="Descripción:", font=('Arial', 10, 'bold')).grid(row=7, column=0, sticky='nw', pady=5)
            frame_desc = ttk.Frame(frame_principal, borderwidth=1, relief='solid')
            frame_desc.grid(row=8, column=0, columnspan=2, sticky='we', padx=5, pady=2)
            
            text_desc = tk.Text(frame_desc, height=5, width=60, wrap='word', padx=5, pady=5)
            text_desc.insert('1.0', reporte[3])
            text_desc.config(state='disabled')
            text_desc.pack(fill='both', expand=True)
            
            # Solución (si existe)
            if reporte[6]:
                ttk.Label(frame_principal, text="Solución:", font=('Arial', 10, 'bold')).grid(row=9, column=0, sticky='nw', pady=5)
                frame_sol = ttk.Frame(frame_principal, borderwidth=1, relief='solid')
                frame_sol.grid(row=10, column=0, columnspan=2, sticky='we', padx=5, pady=2)
                
                text_sol = tk.Text(frame_sol, height=5, width=60, wrap='word', padx=5, pady=5)
                text_sol.insert('1.0', reporte[6])
                text_sol.config(state='disabled')
                text_sol.pack(fill='both', expand=True)
            else:
                ttk.Label(frame_principal, text="Solución: No registrada", font=('Arial', 10)).grid(row=9, column=0, sticky='nw', pady=5)
            
            # Botones para cambiar estado si está abierto
            if reporte[5] == "Abierto":
                frame_botones = ttk.Frame(ventana)
                frame_botones.pack(fill='x', padx=10, pady=10)
                
                btn_en_progreso = ttk.Button(frame_botones, text="Marcar en Progreso", 
                                            command=lambda: self.cambiar_estado_reporte(reporte_id, "En Progreso", ventana))
                btn_en_progreso.pack(side='left', padx=5)
                
                btn_resolver = ttk.Button(frame_botones, text="Resolver", 
                                        command=lambda: self.resolver_reporte(reporte_id, ventana))
                btn_resolver.pack(side='left', padx=5)
            
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo cargar el detalle: {e}")
    
    def cambiar_estado_reporte(self, reporte_id, estado, ventana):
        """Cambia el estado de un reporte"""
        try:
            self.c.execute("UPDATE reportes SET estado = ? WHERE id = ?", (estado, reporte_id))
            self.conn.commit()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Cambió estado de reporte {reporte_id} a {estado}")
            
            messagebox.showinfo("Éxito", f"Reporte marcado como {estado}")
            ventana.destroy()
            self.buscar_reportes()
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo cambiar el estado: {e}")
    
    def resolver_reporte(self, reporte_id, ventana):
        """Abre ventana para registrar solución a un reporte"""
        ventana_sol = tk.Toplevel(self.root)
        ventana_sol.title("Registrar Solución")
        ventana_sol.geometry("500x300")
        ventana_sol.grab_set()
        
        # Frame principal
        frame_principal = ttk.Frame(ventana_sol, padding=10)
        frame_principal.pack(fill='both', expand=True)
        
        ttk.Label(frame_principal, text="Solución aplicada:").pack(anchor='w', pady=5)
        text_sol = tk.Text(frame_principal, height=10, width=50, wrap='word')
        text_sol.pack(fill='both', expand=True, pady=5)
        
        # Barra de desplazamiento
        scrollbar = ttk.Scrollbar(frame_principal, orient='vertical', command=text_sol.yview)
        scrollbar.pack(side='right', fill='y')
        text_sol.config(yscrollcommand=scrollbar.set)
        
        # Botones
        frame_botones = ttk.Frame(ventana_sol)
        frame_botones.pack(fill='x', padx=10, pady=10)
        
        btn_guardar = ttk.Button(frame_botones, text="Guardar", 
                                command=lambda: self.guardar_solucion(reporte_id, text_sol.get("1.0", "end").strip(), 
                                                                    ventana_sol, ventana))
        btn_guardar.pack(side='left', padx=5)
        
        btn_cancelar = ttk.Button(frame_botones, text="Cancelar", command=ventana_sol.destroy)
        btn_cancelar.pack(side='left', padx=5)
    
    def guardar_solucion(self, reporte_id, solucion, ventana_sol, ventana_detalle):
        """Guarda la solución y marca el reporte como resuelto"""
        if not solucion:
            messagebox.showwarning("Advertencia", "La solución es obligatoria")
            return
            
        try:
            self.c.execute("UPDATE reportes SET solucion = ?, estado = 'Resuelto' WHERE id = ?", 
                          (solucion, reporte_id))
            self.conn.commit()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Resolvió reporte {reporte_id}")
            
            messagebox.showinfo("Éxito", "Solución registrada y reporte marcado como resuelto")
            ventana_sol.destroy()
            ventana_detalle.destroy()
            self.buscar_reportes()
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo guardar la solución: {e}")
    
    def exportar_reportes(self):
        """Exporta los reportes a un archivo Excel"""
        try:
            # Obtener todos los reportes visibles
            items = self.tree_reportes.get_children()
            data = [self.tree_reportes.item(item, 'values') for item in items]
            
            if not data:
                messagebox.showwarning("Advertencia", "No hay datos para exportar")
                return
                
            # Crear DataFrame
            df = pd.DataFrame(data, columns=["ID", "Equipo", "Tipo", "Descripción", "Fecha", "Estado", "Prioridad"])
            
            # Preguntar dónde guardar
            filepath = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Guardar reportes como"
            )
            
            if not filepath:
                return
                
            # Guardar en Excel
            df.to_excel(filepath, index=False)
            
            messagebox.showinfo("Éxito", f"Reportes exportados a:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar: {e}")
    
    def imprimir_reporte(self):
        """Prepara la impresión de un reporte"""
        seleccion = self.tree_reportes.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione un reporte para imprimir")
            return
            
        reporte_id = self.tree_reportes.item(seleccion[0], 'values')[0]
        
        try:
            self.c.execute("""SELECT r.id, e.nombre, r.tipo, r.descripcion, r.fecha, r.estado, 
                              r.solucion, r.usuario, r.prioridad 
                              FROM reportes r LEFT JOIN equipos e ON r.equipo_id = e.id 
                              WHERE r.id = ?""", (reporte_id,))
            reporte = self.c.fetchone()
            
            # Crear contenido HTML para imprimir
            html = f"""
            <html>
            <head>
                <title>Reporte #{reporte[0]}</title>
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; }}
                    h1 {{ color: #333; border-bottom: 1px solid #333; }}
                    .info {{ margin-bottom: 15px; }}
                    .label {{ font-weight: bold; }}
                    .desc {{ border: 1px solid #ddd; padding: 10px; margin: 10px 0; }}
                    .sol {{ border: 1px solid #ddd; padding: 10px; margin: 10px 0; }}
                </style>
            </head>
            <body>
                <h1>Reporte #{reporte[0]}</h1>
                
                <div class="info">
                    <span class="label">Equipo:</span> {reporte[1]}<br>
                    <span class="label">Tipo:</span> {reporte[2]}<br>
                    <span class="label">Prioridad:</span> {reporte[8]}<br>
                    <span class="label">Fecha:</span> {reporte[4]}<br>
                    <span class="label">Estado:</span> {reporte[5]}<br>
                    <span class="label">Reportado por:</span> {reporte[7]}<br>
                </div>
                
                <div class="info">
                    <span class="label">Descripción:</span>
                    <div class="desc">{reporte[3]}</div>
                </div>
            """
            
            if reporte[6]:
                html += f"""
                <div class="info">
                    <span class="label">Solución:</span>
                    <div class="sol">{reporte[6]}</div>
                </div>
                """
            
            html += """
            </body>
            </html>
            """
            
            # Guardar temporalmente el HTML
            temp_file = "temp_reporte.html"
            with open(temp_file, "w", encoding="utf-8") as f:
                f.write(html)
            
            # Abrir en navegador para imprimir
            webbrowser.open(temp_file)
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo preparar el reporte para impresión: {e}")
    
    # ------------------------- Pestaña de Inventario -------------------------
    def inicializar_inventario(self):
        """Configura los componentes de la pestaña de inventario"""
        # Frame superior para controles
        frame_controles = ttk.Frame(self.frame_inventario)
        frame_controles.pack(fill='x', padx=10, pady=5)
        
        # Botones de acción
        btn_agregar = ttk.Button(frame_controles, text="Agregar Componente", command=self.agregar_componente)
        btn_agregar.pack(side='left', padx=5)
        
        btn_editar = ttk.Button(frame_controles, text="Editar", command=self.editar_componente)
        btn_editar.pack(side='left', padx=5)
        
        btn_eliminar = ttk.Button(frame_controles, text="Eliminar", command=self.eliminar_componente)
        btn_eliminar.pack(side='left', padx=5)
        
        btn_actualizar = ttk.Button(frame_controles, text="Actualizar", command=self.actualizar_inventario)
        btn_actualizar.pack(side='left', padx=5)
        
        btn_exportar = ttk.Button(frame_controles, text="Exportar", command=self.exportar_inventario)
        btn_exportar.pack(side='left', padx=5)
        
        # Frame para filtros
        frame_filtros = ttk.LabelFrame(self.frame_inventario, text="Filtros", padding=5)
        frame_filtros.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(frame_filtros, text="Tipo:").grid(row=0, column=0, padx=5, pady=2, sticky='e')
        self.combo_tipo_inventario = ttk.Combobox(frame_filtros, values=["Todos", "Hardware", "Software", "Redes", "Consumibles"], state='readonly')
        self.combo_tipo_inventario.grid(row=0, column=1, padx=5, pady=2, sticky='we')
        self.combo_tipo_inventario.set("Todos")
        self.combo_tipo_inventario.bind("<<ComboboxSelected>>", lambda e: self.actualizar_inventario())
        
        ttk.Label(frame_filtros, text="Ubicación:").grid(row=0, column=2, padx=5, pady=2, sticky='e')
        self.combo_ubicacion_inventario = ttk.Combobox(frame_filtros, state='readonly')
        self.combo_ubicacion_inventario.grid(row=0, column=3, padx=5, pady=2, sticky='we')
        self.combo_ubicacion_inventario.set("Todos")
        self.combo_ubicacion_inventario.bind("<<ComboboxSelected>>", lambda e: self.actualizar_inventario())
        
        # Cargar ubicaciones disponibles
        self.cargar_ubicaciones_inventario()
        
        # Frame para la tabla de inventario
        frame_tabla = ttk.Frame(self.frame_inventario)
        frame_tabla.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Tabla de inventario
        columns = ("ID", "Componente", "Tipo", "Cantidad", "Mínimo", "Proveedor", "Ubicación", "Últ. Actualización")
        self.tree_inventario = ttk.Treeview(frame_tabla, columns=columns, show='headings', selectmode='browse')
        
        for col in columns:
            self.tree_inventario.heading(col, text=col)
            self.tree_inventario.column(col, width=100, anchor='center')
        
        self.tree_inventario.column("Componente", width=150)
        self.tree_inventario.column("Proveedor", width=150)
        self.tree_inventario.column("Ubicación", width=120)
        self.tree_inventario.pack(fill='both', expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree_inventario.yview)
        scrollbar.pack(side='right', fill='y')
        self.tree_inventario.configure(yscrollcommand=scrollbar.set)
        
        # Frame para gráficos
        frame_graficos = ttk.LabelFrame(self.frame_inventario, text="Estadísticas de Inventario", padding=10)
        frame_graficos.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Crear gráficos
        self.crear_graficos_inventario(frame_graficos)
    
    def cargar_ubicaciones_inventario(self):
        """Carga las ubicaciones disponibles para filtrar"""
        try:
            self.c.execute("SELECT DISTINCT ubicacion FROM inventario WHERE ubicacion IS NOT NULL AND ubicacion != '' ORDER BY ubicacion")
            ubicaciones = ["Todos"] + [u[0] for u in self.c.fetchall()]
            self.combo_ubicacion_inventario['values'] = ubicaciones
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudieron cargar las ubicaciones: {e}")
    
    def actualizar_inventario(self):
        """Actualiza la tabla de inventario con los datos de la base de datos"""
        try:
            # Limpiar tabla
            for item in self.tree_inventario.get_children():
                self.tree_inventario.delete(item)
            
            # Construir consulta con filtros
            query = "SELECT id, componente, tipo, cantidad, minimo, proveedor, ubicacion, fecha_actualizacion FROM inventario WHERE 1=1"
            params = []
            
            tipo = self.combo_tipo_inventario.get()
            if tipo != "Todos":
                query += " AND tipo = ?"
                params.append(tipo)
                
            ubicacion = self.combo_ubicacion_inventario.get()
            if ubicacion != "Todos":
                query += " AND ubicacion = ?"
                params.append(ubicacion)
            
            query += " ORDER BY componente"
            
            # Obtener datos
            self.c.execute(query, params)
            componentes = self.c.fetchall()
            
            # Llenar tabla
            for comp in componentes:
                self.tree_inventario.insert('', 'end', values=comp)
                
            # Resaltar componentes bajo mínimo
            for item in self.tree_inventario.get_children():
                valores = self.tree_inventario.item(item, 'values')
                if int(valores[3]) < int(valores[4]):
                    self.tree_inventario.tag_configure('bajo', background='#ffcccc')
                    self.tree_inventario.item(item, tags=('bajo',))
                    
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo cargar el inventario: {e}")
    
    def agregar_componente(self):
        """Abre ventana para agregar nuevo componente al inventario"""
        ventana = tk.Toplevel(self.root)
        ventana.title("Agregar Componente")
        ventana.geometry("500x400")
        ventana.grab_set()
        
        # Variables
        self.var_componente = tk.StringVar()
        self.var_tipo = tk.StringVar(value="Hardware")
        self.var_cantidad = tk.IntVar(value=1)
        self.var_minimo = tk.IntVar(value=1)
        self.var_proveedor = tk.StringVar()
        self.var_ubicacion = tk.StringVar()
        self.var_observaciones = tk.StringVar()
        
        # Frame principal
        frame_principal = ttk.Frame(ventana, padding=10)
        frame_principal.pack(fill='both', expand=True)
        
        # Campos del formulario
        ttk.Label(frame_principal, text="Componente:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_componente).grid(row=0, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Tipo:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        combo_tipo = ttk.Combobox(frame_principal, textvariable=self.var_tipo, 
                                 values=["Hardware", "Software", "Redes", "Consumibles", "Otros"], state='readonly')
        combo_tipo.grid(row=1, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Cantidad:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        ttk.Spinbox(frame_principal, textvariable=self.var_cantidad, from_=0, to=1000).grid(row=2, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Mínimo:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        ttk.Spinbox(frame_principal, textvariable=self.var_minimo, from_=0, to=1000).grid(row=3, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Proveedor:").grid(row=4, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_proveedor).grid(row=4, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Ubicación:").grid(row=5, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_ubicacion).grid(row=5, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Observaciones:").grid(row=6, column=0, padx=5, pady=5, sticky='ne')
        text_obs = tk.Text(frame_principal, height=4, width=30, wrap='word')
        text_obs.grid(row=6, column=1, padx=5, pady=5, sticky='we')
        
        # Botones
        frame_botones = ttk.Frame(ventana)
        frame_botones.pack(fill='x', padx=10, pady=10)
        
        btn_guardar = ttk.Button(frame_botones, text="Guardar", command=lambda: self.guardar_componente(
            self.var_componente.get(),
            self.var_tipo.get(),
            self.var_cantidad.get(),
            self.var_minimo.get(),
            self.var_proveedor.get(),
            self.var_ubicacion.get(),
            text_obs.get("1.0", "end").strip(),
            ventana
        ))
        btn_guardar.pack(side='left', padx=5)
        
        btn_cancelar = ttk.Button(frame_botones, text="Cancelar", command=ventana.destroy)
        btn_cancelar.pack(side='left', padx=5)
    
    def guardar_componente(self, componente, tipo, cantidad, minimo, proveedor, ubicacion, observaciones, ventana):
        """Guarda un nuevo componente en el inventario"""
        if not componente:
            messagebox.showwarning("Advertencia", "El nombre del componente es obligatorio")
            return
            
        try:
            fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            self.c.execute("""INSERT INTO inventario 
                              (componente, tipo, cantidad, minimo, proveedor, ubicacion, 
                               fecha_actualizacion, observaciones) 
                              VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                          (componente, tipo, cantidad, minimo, proveedor, ubicacion, fecha_actual, observaciones))
            self.conn.commit()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Agregó componente al inventario: {componente}")
            
            messagebox.showinfo("Éxito", "Componente agregado al inventario")
            ventana.destroy()
            self.cargar_ubicaciones_inventario()
            self.actualizar_inventario()
        except ValueError:
            messagebox.showerror("Error", "Cantidad y mínimo deben ser números enteros")
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo guardar el componente: {e}")
    
    def editar_componente(self):
        """Abre ventana para editar componente seleccionado"""
        seleccion = self.tree_inventario.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione un componente para editar")
            return
            
        componente_id = self.tree_inventario.item(seleccion[0], 'values')[0]
        
        try:
            self.c.execute("""SELECT componente, tipo, cantidad, minimo, proveedor, ubicacion, observaciones 
                              FROM inventario WHERE id = ?""", (componente_id,))
            comp = self.c.fetchone()
            
            ventana = tk.Toplevel(self.root)
            ventana.title("Editar Componente")
            ventana.geometry("500x400")
            ventana.grab_set()
            
            # Variables
            self.var_componente = tk.StringVar(value=comp[0])
            self.var_tipo = tk.StringVar(value=comp[1])
            self.var_cantidad = tk.IntVar(value=comp[2])
            self.var_minimo = tk.IntVar(value=comp[3])
            self.var_proveedor = tk.StringVar(value=comp[4] if comp[4] else "")
            self.var_ubicacion = tk.StringVar(value=comp[5] if comp[5] else "")
            self.var_observaciones = tk.StringVar(value=comp[6] if comp[6] else "")
            
            # Frame principal
            frame_principal = ttk.Frame(ventana, padding=10)
            frame_principal.pack(fill='both', expand=True)
            
            # Campos del formulario
            ttk.Label(frame_principal, text="Componente:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
            ttk.Entry(frame_principal, textvariable=self.var_componente).grid(row=0, column=1, padx=5, pady=5, sticky='we')
            
            ttk.Label(frame_principal, text="Tipo:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
            combo_tipo = ttk.Combobox(frame_principal, textvariable=self.var_tipo, 
                                     values=["Hardware", "Software", "Redes", "Consumibles", "Otros"], state='readonly')
            combo_tipo.grid(row=1, column=1, padx=5, pady=5, sticky='we')
            
            ttk.Label(frame_principal, text="Cantidad:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
            ttk.Spinbox(frame_principal, textvariable=self.var_cantidad, from_=0, to=1000).grid(row=2, column=1, padx=5, pady=5, sticky='we')
            
            ttk.Label(frame_principal, text="Mínimo:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
            ttk.Spinbox(frame_principal, textvariable=self.var_minimo, from_=0, to=1000).grid(row=3, column=1, padx=5, pady=5, sticky='we')
            
            ttk.Label(frame_principal, text="Proveedor:").grid(row=4, column=0, padx=5, pady=5, sticky='e')
            ttk.Entry(frame_principal, textvariable=self.var_proveedor).grid(row=4, column=1, padx=5, pady=5, sticky='we')
            
            ttk.Label(frame_principal, text="Ubicación:").grid(row=5, column=0, padx=5, pady=5, sticky='e')
            ttk.Entry(frame_principal, textvariable=self.var_ubicacion).grid(row=5, column=1, padx=5, pady=5, sticky='we')
            
            ttk.Label(frame_principal, text="Observaciones:").grid(row=6, column=0, padx=5, pady=5, sticky='ne')
            text_obs = tk.Text(frame_principal, height=4, width=30, wrap='word')
            text_obs.insert('1.0', comp[6] if comp[6] else "")
            text_obs.grid(row=6, column=1, padx=5, pady=5, sticky='we')
            
            # Botones
            frame_botones = ttk.Frame(ventana)
            frame_botones.pack(fill='x', padx=10, pady=10)
            
            btn_guardar = ttk.Button(frame_botones, text="Guardar", command=lambda: self.actualizar_componente(
                componente_id,
                self.var_componente.get(),
                self.var_tipo.get(),
                self.var_cantidad.get(),
                self.var_minimo.get(),
                self.var_proveedor.get(),
                self.var_ubicacion.get(),
                text_obs.get("1.0", "end").strip(),
                ventana
            ))
            btn_guardar.pack(side='left', padx=5)
            
            btn_cancelar = ttk.Button(frame_botones, text="Cancelar", command=ventana.destroy)
            btn_cancelar.pack(side='left', padx=5)
            
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo cargar el componente: {e}")
    
    def actualizar_componente(self, componente_id, componente, tipo, cantidad, minimo, proveedor, ubicacion, observaciones, ventana):
        """Actualiza un componente en la base de datos"""
        if not componente:
            messagebox.showwarning("Advertencia", "El nombre del componente es obligatorio")
            return
            
        try:
            fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            self.c.execute("""UPDATE inventario SET 
                              componente = ?, tipo = ?, cantidad = ?, minimo = ?, 
                              proveedor = ?, ubicacion = ?, fecha_actualizacion = ?, observaciones = ? 
                              WHERE id = ?""",
                          (componente, tipo, cantidad, minimo, proveedor, ubicacion, fecha_actual, observaciones, componente_id))
            self.conn.commit()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Actualizó componente en inventario: {componente}")
            
            messagebox.showinfo("Éxito", "Componente actualizado")
            ventana.destroy()
            self.cargar_ubicaciones_inventario()
            self.actualizar_inventario()
        except ValueError:
            messagebox.showerror("Error", "Cantidad y mínimo deben ser números enteros")
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo actualizar el componente: {e}")
    
    def eliminar_componente(self):
        """Elimina el componente seleccionado"""
        seleccion = self.tree_inventario.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione un componente para eliminar")
            return
            
        componente_id = self.tree_inventario.item(seleccion[0], 'values')[0]
        componente_nombre = self.tree_inventario.item(seleccion[0], 'values')[1]
        
        if messagebox.askyesno("Confirmar", f"¿Eliminar el componente '{componente_nombre}'?"):
            try:
                self.c.execute("DELETE FROM inventario WHERE id = ?", (componente_id,))
                self.conn.commit()
                
                # Registrar en el historial de accesos
                self.registrar_acceso(f"Eliminó componente del inventario: {componente_nombre}")
                
                messagebox.showinfo("Éxito", "Componente eliminado")
                self.cargar_ubicaciones_inventario()
                self.actualizar_inventario()
            except sqlite3.Error as e:
                messagebox.showerror("Error", f"No se pudo eliminar el componente: {e}")
    
    def exportar_inventario(self):
        """Exporta el inventario a un archivo Excel"""
        try:
            # Obtener todos los componentes
            items = self.tree_inventario.get_children()
            data = [self.tree_inventario.item(item, 'values') for item in items]
            
            if not data:
                messagebox.showwarning("Advertencia", "No hay datos para exportar")
                return
                
            # Crear DataFrame
            df = pd.DataFrame(data, columns=["ID", "Componente", "Tipo", "Cantidad", "Mínimo", "Proveedor", "Ubicación", "Últ. Actualización"])
            
            # Preguntar dónde guardar
            filepath = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Guardar inventario como"
            )
            
            if not filepath:
                return
                
            # Guardar en Excel
            df.to_excel(filepath, index=False)
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Exportó inventario a {filepath}")
            
            messagebox.showinfo("Éxito", f"Inventario exportado a:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar: {e}")
    
    def crear_graficos_inventario(self, frame):
        """Crea gráficos estadísticos del inventario"""
        try:
            # Obtener datos para gráficos
            self.c.execute("SELECT componente, cantidad FROM inventario ORDER BY cantidad DESC LIMIT 10")
            top_componentes = self.c.fetchall()
            
            self.c.execute("SELECT tipo, SUM(cantidad) FROM inventario GROUP BY tipo")
            tipos_componentes = self.c.fetchall()
            
            self.c.execute("SELECT proveedor, COUNT(*) FROM inventario WHERE proveedor IS NOT NULL AND proveedor != '' GROUP BY proveedor")
            proveedores = self.c.fetchall()
            
            # Crear figura con subplots
            fig, (ax1, ax2, ax3) = plt.subplots(1, 3, figsize=(15, 5))
            fig.suptitle('Estadísticas de Inventario', fontsize=14)
            
            # Gráfico 1: Top componentes por cantidad
            if top_componentes:
                nombres = [comp[0] for comp in top_componentes]
                cantidades = [comp[1] for comp in top_componentes]
                
                ax1.barh(nombres, cantidades, color='skyblue')
                ax1.set_title('Top 10 Componentes')
                ax1.set_xlabel('Cantidad')
            else:
                ax1.text(0.5, 0.5, 'No hay datos\ndisponibles', ha='center', va='center')
                ax1.set_title('Top 10 Componentes')
            
            # Gráfico 2: Distribución por tipo
            if tipos_componentes:
                tipos = [tipo[0] for tipo in tipos_componentes]
                cantidades = [tipo[1] for tipo in tipos_componentes]
                
                ax2.pie(cantidades, labels=tipos, autopct='%1.1f%%', startangle=90)
                ax2.set_title('Distribución por Tipo')
            else:
                ax2.text(0.5, 0.5, 'No hay datos\ndisponibles', ha='center', va='center')
                ax2.set_title('Distribución por Tipo')
            
            # Gráfico 3: Distribución por proveedor
            if proveedores and len(proveedores) > 1:
                prov_nombres = [prov[0] for prov in proveedores]
                prov_cantidades = [prov[1] for prov in proveedores]
                
                ax3.bar(prov_nombres, prov_cantidades, color='lightgreen')
                ax3.set_title('Componentes por Proveedor')
                ax3.tick_params(axis='x', rotation=45)
            else:
                ax3.text(0.5, 0.5, 'No hay suficientes\ndatos de proveedores', ha='center', va='center')
                ax3.set_title('Componentes por Proveedor')
            
            fig.tight_layout()
            
            # Integrar gráfico en Tkinter
            canvas = FigureCanvasTkAgg(fig, master=frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            
            # Guardar referencia para evitar garbage collection
            self.canvas_inventario = canvas
            
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudieron generar los gráficos: {e}")
    
    # ------------------------- Pestaña de Gestión -------------------------
    def inicializar_gestion(self):
        """Configura los componentes de la pestaña de gestión del laboratorio"""
        # Frame para controles principales
        frame_controles = ttk.Frame(self.frame_gestion)
        frame_controles.pack(fill='x', padx=10, pady=5)
        
        # Pestañas internas para gestión
        self.notebook_gestion = ttk.Notebook(frame_controles)
        self.notebook_gestion.pack(fill='both', expand=True)
        
        # Frames para cada subpestaña
        self.frame_equipos = ttk.Frame(self.notebook_gestion)
        self.frame_reservas = ttk.Frame(self.notebook_gestion)
        self.frame_mantenimiento = ttk.Frame(self.notebook_gestion)
        self.frame_configuracion = ttk.Frame(self.notebook_gestion)
        
        self.notebook_gestion.add(self.frame_equipos, text="Gestión de Equipos")
        self.notebook_gestion.add(self.frame_reservas, text="Reservas")
        self.notebook_gestion.add(self.frame_mantenimiento, text="Mantenimiento")
        self.notebook_gestion.add(self.frame_configuracion, text="Configuración")
        
        # Inicializar subpestañas
        self.inicializar_gestion_equipos()
        self.inicializar_gestion_reservas()
        self.inicializar_gestion_mantenimiento()
        self.inicializar_configuracion()
        
        # Frame para estadísticas
        frame_estadisticas = ttk.LabelFrame(self.frame_gestion, text="Estadísticas del Laboratorio", padding=10)
        frame_estadisticas.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Crear gráficos de estadísticas
        self.crear_estadisticas_laboratorio(frame_estadisticas)
    
    def inicializar_gestion_equipos(self):
        """Configura la subpestaña de gestión de equipos"""
        # Frame para controles
        frame_controles = ttk.Frame(self.frame_equipos)
        frame_controles.pack(fill='x', padx=10, pady=5)
        
        # Botones de acción
        btn_agregar = ttk.Button(frame_controles, text="Agregar Equipo", command=self.agregar_equipo)
        btn_agregar.pack(side='left', padx=5)
        
        btn_editar = ttk.Button(frame_controles, text="Editar", command=self.editar_equipo)
        btn_editar.pack(side='left', padx=5)
        
        btn_eliminar = ttk.Button(frame_controles, text="Eliminar", command=self.eliminar_equipo)
        btn_eliminar.pack(side='left', padx=5)
        
        btn_actualizar = ttk.Button(frame_controles, text="Actualizar", command=self.actualizar_equipos)
        btn_actualizar.pack(side='left', padx=5)
        
        btn_exportar = ttk.Button(frame_controles, text="Exportar", command=self.exportar_equipos)
        btn_exportar.pack(side='left', padx=5)
        
        # Frame para filtros
        frame_filtros = ttk.LabelFrame(self.frame_equipos, text="Filtros", padding=5)
        frame_filtros.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(frame_filtros, text="Tipo:").grid(row=0, column=0, padx=5, pady=2, sticky='e')
        self.combo_tipo_equipo = ttk.Combobox(frame_filtros, values=["Todos", "Computadora", "Servidor", "Switch", "Router", "Impresora", "Otros"], state='readonly')
        self.combo_tipo_equipo.grid(row=0, column=1, padx=5, pady=2, sticky='we')
        self.combo_tipo_equipo.set("Todos")
        self.combo_tipo_equipo.bind("<<ComboboxSelected>>", lambda e: self.actualizar_equipos())
        
        ttk.Label(frame_filtros, text="Estado:").grid(row=0, column=2, padx=5, pady=2, sticky='e')
        self.combo_estado_equipo = ttk.Combobox(frame_filtros, values=["Todos", "Operativo", "Mantenimiento", "Dañado", "Retirado"], state='readonly')
        self.combo_estado_equipo.grid(row=0, column=3, padx=5, pady=2, sticky='we')
        self.combo_estado_equipo.set("Todos")
        self.combo_estado_equipo.bind("<<ComboboxSelected>>", lambda e: self.actualizar_equipos())
        
        ttk.Label(frame_filtros, text="Ubicación:").grid(row=0, column=4, padx=5, pady=2, sticky='e')
        self.combo_ubicacion_equipo = ttk.Combobox(frame_filtros, state='readonly')
        self.combo_ubicacion_equipo.grid(row=0, column=5, padx=5, pady=2, sticky='we')
        self.combo_ubicacion_equipo.set("Todos")
        self.combo_ubicacion_equipo.bind("<<ComboboxSelected>>", lambda e: self.actualizar_equipos())
        
        # Cargar ubicaciones disponibles
        self.cargar_ubicaciones_equipos()
        
        # Frame para la tabla de equipos
        frame_tabla = ttk.Frame(self.frame_equipos)
        frame_tabla.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Tabla de equipos
        columns = ("ID", "Nombre", "Tipo", "Modelo", "Serial", "Estado", "Ubicación", "Adquisición", "Últ. Mant.")
        self.tree_equipos = ttk.Treeview(frame_tabla, columns=columns, show='headings', selectmode='browse')
        
        for col in columns:
            self.tree_equipos.heading(col, text=col)
            self.tree_equipos.column(col, width=100, anchor='center')
        
        self.tree_equipos.column("Nombre", width=150)
        self.tree_equipos.column("Modelo", width=120)
        self.tree_equipos.column("Serial", width=120)
        self.tree_equipos.column("Ubicación", width=100)
        self.tree_equipos.pack(fill='both', expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree_equipos.yview)
        scrollbar.pack(side='right', fill='y')
        self.tree_equipos.configure(yscrollcommand=scrollbar.set)
        
        # Cargar datos iniciales
        self.actualizar_equipos()
    
    def cargar_ubicaciones_equipos(self):
        """Carga las ubicaciones disponibles para filtrar"""
        try:
            self.c.execute("SELECT DISTINCT ubicacion FROM equipos WHERE ubicacion IS NOT NULL AND ubicacion != '' ORDER BY ubicacion")
            ubicaciones = ["Todos"] + [u[0] for u in self.c.fetchall()]
            self.combo_ubicacion_equipo['values'] = ubicaciones
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudieron cargar las ubicaciones: {e}")
    
    def actualizar_equipos(self):
        """Actualiza la tabla de equipos con los datos de la base de datos"""
        try:
            # Limpiar tabla
            for item in self.tree_equipos.get_children():
                self.tree_equipos.delete(item)
            
            # Construir consulta con filtros
            query = """SELECT id, nombre, tipo, modelo, serial, estado, ubicacion, 
                        fecha_adquisicion, ultimo_mantenimiento 
                        FROM equipos WHERE 1=1"""
            params = []
            
            tipo = self.combo_tipo_equipo.get()
            if tipo != "Todos":
                query += " AND tipo = ?"
                params.append(tipo)
                
            estado = self.combo_estado_equipo.get()
            if estado != "Todos":
                query += " AND estado = ?"
                params.append(estado)
                
            ubicacion = self.combo_ubicacion_equipo.get()
            if ubicacion != "Todos":
                query += " AND ubicacion = ?"
                params.append(ubicacion)
            
            query += " ORDER BY nombre"
            
            # Obtener datos
            self.c.execute(query, params)
            equipos = self.c.fetchall()
            
            # Llenar tabla
            for equipo in equipos:
                self.tree_equipos.insert('', 'end', values=equipo)
                
            # Resaltar equipos con estado diferente a "Operativo"
            for item in self.tree_equipos.get_children():
                valores = self.tree_equipos.item(item, 'values')
                if valores[5] != "Operativo":
                    self.tree_equipos.tag_configure('no_operativo', background='#ffcccc')
                    self.tree_equipos.item(item, tags=('no_operativo',))
                    
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo cargar los equipos: {e}")
    
    def agregar_equipo(self):
        """Abre ventana para agregar nuevo equipo"""
        ventana = tk.Toplevel(self.root)
        ventana.title("Agregar Equipo")
        ventana.geometry("600x500")
        ventana.grab_set()
        
        # Variables
        self.var_nombre = tk.StringVar()
        self.var_tipo = tk.StringVar(value="Computadora")
        self.var_modelo = tk.StringVar()
        self.var_serial = tk.StringVar()
        self.var_estado = tk.StringVar(value="Operativo")
        self.var_ubicacion = tk.StringVar()
        self.var_fecha_adq = tk.StringVar()
        self.var_ult_mant = tk.StringVar()
        self.var_observaciones = tk.StringVar()
        
        # Frame principal con scrollbar
        frame_principal = ttk.Frame(ventana)
        frame_principal.pack(fill='both', expand=True)
        
        canvas = tk.Canvas(frame_principal)
        scrollbar = ttk.Scrollbar(frame_principal, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Campos del formulario
        ttk.Label(scrollable_frame, text="Nombre:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
        ttk.Entry(scrollable_frame, textvariable=self.var_nombre).grid(row=0, column=1, padx=10, pady=5, sticky='we')
        
        ttk.Label(scrollable_frame, text="Tipo:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
        combo_tipo = ttk.Combobox(scrollable_frame, textvariable=self.var_tipo, 
                                 values=["Computadora", "Servidor", "Switch", "Router", "Impresora", "Otros"], 
                                 state='readonly')
        combo_tipo.grid(row=1, column=1, padx=10, pady=5, sticky='we')
        
        ttk.Label(scrollable_frame, text="Modelo:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
        ttk.Entry(scrollable_frame, textvariable=self.var_modelo).grid(row=2, column=1, padx=10, pady=5, sticky='we')
        
        ttk.Label(scrollable_frame, text="Número de Serie:").grid(row=3, column=0, padx=10, pady=5, sticky='e')
        ttk.Entry(scrollable_frame, textvariable=self.var_serial).grid(row=3, column=1, padx=10, pady=5, sticky='we')
        
        ttk.Label(scrollable_frame, text="Estado:").grid(row=4, column=0, padx=10, pady=5, sticky='e')
        combo_estado = ttk.Combobox(scrollable_frame, textvariable=self.var_estado, 
                                   values=["Operativo", "Mantenimiento", "Dañado", "Retirado"], 
                                   state='readonly')
        combo_estado.grid(row=4, column=1, padx=10, pady=5, sticky='we')
        
        ttk.Label(scrollable_frame, text="Ubicación:").grid(row=5, column=0, padx=10, pady=5, sticky='e')
        ttk.Entry(scrollable_frame, textvariable=self.var_ubicacion).grid(row=5, column=1, padx=10, pady=5, sticky='we')
        
        ttk.Label(scrollable_frame, text="Fecha de Adquisición:").grid(row=6, column=0, padx=10, pady=5, sticky='e')
        ttk.Entry(scrollable_frame, textvariable=self.var_fecha_adq).grid(row=6, column=1, padx=10, pady=5, sticky='we')
        
        ttk.Label(scrollable_frame, text="Último Mantenimiento:").grid(row=7, column=0, padx=10, pady=5, sticky='e')
        ttk.Entry(scrollable_frame, textvariable=self.var_ult_mant).grid(row=7, column=1, padx=10, pady=5, sticky='we')
        
        ttk.Label(scrollable_frame, text="Observaciones:").grid(row=8, column=0, padx=10, pady=5, sticky='ne')
        text_obs = tk.Text(scrollable_frame, height=5, width=30, wrap='word')
        text_obs.grid(row=8, column=1, padx=10, pady=5, sticky='we')
        
        # Botones
        frame_botones = ttk.Frame(ventana)
        frame_botones.pack(fill='x', padx=10, pady=10)
        
        btn_guardar = ttk.Button(frame_botones, text="Guardar", command=lambda: self.guardar_equipo(
            self.var_nombre.get(),
            self.var_tipo.get(),
            self.var_modelo.get(),
            self.var_serial.get(),
            self.var_estado.get(),
            self.var_ubicacion.get(),
            self.var_fecha_adq.get(),
            self.var_ult_mant.get(),
            text_obs.get("1.0", "end").strip(),
            ventana
        ))
        btn_guardar.pack(side='left', padx=5)
        
        btn_cancelar = ttk.Button(frame_botones, text="Cancelar", command=ventana.destroy)
        btn_cancelar.pack(side='left', padx=5)
    
    def guardar_equipo(self, nombre, tipo, modelo, serial, estado, ubicacion, fecha_adq, ult_mant, observaciones, ventana):
        """Guarda un nuevo equipo en la base de datos"""
        if not nombre:
            messagebox.showwarning("Advertencia", "El nombre del equipo es obligatorio")
            return
            
        try:
            self.c.execute("""INSERT INTO equipos 
                              (nombre, tipo, modelo, serial, estado, ubicacion, 
                               fecha_adquisicion, ultimo_mantenimiento, observaciones) 
                              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                          (nombre, tipo, modelo, serial, estado, ubicacion, fecha_adq, ult_mant, observaciones))
            self.conn.commit()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Agregó nuevo equipo: {nombre}")
            
            messagebox.showinfo("Éxito", "Equipo agregado al sistema")
            ventana.destroy()
            self.cargar_ubicaciones_equipos()
            self.actualizar_equipos()
        except sqlite3.Error as e:
            if "UNIQUE" in str(e):
                messagebox.showerror("Error", "El número de serie ya existe en el sistema")
            else:
                messagebox.showerror("Error", f"No se pudo guardar el equipo: {e}")
    
    def editar_equipo(self):
        """Abre ventana para editar equipo seleccionado"""
        seleccion = self.tree_equipos.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione un equipo para editar")
            return
            
        equipo_id = self.tree_equipos.item(seleccion[0], 'values')[0]
        
        try:
            self.c.execute("""SELECT nombre, tipo, modelo, serial, estado, ubicacion, 
                              fecha_adquisicion, ultimo_mantenimiento, observaciones 
                              FROM equipos WHERE id = ?""", (equipo_id,))
            equipo = self.c.fetchone()
            
            ventana = tk.Toplevel(self.root)
            ventana.title("Editar Equipo")
            ventana.geometry("600x500")
            ventana.grab_set()
            
            # Variables
            self.var_nombre = tk.StringVar(value=equipo[0])
            self.var_tipo = tk.StringVar(value=equipo[1])
            self.var_modelo = tk.StringVar(value=equipo[2])
            self.var_serial = tk.StringVar(value=equipo[3])
            self.var_estado = tk.StringVar(value=equipo[4])
            self.var_ubicacion = tk.StringVar(value=equipo[5] if equipo[5] else "")
            self.var_fecha_adq = tk.StringVar(value=equipo[6] if equipo[6] else "")
            self.var_ult_mant = tk.StringVar(value=equipo[7] if equipo[7] else "")
            self.var_observaciones = tk.StringVar(value=equipo[8] if equipo[8] else "")
            
            # Frame principal con scrollbar
            frame_principal = ttk.Frame(ventana)
            frame_principal.pack(fill='both', expand=True)
            
            canvas = tk.Canvas(frame_principal)
            scrollbar = ttk.Scrollbar(frame_principal, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(
                    scrollregion=canvas.bbox("all")
                )
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Campos del formulario
            ttk.Label(scrollable_frame, text="Nombre:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
            ttk.Entry(scrollable_frame, textvariable=self.var_nombre).grid(row=0, column=1, padx=10, pady=5, sticky='we')
            
            ttk.Label(scrollable_frame, text="Tipo:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
            combo_tipo = ttk.Combobox(scrollable_frame, textvariable=self.var_tipo, 
                                     values=["Computadora", "Servidor", "Switch", "Router", "Impresora", "Otros"], 
                                     state='readonly')
            combo_tipo.grid(row=1, column=1, padx=10, pady=5, sticky='we')
            
            ttk.Label(scrollable_frame, text="Modelo:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
            ttk.Entry(scrollable_frame, textvariable=self.var_modelo).grid(row=2, column=1, padx=10, pady=5, sticky='we')
            
            ttk.Label(scrollable_frame, text="Número de Serie:").grid(row=3, column=0, padx=10, pady=5, sticky='e')
            ttk.Entry(scrollable_frame, textvariable=self.var_serial).grid(row=3, column=1, padx=10, pady=5, sticky='we')
            
            ttk.Label(scrollable_frame, text="Estado:").grid(row=4, column=0, padx=10, pady=5, sticky='e')
            combo_estado = ttk.Combobox(scrollable_frame, textvariable=self.var_estado, 
                                       values=["Operativo", "Mantenimiento", "Dañado", "Retirado"], 
                                       state='readonly')
            combo_estado.grid(row=4, column=1, padx=10, pady=5, sticky='we')
            
            ttk.Label(scrollable_frame, text="Ubicación:").grid(row=5, column=0, padx=10, pady=5, sticky='e')
            ttk.Entry(scrollable_frame, textvariable=self.var_ubicacion).grid(row=5, column=1, padx=10, pady=5, sticky='we')
            
            ttk.Label(scrollable_frame, text="Fecha de Adquisición:").grid(row=6, column=0, padx=10, pady=5, sticky='e')
            ttk.Entry(scrollable_frame, textvariable=self.var_fecha_adq).grid(row=6, column=1, padx=10, pady=5, sticky='we')
            
            ttk.Label(scrollable_frame, text="Último Mantenimiento:").grid(row=7, column=0, padx=10, pady=5, sticky='e')
            ttk.Entry(scrollable_frame, textvariable=self.var_ult_mant).grid(row=7, column=1, padx=10, pady=5, sticky='we')
            
            ttk.Label(scrollable_frame, text="Observaciones:").grid(row=8, column=0, padx=10, pady=5, sticky='ne')
            text_obs = tk.Text(scrollable_frame, height=5, width=30, wrap='word')
            text_obs.insert('1.0', equipo[8] if equipo[8] else "")
            text_obs.grid(row=8, column=1, padx=10, pady=5, sticky='we')
            
            # Botones
            frame_botones = ttk.Frame(ventana)
            frame_botones.pack(fill='x', padx=10, pady=10)
            
            btn_guardar = ttk.Button(frame_botones, text="Guardar", 
                                    command=lambda: self.actualizar_equipo(
                                        equipo_id,
                                        self.var_nombre.get(),
                                        self.var_tipo.get(),
                                        self.var_modelo.get(),
                                        self.var_serial.get(),
                                        self.var_estado.get(),
                                        self.var_ubicacion.get(),
                                        self.var_fecha_adq.get(),
                                        self.var_ult_mant.get(),
                                        text_obs.get("1.0", "end").strip(),
                                        ventana
                                    ))
            btn_guardar.pack(side='left', padx=5)
            
            btn_cancelar = ttk.Button(frame_botones, text="Cancelar", command=ventana.destroy)
            btn_cancelar.pack(side='left', padx=5)
            
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo cargar el equipo: {e}")
    
    def actualizar_equipo(self, equipo_id, nombre, tipo, modelo, serial, estado, ubicacion, fecha_adq, ult_mant, observaciones, ventana):
        """Actualiza un equipo en la base de datos"""
        if not nombre:
            messagebox.showwarning("Advertencia", "El nombre del equipo es obligatorio")
            return
            
        try:
            self.c.execute("""UPDATE equipos SET 
                              nombre = ?, tipo = ?, modelo = ?, serial = ?, estado = ?, 
                              ubicacion = ?, fecha_adquisicion = ?, ultimo_mantenimiento = ?, observaciones = ? 
                              WHERE id = ?""",
                          (nombre, tipo, modelo, serial, estado, ubicacion, fecha_adq, ult_mant, observaciones, equipo_id))
            self.conn.commit()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Actualizó equipo: {nombre}")
            
            messagebox.showinfo("Éxito", "Equipo actualizado")
            ventana.destroy()
            self.cargar_ubicaciones_equipos()
            self.actualizar_equipos()
        except sqlite3.Error as e:
            if "UNIQUE" in str(e):
                messagebox.showerror("Error", "El número de serie ya existe en el sistema")
            else:
                messagebox.showerror("Error", f"No se pudo actualizar el equipo: {e}")
    
    def eliminar_equipo(self):
        """Elimina el equipo seleccionado"""
        seleccion = self.tree_equipos.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione un equipo para eliminar")
            return
            
        equipo_id = self.tree_equipos.item(seleccion[0], 'values')[0]
        equipo_nombre = self.tree_equipos.item(seleccion[0], 'values')[1]
        
        # Verificar si hay reportes asociados
        self.c.execute("SELECT COUNT(*) FROM reportes WHERE equipo_id = ?", (equipo_id,))
        count_reportes = self.c.fetchone()[0]
        
        # Verificar si hay reservas asociadas
        self.c.execute("SELECT COUNT(*) FROM reservas WHERE equipo_id = ?", (equipo_id,))
        count_reservas = self.c.fetchone()[0]
        
        mensaje = f"¿Eliminar el equipo '{equipo_nombre}'?"
        if count_reportes > 0 or count_reservas > 0:
            mensaje += "\n\nAdvertencia: Este equipo tiene:"
            if count_reportes > 0:
                mensaje += f"\n- {count_reportes} reportes asociados"
            if count_reservas > 0:
                mensaje += f"\n- {count_reservas} reservas asociadas"
            mensaje += "\n\nTodos estos registros también serán eliminados."
        
        if messagebox.askyesno("Confirmar", mensaje):
            try:
                # Eliminar registros asociados primero (por las claves foráneas)
                if count_reportes > 0:
                    self.c.execute("DELETE FROM reportes WHERE equipo_id = ?", (equipo_id,))
                
                if count_reservas > 0:
                    self.c.execute("DELETE FROM reservas WHERE equipo_id = ?", (equipo_id,))
                
                # Eliminar el equipo
                self.c.execute("DELETE FROM equipos WHERE id = ?", (equipo_id,))
                self.conn.commit()
                
                # Registrar en el historial de accesos
                self.registrar_acceso(f"Eliminó equipo: {equipo_nombre}")
                
                messagebox.showinfo("Éxito", "Equipo eliminado")
                self.cargar_ubicaciones_equipos()
                self.actualizar_equipos()
            except sqlite3.Error as e:
                messagebox.showerror("Error", f"No se pudo eliminar el equipo: {e}")
    
    def exportar_equipos(self):
        """Exporta los equipos a un archivo Excel"""
        try:
            # Obtener todos los equipos visibles
            items = self.tree_equipos.get_children()
            data = [self.tree_equipos.item(item, 'values') for item in items]
            
            if not data:
                messagebox.showwarning("Advertencia", "No hay datos para exportar")
                return
                
            # Crear DataFrame
            df = pd.DataFrame(data, columns=["ID", "Nombre", "Tipo", "Modelo", "Serial", "Estado", "Ubicación", "Adquisición", "Últ. Mant."])
            
            # Preguntar dónde guardar
            filepath = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Guardar equipos como"
            )
            
            if not filepath:
                return
                
            # Guardar en Excel
            df.to_excel(filepath, index=False)
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Exportó lista de equipos a {filepath}")
            
            messagebox.showinfo("Éxito", f"Equipos exportados a:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar: {e}")
    
    def inicializar_gestion_reservas(self):
        """Configura la subpestaña de gestión de reservas"""
        # Frame para controles
        frame_controles = ttk.Frame(self.frame_reservas)
        frame_controles.pack(fill='x', padx=10, pady=5)
        
        # Botones de acción
        btn_nueva = ttk.Button(frame_controles, text="Nueva Reserva", command=self.nueva_reserva)
        btn_nueva.pack(side='left', padx=5)
        
        btn_cancelar = ttk.Button(frame_controles, text="Cancelar Reserva", command=self.cancelar_reserva)
        btn_cancelar.pack(side='left', padx=5)
        
        btn_actualizar = ttk.Button(frame_controles, text="Actualizar", command=self.actualizar_reservas)
        btn_actualizar.pack(side='left', padx=5)
        
        btn_exportar = ttk.Button(frame_controles, text="Exportar", command=self.exportar_reservas)
        btn_exportar.pack(side='left', padx=5)
        
        # Frame para filtros
        frame_filtros = ttk.LabelFrame(self.frame_reservas, text="Filtros", padding=5)
        frame_filtros.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(frame_filtros, text="Estado:").grid(row=0, column=0, padx=5, pady=2, sticky='e')
        self.combo_estado_reserva = ttk.Combobox(frame_filtros, values=["Todos", "Confirmada", "Cancelada", "Completada"], state='readonly')
        self.combo_estado_reserva.grid(row=0, column=1, padx=5, pady=2, sticky='we')
        self.combo_estado_reserva.set("Todos")
        self.combo_estado_reserva.bind("<<ComboboxSelected>>", lambda e: self.actualizar_reservas())
        
        ttk.Label(frame_filtros, text="Fecha Desde:").grid(row=0, column=2, padx=5, pady=2, sticky='e')
        self.entry_fecha_desde_reserva = ttk.Entry(frame_filtros)
        self.entry_fecha_desde_reserva.grid(row=0, column=3, padx=5, pady=2, sticky='we')
        
        ttk.Label(frame_filtros, text="Fecha Hasta:").grid(row=0, column=4, padx=5, pady=2, sticky='e')
        self.entry_fecha_hasta_reserva = ttk.Entry(frame_filtros)
        self.entry_fecha_hasta_reserva.grid(row=0, column=5, padx=5, pady=2, sticky='we')
        
        btn_buscar = ttk.Button(frame_filtros, text="Buscar", command=self.actualizar_reservas)
        btn_buscar.grid(row=0, column=6, padx=5, pady=2)
        
        # Frame para la tabla de reservas
        frame_tabla = ttk.Frame(self.frame_reservas)
        frame_tabla.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Tabla de reservas
        columns = ("ID", "Equipo", "Usuario", "Fecha Inicio", "Fecha Fin", "Propósito", "Estado")
        self.tree_reservas = ttk.Treeview(frame_tabla, columns=columns, show='headings', selectmode='browse')
        
        for col in columns:
            self.tree_reservas.heading(col, text=col)
            self.tree_reservas.column(col, width=100, anchor='center')
        
        self.tree_reservas.column("Equipo", width=150)
        self.tree_reservas.column("Usuario", width=120)
        self.tree_reservas.column("Propósito", width=200)
        self.tree_reservas.pack(fill='both', expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree_reservas.yview)
        scrollbar.pack(side='right', fill='y')
        self.tree_reservas.configure(yscrollcommand=scrollbar.set)
        
        # Cargar datos iniciales
        self.actualizar_reservas()
    
    def actualizar_reservas(self):
        """Actualiza la tabla de reservas con los datos de la base de datos"""
        try:
            # Limpiar tabla
            for item in self.tree_reservas.get_children():
                self.tree_reservas.delete(item)
            
            # Construir consulta con filtros
            query = """SELECT r.id, e.nombre, u.nombre, r.fecha_inicio, r.fecha_fin, 
                        r.proposito, r.estado 
                        FROM reservas r 
                        JOIN equipos e ON r.equipo_id = e.id 
                        JOIN usuarios u ON r.usuario_id = u.id 
                        WHERE 1=1"""
            params = []
            
            estado = self.combo_estado_reserva.get()
            if estado != "Todos":
                query += " AND r.estado = ?"
                params.append(estado)
                
            fecha_desde = self.entry_fecha_desde_reserva.get()
            if fecha_desde:
                query += " AND r.fecha_inicio >= ?"
                params.append(fecha_desde)
                
            fecha_hasta = self.entry_fecha_hasta_reserva.get()
            if fecha_hasta:
                query += " AND r.fecha_fin <= ?"
                params.append(fecha_hasta)
            
            query += " ORDER BY r.fecha_inicio"
            
            # Obtener datos
            self.c.execute(query, params)
            reservas = self.c.fetchall()
            
            # Llenar tabla
            for reserva in reservas:
                self.tree_reservas.insert('', 'end', values=reserva)
                
            # Resaltar reservas activas
            fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for item in self.tree_reservas.get_children():
                valores = self.tree_reservas.item(item, 'values')
                if valores[3] <= fecha_actual <= valores[4] and valores[6] == "Confirmada":
                    self.tree_reservas.tag_configure('activa', background='#ccffcc')
                    self.tree_reservas.item(item, tags=('activa',))
                    
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo cargar las reservas: {e}")
    
    def nueva_reserva(self):
        """Abre ventana para crear nueva reserva"""
        ventana = tk.Toplevel(self.root)
        ventana.title("Nueva Reserva")
        ventana.geometry("500x400")
        ventana.grab_set()
        
        # Variables
        self.var_equipo = tk.StringVar()
        self.var_usuario = tk.StringVar()
        self.var_fecha_ini = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d %H:%M"))
        self.var_fecha_fin = tk.StringVar(value=(datetime.now() + timedelta(hours=2)).strftime("%Y-%m-%d %H:%M"))
        self.var_proposito = tk.StringVar()
        
        # Frame principal
        frame_principal = ttk.Frame(ventana, padding=10)
        frame_principal.pack(fill='both', expand=True)
        
        # Campos del formulario
        ttk.Label(frame_principal, text="Equipo:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        combo_equipo = ttk.Combobox(frame_principal, textvariable=self.var_equipo, state='readonly')
        combo_equipo.grid(row=0, column=1, padx=5, pady=5, sticky='we')
        
        # Cargar equipos disponibles
        self.c.execute("SELECT id, nombre FROM equipos WHERE estado = 'Operativo' ORDER BY nombre")
        equipos = self.c.fetchall()
        combo_equipo['values'] = [f"{e[1]} (ID: {e[0]})" for e in equipos]
        
        ttk.Label(frame_principal, text="Usuario:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        combo_usuario = ttk.Combobox(frame_principal, textvariable=self.var_usuario, state='readonly')
        combo_usuario.grid(row=1, column=1, padx=5, pady=5, sticky='we')
        
        # Cargar usuarios disponibles
        self.c.execute("SELECT id, nombre FROM usuarios ORDER BY nombre")
        usuarios = self.c.fetchall()
        combo_usuario['values'] = [f"{u[1]} (ID: {u[0]})" for u in usuarios]
        
        ttk.Label(frame_principal, text="Fecha Inicio:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_fecha_ini).grid(row=2, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Fecha Fin:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_fecha_fin).grid(row=3, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Propósito:").grid(row=4, column=0, padx=5, pady=5, sticky='ne')
        text_proposito = tk.Text(frame_principal, height=5, width=40, wrap='word')
        text_proposito.grid(row=4, column=1, padx=5, pady=5, sticky='we')
        
        # Botones
        frame_botones = ttk.Frame(ventana)
        frame_botones.pack(fill='x', padx=10, pady=10)
        
        btn_guardar = ttk.Button(frame_botones, text="Guardar", command=lambda: self.guardar_reserva(
            self.var_equipo.get().split("(ID: ")[1].rstrip(")") if self.var_equipo.get() else None,
            self.var_usuario.get().split("(ID: ")[1].rstrip(")") if self.var_usuario.get() else None,
            self.var_fecha_ini.get(),
            self.var_fecha_fin.get(),
            text_proposito.get("1.0", "end").strip(),
            ventana
        ))
        btn_guardar.pack(side='left', padx=5)
        
        btn_cancelar = ttk.Button(frame_botones, text="Cancelar", command=ventana.destroy)
        btn_cancelar.pack(side='left', padx=5)
    
    def guardar_reserva(self, equipo_id, usuario_id, fecha_ini, fecha_fin, proposito, ventana):
        """Guarda una nueva reserva en la base de datos"""
        if not equipo_id:
            messagebox.showwarning("Advertencia", "Seleccione un equipo")
            return
        if not usuario_id:
            messagebox.showwarning("Advertencia", "Seleccione un usuario")
            return
        if not fecha_ini or not fecha_fin:
            messagebox.showwarning("Advertencia", "Las fechas son obligatorias")
            return
            
        try:
            # Verificar disponibilidad del equipo
            self.c.execute("""SELECT COUNT(*) FROM reservas 
                              WHERE equipo_id = ? AND estado = 'Confirmada' 
                              AND ((? BETWEEN fecha_inicio AND fecha_fin) 
                              OR (? BETWEEN fecha_inicio AND fecha_fin))""",
                          (equipo_id, fecha_ini, fecha_fin))
            count = self.c.fetchone()[0]
            
            if count > 0:
                messagebox.showerror("Error", "El equipo no está disponible en ese horario")
                return
                
            # Insertar reserva
            fecha_solicitud = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.c.execute("""INSERT INTO reservas 
                              (equipo_id, usuario_id, fecha_inicio, fecha_fin, 
                               proposito, estado, fecha_solicitud) 
                              VALUES (?, ?, ?, ?, ?, ?, ?)""",
                          (equipo_id, usuario_id, fecha_ini, fecha_fin, proposito, "Confirmada", fecha_solicitud))
            self.conn.commit()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Creó reserva para equipo ID: {equipo_id}")
            
            messagebox.showinfo("Éxito", "Reserva creada correctamente")
            ventana.destroy()
            self.actualizar_reservas()
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo guardar la reserva: {e}")
    
    def cancelar_reserva(self):
        """Cancela la reserva seleccionada"""
        seleccion = self.tree_reservas.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione una reserva para cancelar")
            return
            
        reserva_id = self.tree_reservas.item(seleccion[0], 'values')[0]
        estado = self.tree_reservas.item(seleccion[0], 'values')[6]
        
        if estado != "Confirmada":
            messagebox.showwarning("Advertencia", "Solo se pueden cancelar reservas confirmadas")
            return
            
        if messagebox.askyesno("Confirmar", "¿Cancelar esta reserva?"):
            try:
                self.c.execute("UPDATE reservas SET estado = 'Cancelada' WHERE id = ?", (reserva_id,))
                self.conn.commit()
                
                # Registrar en el historial de accesos
                self.registrar_acceso(f"Canceló reserva ID: {reserva_id}")
                
                messagebox.showinfo("Éxito", "Reserva cancelada")
                self.actualizar_reservas()
            except sqlite3.Error as e:
                messagebox.showerror("Error", f"No se pudo cancelar la reserva: {e}")
    
    def exportar_reservas(self):
        """Exporta las reservas a un archivo Excel"""
        try:
            # Obtener todas las reservas visibles
            items = self.tree_reservas.get_children()
            data = [self.tree_reservas.item(item, 'values') for item in items]
            
            if not data:
                messagebox.showwarning("Advertencia", "No hay datos para exportar")
                return
                
            # Crear DataFrame
            df = pd.DataFrame(data, columns=["ID", "Equipo", "Usuario", "Fecha Inicio", "Fecha Fin", "Propósito", "Estado"])
            
            # Preguntar dónde guardar
            filepath = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Guardar reservas como"
            )
            
            if not filepath:
                return
                
            # Guardar en Excel
            df.to_excel(filepath, index=False)
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Exportó lista de reservas a {filepath}")
            
            messagebox.showinfo("Éxito", f"Reservas exportadas a:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar: {e}")
    
    def inicializar_gestion_mantenimiento(self):
        """Configura la subpestaña de gestión de mantenimiento"""
        # Frame para controles
        frame_controles = ttk.Frame(self.frame_mantenimiento)
        frame_controles.pack(fill='x', padx=10, pady=5)
        
        # Botones de acción
        btn_programar = ttk.Button(frame_controles, text="Programar Mantenimiento", 
                                  command=self.programar_mantenimiento)
        btn_programar.pack(side='left', padx=5)
        
        btn_registrar = ttk.Button(frame_controles, text="Registrar Mantenimiento", 
                                 command=self.registrar_mantenimiento)
        btn_registrar.pack(side='left', padx=5)
        
        btn_actualizar = ttk.Button(frame_controles, text="Actualizar", command=self.actualizar_mantenimientos)
        btn_actualizar.pack(side='left', padx=5)
        
        btn_exportar = ttk.Button(frame_controles, text="Exportar", command=self.exportar_mantenimientos)
        btn_exportar.pack(side='left', padx=5)
        
        # Frame para filtros
        frame_filtros = ttk.LabelFrame(self.frame_mantenimiento, text="Filtros", padding=5)
        frame_filtros.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(frame_filtros, text="Tipo:").grid(row=0, column=0, padx=5, pady=2, sticky='e')
        self.combo_tipo_mantenimiento = ttk.Combobox(frame_filtros, values=["Todos", "Preventivo", "Correctivo", "Actualización", "Limpieza"], state='readonly')
        self.combo_tipo_mantenimiento.grid(row=0, column=1, padx=5, pady=2, sticky='we')
        self.combo_tipo_mantenimiento.set("Todos")
        self.combo_tipo_mantenimiento.bind("<<ComboboxSelected>>", lambda e: self.actualizar_mantenimientos())
        
        ttk.Label(frame_filtros, text="Estado:").grid(row=0, column=2, padx=5, pady=2, sticky='e')
        self.combo_estado_mantenimiento = ttk.Combobox(frame_filtros, values=["Todos", "Pendiente", "En Progreso", "Completado", "Cancelado"], state='readonly')
        self.combo_estado_mantenimiento.grid(row=0, column=3, padx=5, pady=2, sticky='we')
        self.combo_estado_mantenimiento.set("Todos")
        self.combo_estado_mantenimiento.bind("<<ComboboxSelected>>", lambda e: self.actualizar_mantenimientos())
        
        # Frame para la tabla de mantenimientos
        frame_tabla = ttk.Frame(self.frame_mantenimiento)
        frame_tabla.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Tabla de mantenimientos
        columns = ("ID", "Equipo", "Tipo", "Fecha Programada", "Fecha Realizado", "Técnico", "Estado")
        self.tree_mantenimientos = ttk.Treeview(frame_tabla, columns=columns, show='headings', selectmode='browse')
        
        for col in columns:
            self.tree_mantenimientos.heading(col, text=col)
            self.tree_mantenimientos.column(col, width=100, anchor='center')
        
        self.tree_mantenimientos.column("Equipo", width=150)
        self.tree_mantenimientos.column("Técnico", width=120)
        self.tree_mantenimientos.pack(fill='both', expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree_mantenimientos.yview)
        scrollbar.pack(side='right', fill='y')
        self.tree_mantenimientos.configure(yscrollcommand=scrollbar.set)
        
        # Cargar datos iniciales
        self.actualizar_mantenimientos()
    
    def actualizar_mantenimientos(self):
        """Actualiza la tabla de mantenimientos con los datos de la base de datos"""
        try:
            # Limpiar tabla
            for item in self.tree_mantenimientos.get_children():
                self.tree_mantenimientos.delete(item)
            
            # Construir consulta con filtros
            query = """SELECT m.id, e.nombre, m.tipo, m.fecha_programada, 
                        m.fecha_realizado, m.tecnico, m.estado 
                        FROM mantenimientos m 
                        JOIN equipos e ON m.equipo_id = e.id 
                        WHERE 1=1"""
            params = []
            
            tipo = self.combo_tipo_mantenimiento.get()
            if tipo != "Todos":
                query += " AND m.tipo = ?"
                params.append(tipo)
                
            estado = self.combo_estado_mantenimiento.get()
            if estado != "Todos":
                query += " AND m.estado = ?"
                params.append(estado)
            
            query += " ORDER BY m.fecha_programada"
            
            # Obtener datos
            self.c.execute(query, params)
            mantenimientos = self.c.fetchall()
            
            # Llenar tabla
            for mant in mantenimientos:
                self.tree_mantenimientos.insert('', 'end', values=mant)
                
            # Resaltar mantenimientos
            fecha_actual = datetime.now().strftime("%Y-%m-%d")
            for item in self.tree_mantenimientos.get_children():
                valores = self.tree_mantenimientos.item(item, 'values')
                
                if valores[6] == "Pendiente" and valores[3] < fecha_actual:
                    self.tree_mantenimientos.tag_configure('atrasado', background='#ff9999')
                    self.tree_mantenimientos.item(item, tags=('atrasado',))
                elif valores[6] == "Pendiente":
                    self.tree_mantenimientos.tag_configure('pendiente', background='#ffff99')
                    self.tree_mantenimientos.item(item, tags=('pendiente',))
                elif valores[6] == "Completado":
                    self.tree_mantenimientos.tag_configure('completado', background='#ccffcc')
                    self.tree_mantenimientos.item(item, tags=('completado',))
                    
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo cargar los mantenimientos: {e}")
    
    def programar_mantenimiento(self):
        """Abre ventana para programar mantenimiento"""
        ventana = tk.Toplevel(self.root)
        ventana.title("Programar Mantenimiento")
        ventana.geometry("500x400")
        ventana.grab_set()
        
        # Variables
        self.var_equipo = tk.StringVar()
        self.var_tipo = tk.StringVar(value="Preventivo")
        self.var_fecha_prog = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.var_tecnico = tk.StringVar()
        self.var_costo = tk.DoubleVar(value=0.0)
        self.var_observaciones = tk.StringVar()
        
        # Frame principal
        frame_principal = ttk.Frame(ventana, padding=10)
        frame_principal.pack(fill='both', expand=True)
        
        # Campos del formulario
        ttk.Label(frame_principal, text="Equipo:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        combo_equipo = ttk.Combobox(frame_principal, textvariable=self.var_equipo, state='readonly')
        combo_equipo.grid(row=0, column=1, padx=5, pady=5, sticky='we')
        
        # Cargar equipos
        self.c.execute("SELECT id, nombre FROM equipos ORDER BY nombre")
        equipos = self.c.fetchall()
        combo_equipo['values'] = [f"{e[1]} (ID: {e[0]})" for e in equipos]
        
        ttk.Label(frame_principal, text="Tipo:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        combo_tipo = ttk.Combobox(frame_principal, textvariable=self.var_tipo, 
                                 values=["Preventivo", "Correctivo", "Actualización", "Limpieza"], 
                                 state='readonly')
        combo_tipo.grid(row=1, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Fecha Programada:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_fecha_prog).grid(row=2, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Técnico:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_tecnico).grid(row=3, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Costo Estimado:").grid(row=4, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_costo).grid(row=4, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Descripción:").grid(row=5, column=0, padx=5, pady=5, sticky='ne')
        text_desc = tk.Text(frame_principal, height=5, width=40, wrap='word')
        text_desc.grid(row=5, column=1, padx=5, pady=5, sticky='we')
        
        # Botones
        frame_botones = ttk.Frame(ventana)
        frame_botones.pack(fill='x', padx=10, pady=10)
        
        btn_guardar = ttk.Button(frame_botones, text="Guardar", command=lambda: self.guardar_mantenimiento(
            self.var_equipo.get().split("(ID: ")[1].rstrip(")") if self.var_equipo.get() else None,
            self.var_tipo.get(),
            self.var_fecha_prog.get(),
            None,  # fecha_realizado
            text_desc.get("1.0", "end").strip(),
            self.var_tecnico.get(),
            "Pendiente",
            self.var_costo.get(),
            "",
            ventana
        ))
        btn_guardar.pack(side='left', padx=5)
        
        btn_cancelar = ttk.Button(frame_botones, text="Cancelar", command=ventana.destroy)
        btn_cancelar.pack(side='left', padx=5)
    
    def registrar_mantenimiento(self):
        """Abre ventana para registrar mantenimiento realizado"""
        seleccion = self.tree_mantenimientos.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione un mantenimiento para registrar")
            return
            
        mantenimiento_id = self.tree_mantenimientos.item(seleccion[0], 'values')[0]
        estado = self.tree_mantenimientos.item(seleccion[0], 'values')[6]
        
        if estado != "Pendiente":
            messagebox.showwarning("Advertencia", "Solo se pueden registrar mantenimientos pendientes")
            return
            
        ventana = tk.Toplevel(self.root)
        ventana.title("Registrar Mantenimiento")
        ventana.geometry("500x300")
        ventana.grab_set()
        
        # Variables
        self.var_fecha_real = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.var_costo_real = tk.DoubleVar(value=0.0)
        self.var_observaciones = tk.StringVar()
        
        # Frame principal
        frame_principal = ttk.Frame(ventana, padding=10)
        frame_principal.pack(fill='both', expand=True)
        
        # Campos del formulario
        ttk.Label(frame_principal, text="Fecha Realizado:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_fecha_real).grid(row=0, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Costo Real:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_costo_real).grid(row=1, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Observaciones:").grid(row=2, column=0, padx=5, pady=5, sticky='ne')
        text_obs = tk.Text(frame_principal, height=5, width=40, wrap='word')
        text_obs.grid(row=2, column=1, padx=5, pady=5, sticky='we')
        
        # Botones
        frame_botones = ttk.Frame(ventana)
        frame_botones.pack(fill='x', padx=10, pady=10)
        
        btn_guardar = ttk.Button(frame_botones, text="Guardar", command=lambda: self.actualizar_mantenimiento_db(
            mantenimiento_id,
            self.var_fecha_real.get(),
            self.var_costo_real.get(),
            text_obs.get("1.0", "end").strip(),
            ventana
        ))
        btn_guardar.pack(side='left', padx=5)
        
        btn_cancelar = ttk.Button(frame_botones, text="Cancelar", command=ventana.destroy)
        btn_cancelar.pack(side='left', padx=5)
    
    def guardar_mantenimiento(self, equipo_id, tipo, fecha_prog, fecha_real, descripcion, tecnico, estado, costo, observaciones, ventana):
        """Guarda un nuevo mantenimiento en la base de datos"""
        if not equipo_id:
            messagebox.showwarning("Advertencia", "Seleccione un equipo")
            return
        if not fecha_prog:
            messagebox.showwarning("Advertencia", "La fecha programada es obligatoria")
            return
            
        try:
            # Insertar mantenimiento
            self.c.execute("""INSERT INTO mantenimientos 
                              (equipo_id, tipo, fecha_programada, fecha_realizado, 
                               descripcion, tecnico, estado, costo, observaciones) 
                              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                          (equipo_id, tipo, fecha_prog, fecha_real, descripcion, tecnico, estado, costo, observaciones))
            self.conn.commit()
            
            # Actualizar fecha de último mantenimiento en el equipo si ya se realizó
            if fecha_real:
                self.c.execute("UPDATE equipos SET ultimo_mantenimiento = ? WHERE id = ?", 
                              (fecha_real, equipo_id))
                self.conn.commit()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Programó mantenimiento para equipo ID: {equipo_id}")
            
            messagebox.showinfo("Éxito", "Mantenimiento registrado correctamente")
            ventana.destroy()
            self.actualizar_mantenimientos()
            self.actualizar_equipos()  # Para actualizar la fecha de último mantenimiento
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo guardar el mantenimiento: {e}")
    
    def actualizar_mantenimiento_db(self, mantenimiento_id, fecha_real, costo_real, observaciones, ventana):
        """Actualiza un mantenimiento en la base de datos"""
        if not fecha_real:
            messagebox.showwarning("Advertencia", "La fecha de realización es obligatoria")
            return
            
        try:
            # Obtener equipo_id para actualizar la fecha de último mantenimiento
            self.c.execute("SELECT equipo_id FROM mantenimientos WHERE id = ?", (mantenimiento_id,))
            equipo_id = self.c.fetchone()[0]
            
            # Actualizar mantenimiento
            self.c.execute("""UPDATE mantenimientos SET 
                              fecha_realizado = ?, 
                              estado = 'Completado',
                              costo = ?,
                              observaciones = ?
                              WHERE id = ?""",
                          (fecha_real, costo_real, observaciones, mantenimiento_id))
            self.conn.commit()
            
            # Actualizar fecha de último mantenimiento en el equipo
            self.c.execute("UPDATE equipos SET ultimo_mantenimiento = ? WHERE id = ?", 
                          (fecha_real, equipo_id))
            self.conn.commit()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Registró mantenimiento ID: {mantenimiento_id}")
            
            messagebox.showinfo("Éxito", "Mantenimiento registrado correctamente")
            ventana.destroy()
            self.actualizar_mantenimientos()
            self.actualizar_equipos()
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo actualizar el mantenimiento: {e}")
    
    def exportar_mantenimientos(self):
        """Exporta los mantenimientos a un archivo Excel"""
        try:
            # Obtener todos los mantenimientos visibles
            items = self.tree_mantenimientos.get_children()
            data = [self.tree_mantenimientos.item(item, 'values') for item in items]
            
            if not data:
                messagebox.showwarning("Advertencia", "No hay datos para exportar")
                return
                
            # Crear DataFrame
            df = pd.DataFrame(data, columns=["ID", "Equipo", "Tipo", "Fecha Programada", "Fecha Realizado", "Técnico", "Estado"])
            
            # Preguntar dónde guardar
            filepath = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Guardar mantenimientos como"
            )
            
            if not filepath:
                return
                
            # Guardar en Excel
            df.to_excel(filepath, index=False)
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Exportó lista de mantenimientos a {filepath}")
            
            messagebox.showinfo("Éxito", f"Mantenimientos exportados a:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar: {e}")
    
    def inicializar_configuracion(self):
        """Configura la subpestaña de configuración del laboratorio"""
        frame_principal = ttk.Frame(self.frame_configuracion, padding=20)
        frame_principal.pack(fill='both', expand=True)
        
        ttk.Label(frame_principal, text="Configuración del Laboratorio", font=('Arial', 12, 'bold')).pack(pady=10)
        
        # Configuración de red
        frame_red = ttk.LabelFrame(frame_principal, text="Configuración de Red", padding=10)
        frame_red.pack(fill='x', pady=5)
        
        ttk.Label(frame_red, text="Dirección IP:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.entry_ip = ttk.Entry(frame_red)
        self.entry_ip.grid(row=0, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_red, text="Máscara de Red:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.entry_mascara = ttk.Entry(frame_red)
        self.entry_mascara.grid(row=1, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_red, text="Puerta de Enlace:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.entry_gateway = ttk.Entry(frame_red)
        self.entry_gateway.grid(row=2, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_red, text="DNS Primario:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        self.entry_dns1 = ttk.Entry(frame_red)
        self.entry_dns1.grid(row=3, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_red, text="DNS Secundario:").grid(row=4, column=0, padx=5, pady=5, sticky='e')
        self.entry_dns2 = ttk.Entry(frame_red)
        self.entry_dns2.grid(row=4, column=1, padx=5, pady=5, sticky='we')
        
        # Configuración general
        frame_general = ttk.LabelFrame(frame_principal, text="Configuración General", padding=10)
        frame_general.pack(fill='x', pady=5)
        
        ttk.Label(frame_general, text="Nombre del Laboratorio:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.entry_nombre_lab = ttk.Entry(frame_general)
        self.entry_nombre_lab.grid(row=0, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_general, text="Ubicación:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.entry_ubicacion_lab = ttk.Entry(frame_general)
        self.entry_ubicacion_lab.grid(row=1, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_general, text="Responsable:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.entry_responsable = ttk.Entry(frame_general)
        self.entry_responsable.grid(row=2, column=1, padx=5, pady=5, sticky='we')
        
        # Botones
        frame_botones = ttk.Frame(frame_principal)
        frame_botones.pack(fill='x', pady=10)
        
        btn_guardar = ttk.Button(frame_botones, text="Guardar Configuración", command=self.guardar_configuracion)
        btn_guardar.pack(side='left', padx=5)
        
        btn_cargar = ttk.Button(frame_botones, text="Cargar Configuración", command=self.cargar_configuracion)
        btn_cargar.pack(side='left', padx=5)
        
        # Cargar configuración actual
        self.cargar_configuracion()
    
    def cargar_configuracion(self):
        """Carga la configuración actual del laboratorio"""
        try:
            # En una implementación real, esto cargaría desde una tabla de configuración
            # Aquí solo mostramos valores de ejemplo
            self.entry_ip.delete(0, 'end')
            self.entry_ip.insert(0, "192.168.1.100")
            
            self.entry_mascara.delete(0, 'end')
            self.entry_mascara.insert(0, "255.255.255.0")
            
            self.entry_gateway.delete(0, 'end')
            self.entry_gateway.insert(0, "192.168.1.1")
            
            self.entry_dns1.delete(0, 'end')
            self.entry_dns1.insert(0, "8.8.8.8")
            
            self.entry_dns2.delete(0, 'end')
            self.entry_dns2.insert(0, "8.8.4.4")
            
            self.entry_nombre_lab.delete(0, 'end')
            self.entry_nombre_lab.insert(0, "Laboratorio de Computación y Redes")
            
            self.entry_ubicacion_lab.delete(0, 'end')
            self.entry_ubicacion_lab.insert(0, "Edificio A, Piso 2, Sala 205")
            
            self.entry_responsable.delete(0, 'end')
            self.entry_responsable.insert(0, "Ing. Juan Pérez")
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar la configuración: {e}")
    
    def guardar_configuracion(self):
        """Guarda la configuración del laboratorio"""
        try:
            # En una implementación real, esto guardaría en una tabla de configuración
            ip = self.entry_ip.get()
            mascara = self.entry_mascara.get()
            gateway = self.entry_gateway.get()
            dns1 = self.entry_dns1.get()
            dns2 = self.entry_dns2.get()
            nombre = self.entry_nombre_lab.get()
            ubicacion = self.entry_ubicacion_lab.get()
            responsable = self.entry_responsable.get()
            
            # Validaciones básicas
            if not nombre or not responsable:
                messagebox.showwarning("Advertencia", "Nombre del laboratorio y responsable son obligatorios")
                return
                
            # Registrar en el historial de accesos
            self.registrar_acceso("Actualizó configuración del laboratorio")
            
            messagebox.showinfo("Éxito", "Configuración guardada correctamente")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la configuración: {e}")
    
    def crear_estadisticas_laboratorio(self, frame):
        """Crea gráficos estadísticos del laboratorio"""
        try:
            # Obtener datos para gráficos
            self.c.execute("SELECT estado, COUNT(*) FROM equipos GROUP BY estado")
            estados_equipos = self.c.fetchall()
            
            self.c.execute("SELECT tipo, COUNT(*) FROM reportes GROUP BY tipo")
            tipos_reportes = self.c.fetchall()
            
            self.c.execute("SELECT strftime('%Y-%m', fecha) as mes, COUNT(*) FROM reportes GROUP BY mes ORDER BY mes")
            reportes_por_mes = self.c.fetchall()
            
            self.c.execute("SELECT estado, COUNT(*) FROM mantenimientos GROUP BY estado")
            estados_mantenimientos = self.c.fetchall()
            
            # Crear figura con subplots
            fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 8))
            fig.suptitle('Estadísticas Generales del Laboratorio', fontsize=14)
            
            # Gráfico 1: Estado de los equipos
            if estados_equipos:
                estados = [est[0] for est in estados_equipos]
                cantidades = [est[1] for est in estados_equipos]
                
                ax1.bar(estados, cantidades, color=['#4CAF50', '#FFC107', '#F44336', '#9E9E9E'])
                ax1.set_title('Estado de los Equipos')
                ax1.set_ylabel('Cantidad')
                ax1.tick_params(axis='x', rotation=45)
            else:
                ax1.text(0.5, 0.5, 'No hay datos\nde equipos', ha='center', va='center')
                ax1.set_title('Estado de los Equipos')
            
            # Gráfico 2: Tipos de reportes
            if tipos_reportes:
                tipos = [tipo[0] for tipo in tipos_reportes]
                cantidades = [tipo[1] for tipo in tipos_reportes]
                
                ax2.pie(cantidades, labels=tipos, autopct='%1.1f%%', startangle=90)
                ax2.set_title('Distribución de Reportes')
            else:
                ax2.text(0.5, 0.5, 'No hay datos\nde reportes', ha='center', va='center')
                ax2.set_title('Distribución de Reportes')
            
            # Gráfico 3: Reportes por mes
            if reportes_por_mes and len(reportes_por_mes) > 1:
                meses = [mes[0] for mes in reportes_por_mes]
                cantidades = [mes[1] for mes in reportes_por_mes]
                
                ax3.plot(meses, cantidades, marker='o')
                ax3.set_title('Reportes por Mes')
                ax3.set_ylabel('Cantidad')
                ax3.tick_params(axis='x', rotation=45)
            else:
                ax3.text(0.5, 0.5, 'No hay suficientes datos\npor mes', ha='center', va='center')
                ax3.set_title('Reportes por Mes')
            
            # Gráfico 4: Estado de mantenimientos
            if estados_mantenimientos:
                estados = [est[0] for est in estados_mantenimientos]
                cantidades = [est[1] for est in estados_mantenimientos]
                
                ax4.barh(estados, cantidades, color=['#FFC107', '#4CAF50', '#F44336', '#9E9E9E'])
                ax4.set_title('Estado de Mantenimientos')
                ax4.set_xlabel('Cantidad')
            else:
                ax4.text(0.5, 0.5, 'No hay datos\nde mantenimientos', ha='center', va='center')
                ax4.set_title('Estado de Mantenimientos')
            
            fig.tight_layout()
            
            # Integrar gráfico en Tkinter
            canvas = FigureCanvasTkAgg(fig, master=frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            
            # Guardar referencia para evitar garbage collection
            self.canvas_estadisticas = canvas
            
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudieron generar los gráficos: {e}")
    
    # ------------------------- Pestaña de Usuarios -------------------------
    def inicializar_usuarios(self):
        """Configura los componentes de la pestaña de usuarios"""
        # Frame para controles
        frame_controles = ttk.Frame(self.frame_usuarios)
        frame_controles.pack(fill='x', padx=10, pady=5)
        
        # Botones de acción
        btn_agregar = ttk.Button(frame_controles, text="Agregar Usuario", command=self.agregar_usuario)
        btn_agregar.pack(side='left', padx=5)
        
        btn_editar = ttk.Button(frame_controles, text="Editar", command=self.editar_usuario)
        btn_editar.pack(side='left', padx=5)
        
        btn_eliminar = ttk.Button(frame_controles, text="Eliminar", command=self.eliminar_usuario)
        btn_eliminar.pack(side='left', padx=5)
        
        btn_actualizar = ttk.Button(frame_controles, text="Actualizar", command=self.actualizar_usuarios)
        btn_actualizar.pack(side='left', padx=5)
        
        # Frame para la tabla de usuarios
        frame_tabla = ttk.Frame(self.frame_usuarios)
        frame_tabla.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Tabla de usuarios
        columns = ("ID", "Nombre", "Apellido", "Email", "Rol", "Usuario", "Estado")
        self.tree_usuarios = ttk.Treeview(frame_tabla, columns=columns, show='headings', selectmode='browse')
        
        for col in columns:
            self.tree_usuarios.heading(col, text=col)
            self.tree_usuarios.column(col, width=100, anchor='center')
        
        self.tree_usuarios.column("Nombre", width=120)
        self.tree_usuarios.column("Apellido", width=120)
        self.tree_usuarios.column("Email", width=150)
        self.tree_usuarios.pack(fill='both', expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree_usuarios.yview)
        scrollbar.pack(side='right', fill='y')
        self.tree_usuarios.configure(yscrollcommand=scrollbar.set)
        
        # Frame para historial de accesos
        frame_historial = ttk.LabelFrame(self.frame_usuarios, text="Historial de Accesos", padding=10)
        frame_historial.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Tabla de historial
        columns_historial = ("ID", "Usuario", "Fecha/Hora", "Acción", "Detalles")
        self.tree_historial = ttk.Treeview(frame_historial, columns=columns_historial, show='headings', selectmode='browse')
        
        for col in columns_historial:
            self.tree_historial.heading(col, text=col)
            self.tree_historial.column(col, width=100, anchor='center')
        
        self.tree_historial.column("Usuario", width=120)
        self.tree_historial.column("Fecha/Hora", width=120)
        self.tree_historial.column("Acción", width=120)
        self.tree_historial.column("Detalles", width=200)
        self.tree_historial.pack(fill='both', expand=True)
        
        # Scrollbar
        scrollbar_hist = ttk.Scrollbar(frame_historial, orient="vertical", command=self.tree_historial.yview)
        scrollbar_hist.pack(side='right', fill='y')
        self.tree_historial.configure(yscrollcommand=scrollbar_hist.set)
        
        # Cargar datos iniciales
        self.actualizar_usuarios()
        self.actualizar_historial_accesos()
    
    def actualizar_usuarios(self):
        """Actualiza la tabla de usuarios con los datos de la base de datos"""
        try:
            # Limpiar tabla
            for item in self.tree_usuarios.get_children():
                self.tree_usuarios.delete(item)
            
            # Obtener datos
            self.c.execute("SELECT id, nombre, apellido, email, rol, usuario, estado FROM usuarios ORDER BY nombre")
            usuarios = self.c.fetchall()
            
            # Llenar tabla
            for usuario in usuarios:
                self.tree_usuarios.insert('', 'end', values=usuario)
                
            # Resaltar usuarios inactivos
            for item in self.tree_usuarios.get_children():
                valores = self.tree_usuarios.item(item, 'values')
                if valores[6] == "Inactivo":
                    self.tree_usuarios.tag_configure('inactivo', foreground='gray')
                    self.tree_usuarios.item(item, tags=('inactivo',))
                    
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo cargar los usuarios: {e}")
    
    def agregar_usuario(self):
        """Abre ventana para agregar nuevo usuario"""
        ventana = tk.Toplevel(self.root)
        ventana.title("Agregar Usuario")
        ventana.geometry("400x400")
        ventana.grab_set()
        
        # Variables
        self.var_nombre = tk.StringVar()
        self.var_apellido = tk.StringVar()
        self.var_email = tk.StringVar()
        self.var_rol = tk.StringVar(value="Usuario")
        self.var_usuario = tk.StringVar()
        self.var_contrasena = tk.StringVar()
        self.var_estado = tk.StringVar(value="Activo")
        
        # Frame principal
        frame_principal = ttk.Frame(ventana, padding=10)
        frame_principal.pack(fill='both', expand=True)
        
        # Campos del formulario
        ttk.Label(frame_principal, text="Nombre:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_nombre).grid(row=0, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Apellido:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_apellido).grid(row=1, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Email:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_email).grid(row=2, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Rol:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        combo_rol = ttk.Combobox(frame_principal, textvariable=self.var_rol, 
                                values=["Administrador", "Técnico", "Usuario"], state='readonly')
        combo_rol.grid(row=3, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Usuario:").grid(row=4, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_usuario).grid(row=4, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Contraseña:").grid(row=5, column=0, padx=5, pady=5, sticky='e')
        ttk.Entry(frame_principal, textvariable=self.var_contrasena, show="*").grid(row=5, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(frame_principal, text="Estado:").grid(row=6, column=0, padx=5, pady=5, sticky='e')
        combo_estado = ttk.Combobox(frame_principal, textvariable=self.var_estado, 
                                   values=["Activo", "Inactivo"], state='readonly')
        combo_estado.grid(row=6, column=1, padx=5, pady=5, sticky='we')
        
        # Botones
        frame_botones = ttk.Frame(ventana)
        frame_botones.pack(fill='x', padx=10, pady=10)
        
        btn_guardar = ttk.Button(frame_botones, text="Guardar", command=lambda: self.guardar_usuario(
            self.var_nombre.get(),
            self.var_apellido.get(),
            self.var_email.get(),
            self.var_rol.get(),
            self.var_usuario.get(),
            self.var_contrasena.get(),
            self.var_estado.get(),
            ventana
        ))
        btn_guardar.pack(side='left', padx=5)
        
        btn_cancelar = ttk.Button(frame_botones, text="Cancelar", command=ventana.destroy)
        btn_cancelar.pack(side='left', padx=5)
    
    def guardar_usuario(self, nombre, apellido, email, rol, usuario, contrasena, estado, ventana):
        """Guarda un nuevo usuario en la base de datos"""
        if not nombre or not usuario or not contrasena:
            messagebox.showwarning("Advertencia", "Nombre, usuario y contraseña son obligatorios")
            return
            
        try:
            fecha_registro = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            self.c.execute("""INSERT INTO usuarios 
                              (nombre, apellido, email, rol, usuario, contrasena, estado, fecha_registro) 
                              VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                          (nombre, apellido, email, rol, usuario, contrasena, estado, fecha_registro))
            self.conn.commit()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Agregó nuevo usuario: {usuario}")
            
            messagebox.showinfo("Éxito", "Usuario agregado correctamente")
            ventana.destroy()
            self.actualizar_usuarios()
        except sqlite3.Error as e:
            if "UNIQUE" in str(e):
                if "email" in str(e):
                    messagebox.showerror("Error", "El email ya está registrado")
                else:
                    messagebox.showerror("Error", "El nombre de usuario ya existe")
            else:
                messagebox.showerror("Error", f"No se pudo guardar el usuario: {e}")
    
    def editar_usuario(self):
        """Abre ventana para editar usuario seleccionado"""
        seleccion = self.tree_usuarios.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione un usuario para editar")
            return
            
        usuario_id = self.tree_usuarios.item(seleccion[0], 'values')[0]
        
        try:
            self.c.execute("SELECT nombre, apellido, email, rol, usuario, estado FROM usuarios WHERE id = ?", (usuario_id,))
            usuario = self.c.fetchone()
            
            ventana = tk.Toplevel(self.root)
            ventana.title("Editar Usuario")
            ventana.geometry("400x400")
            ventana.grab_set()
            
            # Variables
            self.var_nombre = tk.StringVar(value=usuario[0])
            self.var_apellido = tk.StringVar(value=usuario[1])
            self.var_email = tk.StringVar(value=usuario[2])
            self.var_rol = tk.StringVar(value=usuario[3])
            self.var_usuario = tk.StringVar(value=usuario[4])
            self.var_contrasena = tk.StringVar()
            self.var_estado = tk.StringVar(value=usuario[5])
            
            # Frame principal
            frame_principal = ttk.Frame(ventana, padding=10)
            frame_principal.pack(fill='both', expand=True)
            
            # Campos del formulario
            ttk.Label(frame_principal, text="Nombre:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
            ttk.Entry(frame_principal, textvariable=self.var_nombre).grid(row=0, column=1, padx=5, pady=5, sticky='we')
            
            ttk.Label(frame_principal, text="Apellido:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
            ttk.Entry(frame_principal, textvariable=self.var_apellido).grid(row=1, column=1, padx=5, pady=5, sticky='we')
            
            ttk.Label(frame_principal, text="Email:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
            ttk.Entry(frame_principal, textvariable=self.var_email).grid(row=2, column=1, padx=5, pady=5, sticky='we')
            
            ttk.Label(frame_principal, text="Rol:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
            combo_rol = ttk.Combobox(frame_principal, textvariable=self.var_rol, 
                                    values=["Administrador", "Técnico", "Usuario"], state='readonly')
            combo_rol.grid(row=3, column=1, padx=5, pady=5, sticky='we')
            
            ttk.Label(frame_principal, text="Usuario:").grid(row=4, column=0, padx=5, pady=5, sticky='e')
            ttk.Entry(frame_principal, textvariable=self.var_usuario, state='readonly').grid(row=4, column=1, padx=5, pady=5, sticky='we')
            
            ttk.Label(frame_principal, text="Nueva Contraseña:").grid(row=5, column=0, padx=5, pady=5, sticky='e')
            ttk.Entry(frame_principal, textvariable=self.var_contrasena, show="*").grid(row=5, column=1, padx=5, pady=5, sticky='we')
            
            ttk.Label(frame_principal, text="Estado:").grid(row=6, column=0, padx=5, pady=5, sticky='e')
            combo_estado = ttk.Combobox(frame_principal, textvariable=self.var_estado, 
                                       values=["Activo", "Inactivo"], state='readonly')
            combo_estado.grid(row=6, column=1, padx=5, pady=5, sticky='we')
            
            # Botones
            frame_botones = ttk.Frame(ventana)
            frame_botones.pack(fill='x', padx=10, pady=10)
            
            btn_guardar = ttk.Button(frame_botones, text="Guardar", 
                                   command=lambda: self.actualizar_usuario(
                                       usuario_id,
                                       self.var_nombre.get(),
                                       self.var_apellido.get(),
                                       self.var_email.get(),
                                       self.var_rol.get(),
                                       self.var_contrasena.get(),
                                       self.var_estado.get(),
                                       ventana
                                   ))
            btn_guardar.pack(side='left', padx=5)
            
            btn_cancelar = ttk.Button(frame_botones, text="Cancelar", command=ventana.destroy)
            btn_cancelar.pack(side='left', padx=5)
            
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo cargar el usuario: {e}")
    
    def actualizar_usuario(self, usuario_id, nombre, apellido, email, rol, contrasena, estado, ventana):
        """Actualiza un usuario en la base de datos"""
        if not nombre:
            messagebox.showwarning("Advertencia", "El nombre es obligatorio")
            return
            
        try:
            if contrasena:
                # Actualizar con nueva contraseña
                self.c.execute("""UPDATE usuarios SET 
                                  nombre = ?, apellido = ?, email = ?, rol = ?, 
                                  contrasena = ?, estado = ? 
                                  WHERE id = ?""",
                              (nombre, apellido, email, rol, contrasena, estado, usuario_id))
            else:
                # Actualizar sin cambiar contraseña
                self.c.execute("""UPDATE usuarios SET 
                                  nombre = ?, apellido = ?, email = ?, rol = ?, 
                                  estado = ? 
                                  WHERE id = ?""",
                              (nombre, apellido, email, rol, estado, usuario_id))
                
            self.conn.commit()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Actualizó usuario ID: {usuario_id}")
            
            messagebox.showinfo("Éxito", "Usuario actualizado")
            ventana.destroy()
            self.actualizar_usuarios()
        except sqlite3.Error as e:
            if "UNIQUE" in str(e):
                messagebox.showerror("Error", "El email ya está registrado")
            else:
                messagebox.showerror("Error", f"No se pudo actualizar el usuario: {e}")
    
    def eliminar_usuario(self):
        """Elimina el usuario seleccionado"""
        seleccion = self.tree_usuarios.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione un usuario para eliminar")
            return
            
        usuario_id = self.tree_usuarios.item(seleccion[0], 'values')[0]
        usuario_nombre = self.tree_usuarios.item(seleccion[0], 'values')[1]
        
        # Verificar si es el usuario actual
        if usuario_id == 1:  # Suponiendo que el admin tiene ID 1
            messagebox.showerror("Error", "No se puede eliminar al usuario administrador principal")
            return
            
        if messagebox.askyesno("Confirmar", f"¿Eliminar al usuario '{usuario_nombre}'?"):
            try:
                self.c.execute("DELETE FROM usuarios WHERE id = ?", (usuario_id,))
                self.conn.commit()
                
                # Registrar en el historial de accesos
                self.registrar_acceso(f"Eliminó usuario: {usuario_nombre}")
                
                messagebox.showinfo("Éxito", "Usuario eliminado")
                self.actualizar_usuarios()
            except sqlite3.Error as e:
                messagebox.showerror("Error", f"No se pudo eliminar el usuario: {e}")
    
    def actualizar_historial_accesos(self):
        """Actualiza la tabla de historial de accesos"""
        try:
            # Limpiar tabla
            for item in self.tree_historial.get_children():
                self.tree_historial.delete(item)
            
            # Obtener datos
            self.c.execute("""SELECT a.id, u.nombre, a.fecha_hora, a.accion, a.detalles 
                              FROM accesos a 
                              JOIN usuarios u ON a.usuario_id = u.id 
                              ORDER BY a.fecha_hora DESC 
                              LIMIT 100""")
            accesos = self.c.fetchall()
            
            # Llenar tabla
            for acceso in accesos:
                self.tree_historial.insert('', 'end', values=acceso)
                
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo cargar el historial: {e}")
    
    def registrar_acceso(self, accion, detalles=None):
        """Registra una acción en el historial de accesos"""
        try:
            usuario_id = 1  # En una aplicación real, obtendríamos el ID del usuario actual
            fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            if detalles is None:
                detalles = accion
            
            self.c.execute("""INSERT INTO accesos 
                              (usuario_id, fecha_hora, accion, detalles) 
                              VALUES (?, ?, ?, ?)""",
                          (usuario_id, fecha_hora, accion, detalles))
            self.conn.commit()
            
            # Actualizar el historial si estamos en la pestaña de usuarios
            if self.notebook.index(self.notebook.select()) == 3:  # Índice de pestaña de usuarios
                self.actualizar_historial_accesos()
        except sqlite3.Error as e:
            print(f"Error al registrar acceso: {e}")  # No mostramos mensaje para no molestar al usuario
    
    # ------------------------- Funciones generales -------------------------
    def crear_respaldo(self):
        """Crea una copia de seguridad de la base de datos"""
        try:
            # Preguntar dónde guardar el respaldo
            filepath = filedialog.asksaveasfilename(
                defaultextension=".db",
                filetypes=[("Database files", "*.db"), ("All files", "*.*")],
                title="Guardar copia de seguridad como"
            )
            
            if not filepath:
                return
                
            # Cerrar la conexión actual para poder copiar el archivo
            if hasattr(self, 'conn'):
                self.conn.close()
            
            # Copiar el archivo de la base de datos
            import shutil
            shutil.copy2('laboratorio.db', filepath)
            
            # Reconectar a la base de datos
            self.conexion_db()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Creó copia de seguridad en {filepath}")
            
            messagebox.showinfo("Éxito", f"Copia de seguridad creada en:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear la copia de seguridad: {e}")
            # Intentar reconectar si falló
            self.conexion_db()
    
    def restaurar_respaldo(self):
        """Restaura la base de datos desde una copia de seguridad"""
        try:
            # Preguntar por el archivo de respaldo
            filepath = filedialog.askopenfilename(
                filetypes=[("Database files", "*.db"), ("All files", "*.*")],
                title="Seleccionar archivo de respaldo"
            )
            
            if not filepath:
                return
                
            # Confirmar con el usuario
            if not messagebox.askyesno("Confirmar", "¿Restaurar desde esta copia de seguridad?\nTodos los datos actuales serán reemplazados."):
                return
                
            # Cerrar la conexión actual
            if hasattr(self, 'conn'):
                self.conn.close()
            
            # Copiar el archivo de respaldo
            import shutil
            shutil.copy2(filepath, 'laboratorio.db')
            
            # Reconectar a la base de datos
            self.conexion_db()
            
            # Registrar en el historial de accesos
            self.registrar_acceso(f"Restauró base de datos desde {filepath}")
            
            # Recargar todos los datos
            self.cargar_datos_iniciales()
            
            messagebox.showinfo("Éxito", "Base de datos restaurada correctamente")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo restaurar la copia de seguridad: {e}")
            # Intentar reconectar si falló
            self.conexion_db()
    
    def mostrar_documentacion(self):
        """Muestra la documentación del sistema"""
        ventana = tk.Toplevel(self.root)
        ventana.title("Documentación del Sistema")
        ventana.geometry("800x600")
        
        # Frame principal con scrollbar
        frame_principal = ttk.Frame(ventana)
        frame_principal.pack(fill='both', expand=True)
        
        canvas = tk.Canvas(frame_principal)
        scrollbar = ttk.Scrollbar(frame_principal, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Contenido de la documentación
        ttk.Label(scrollable_frame, text="Documentación del Sistema", font=('Arial', 14, 'bold')).pack(pady=10)
        
        secciones = [
            ("1. Reportes de Equipos", """
            - Crear nuevos reportes de problemas técnicos
            - Seguir el estado de los reportes (Abierto, En Progreso, Resuelto)
            - Asignar prioridades a los reportes
            - Registrar soluciones aplicadas
            - Exportar reportes a Excel
            """),
            
            ("2. Gestión de Inventario", """
            - Registrar componentes de hardware y software
            - Controlar niveles de stock y mínimos requeridos
            - Registrar proveedores y ubicaciones
            - Visualizar estadísticas de inventario
            - Exportar inventario a Excel
            """),
            
            ("3. Gestión del Laboratorio", """
            - Registrar equipos con sus características técnicas
            - Programar y realizar reservas de equipos
            - Planificar mantenimientos preventivos y correctivos
            - Visualizar estadísticas generales del laboratorio
            - Configurar parámetros del laboratorio
            """),
            
            ("4. Usuarios y Accesos", """
            - Administrar usuarios del sistema
            - Asignar roles (Administrador, Técnico, Usuario)
            - Ver historial de actividades realizadas
            - Activar/desactivar usuarios
            """),
            
            ("5. Copias de Seguridad", """
            - Crear copias de seguridad de la base de datos
            - Restaurar desde una copia de seguridad
            - Exportar datos a formatos externos
            """)
        ]
        
        for titulo, contenido in secciones:
            frame_seccion = ttk.Frame(scrollable_frame, borderwidth=1, relief='solid', padding=10)
            frame_seccion.pack(fill='x', padx=10, pady=5)
            
            ttk.Label(frame_seccion, text=titulo, font=('Arial', 12, 'bold')).pack(anchor='w')
            ttk.Label(frame_seccion, text=contenido.strip(), justify='left').pack(anchor='w')
    
    def mostrar_acerca_de(self):
        """Muestra la información 'Acerca de' del sistema"""
        ventana = tk.Toplevel(self.root)
        ventana.title("Acerca de")
        ventana.geometry("400x300")
        
        frame_principal = ttk.Frame(ventana, padding=20)
        frame_principal.pack(fill='both', expand=True)
        
        ttk.Label(frame_principal, text="Sistema de Gestión de Laboratorio", font=('Arial', 14, 'bold')).pack(pady=10)
        ttk.Label(frame_principal, text="Versión 1.0").pack(pady=5)
        ttk.Label(frame_principal, text="Desarrollado por:").pack(pady=5)
        ttk.Label(frame_principal, text="Equipo de Desarrollo TI", font=('Arial', 10, 'bold')).pack(pady=5)
        ttk.Label(frame_principal, text="© 2023 Todos los derechos reservados").pack(pady=20)
        
        btn_cerrar = ttk.Button(frame_principal, text="Cerrar", command=ventana.destroy)
        btn_cerrar.pack(pady=10)
    
    def cerrar_aplicacion(self):
        """Cierra la aplicación de manera segura"""
        if messagebox.askokcancel("Salir", "¿Está seguro que desea salir de la aplicación?"):
            if hasattr(self, 'conn'):
                self.conn.close()
            self.root.destroy()
    
    def __del__(self):
        """Destructor para cerrar la conexión a la base de datos"""
        if hasattr(self, 'conn'):
            self.conn.close()

# Función principal
if __name__ == "__main__":
    root = tk.Tk()
    app = SistemaGestionLaboratorio(root)
    root.mainloop()