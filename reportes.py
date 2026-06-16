# ==============================================================================
# INSTRUCCIONES DE INSTALACIÓN Y EMPAQUETADO (LEER ANTES DE EJECUTAR)
# ==============================================================================
# 1. Instalación de dependencias:
#    Abra su terminal y ejecute:
#    pip install openpyxl pandas
#
# 2. Ejecución en desarrollo:
#    python app.py
#
# 3. Instrucciones para empaquetar como ejecutable (.exe en Windows):
#    Instale PyInstaller: pip install pyinstaller
#    Ejecute el siguiente comando en la terminal:
#    pyinstaller --noconsole --onefile --name="Galeria_Inventario" app.py
# ==============================================================================

# ============================================
# SECCIÓN 1: IMPORTACIONES Y CONFIGURACIÓN
# ============================================
import os
import datetime
from datetime import timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Configuración global de rutas por defecto si no se seleccionan
DB_MAESTRO_DEFAULT = "inventario_maestro.xlsx"
DB_REPORTES_DEFAULT = f"reportes_ventas_{datetime.datetime.now().year}.xlsx"

# ============================================
# SECCIÓN 2: ESTILOS Y DISEÑO DE INTERFAZ
# ============================================
# Paleta de colores elegante para una Galería de Arte
COLOR_PRIMARY = "#2C3E50"    # Azul Medianoche
COLOR_SECONDARY = "#16A085"  # Verde Azulado
COLOR_BG = "#F8F9FA"         # Blanco Grisáceo
COLOR_TEXT = "#34495E"       # Gris Oscuro
COLOR_ACCENT = "#E74C3C"     # Coral para alertas/bajas

FONT_TITLE = ("Segoe UI", 14, "bold")
FONT_SUBTITLE = ("Segoe UI", 11, "bold")
FONT_BODY = ("Segoe UI", 10)

def aplicar_estilos_modernos():
    """Configura el catálogo de estilos ttk para acoplarse a la paleta visual."""
    style = ttk.Style()
    style.theme_use('clam')
    
    style.configure(".", background=COLOR_BG, foreground=COLOR_TEXT, font=FONT_BODY)
    style.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT)
    style.configure("TFrame", background=COLOR_BG)
    
    # Notebook (Pestañas)
    style.configure("TNotebook", background=COLOR_BG, borderwidth=0)
    style.configure("TNotebook.Tab", background="#BDC3C7", foreground=COLOR_TEXT, padding=[15, 5], font=FONT_BODY)
    style.map("TNotebook.Tab", background=[("selected", COLOR_PRIMARY)], foreground=[("selected", "white")])
    
    # Botones
    style.configure("TButton", background=COLOR_PRIMARY, foreground="white", borderwidth=0, padding=6, focuscolor=COLOR_PRIMARY)
    style.map("TButton", background=[("active", COLOR_SECONDARY)])
    style.configure("Accent.TButton", background=COLOR_SECONDARY, foreground="white")
    style.map("Accent.TButton", background=[("active", COLOR_PRIMARY)])
    
    # Entradas de texto
    style.configure("TEntry", fieldbackground="white", borderwidth=1)
    style.configure("TCombobox", fieldbackground="white", borderwidth=1)

# ============================================
# SECCIÓN 3: LÓGICA DE NEGOCIO (EXCEL)
# ============================================

def inicializar_excel_maestro(ruta):
    """
    Crea un archivo maestro de inventario básico si este no existe.
    :param ruta: str, camino del archivo de Excel.
    """
    if not os.path.exists(ruta):
        df = pd.DataFrame(columns=["Código", "Nombre", "Departamento", "Precio Venta", "Cantidad"])
        df.to_excel(ruta, index=False)

def buscar_producto_maestro(ruta, codigo):
    """
    Busca un producto en el archivo maestro por su código.
    :param ruta: str, ruta del archivo maestro.
    :param codigo: str, código del producto.
    :return: dict o None si no se encuentra.
    """
    if not os.path.exists(ruta):
        return None
    try:
        df = pd.read_excel(ruta, dtype={"Código": str})
        res = df[df["Código"] == str(codigo).strip()]
        if not res.empty:
            return res.iloc[0].to_dict()
    except Exception as e:
        messagebox.showerror("Error de lectura", f"No se pudo leer el maestro:\n{e}")
    return None

def guardar_o_actualizar_producto(ruta, item_dict):
    """
    Inserta o actualiza un producto en el maestro manteniendo orden por Departamento y Código.
    :param ruta: str, ruta del archivo maestro.
    :param item_dict: dict, datos del producto.
    """
    inicializar_excel_maestro(ruta)
    try:
        df = pd.read_excel(ruta, dtype={"Código": str})
        # Limpiar datos
        item_dict["Código"] = str(item_dict["Código"]).strip()
        
        # Eliminar si ya existe para reinsertar ordenado
        df = df[df["Código"] != item_dict["Código"]]
        
        # Convertir el dict en DataFrame seguro
        new_row = pd.DataFrame([item_dict])
        df = pd.concat([df, new_row], ignore_index=True)
        
        # Ordenar por Departamento y luego por Código
        df.sort_values(by=["Departamento", "Código"], inplace=True)
        df.to_excel(ruta, index=False)
        return True
    except Exception as e:
        messagebox.showerror("Error de escritura", f"No se pudo guardar en maestro:\n{e}")
        return False

def modificar_stock_maestro(ruta, codigo, cantidad_cambio):
    """
    Modifica la cantidad existente de un producto en el maestro (Suma algebráica).
    :param ruta: str, ruta del archivo maestro.
    :param codigo: str, código del ítem.
    :param cantidad_cambio: int, cantidad a sumar o restar.
    """
    try:
        df = pd.read_excel(ruta, dtype={"Código": str})
        codigo_str = str(codigo).strip()
        if codigo_str in df["Código"].values:
            idx = df[df["Código"] == codigo_str].index[0]
            nueva_cant = df.at[idx, "Cantidad"] + cantidad_cambio
            if nueva_cant < 0:
                return False, "La cantidad en inventario no puede ser negativa."
            df.at[idx, "Cantidad"] = nueva_cant
            df.to_excel(ruta, index=False)
            return True, nueva_cant
        return False, "Código no encontrado."
    except Exception as e:
        return False, str(e)

def obtener_semana_y_mes_asignado(fecha):
    """
    Calcula el rango de la semana (Lunes a Domingo) y determina a qué mes 
    pertenece basándose en la regla de la mayoría de días.
    :param fecha: datetime.date
    :return: (str nombre_mes, date lunes, date domingo)
    """
    lunes = fecha - timedelta(days=fecha.weekday())
    domingo = lunes + timedelta(days=6)
    
    # Contar cuántos días de esta semana caen en el mes del lunes y cuántos en el del domingo
    mes_lunes = lunes.month
    mes_domingo = domingo.month
    
    if mes_lunes == mes_domingo:
        mes_asignado = lunes.strftime("%B").capitalize()
    else:
        # Contar días en el mes de inicio
        dias_mes_lunes = (datetime.date(lunes.year, mes_lunes, 1) + timedelta(days=32)).replace(day=1) - lunes
        dias_mes_lunes = dias_mes_lunes.days if dias_mes_lunes.days < 7 else 7
        dias_mes_domingo = 7 - dias_mes_lunes
        
        if dias_mes_lunes >= dias_mes_domingo:
            mes_asignado = lunes.strftime("%B").capitalize()
        else:
            mes_asignado = domingo.strftime("%B").capitalize()
            
    # Traducción rápida de meses al español
    meses_es = {
        "January": "Enero", "February": "Febrero", "March": "Marzo", "April": "Abril",
        "May": "Mayo", "June": "Junio", "July": "Julio", "August": "Agosto",
        "September": "Septiembre", "October": "Octubre", "November": "Noviembre", "December": "Diciembre"
    }
    return meses_es.get(mes_asignado, mes_asignado), lunes, domingo

def registrar_venta_en_reporte(ruta_reporte, venta_data):
    """
    Guarda una transacción en el archivo de reportes mensuales segmentado por hojas de mes.
    :param ruta_reporte: str, ruta de destino.
    :param venta_data: dict, datos de la venta.
    """
    fecha_v = datetime.datetime.strptime(venta_data["Fecha"], "%Y-%m-%d").date()
    mes_hoja, lunes, domingo = obtener_semana_y_mes_asignado(fecha_v)
    
    # Crear archivo si no existe
    if not os.path.exists(ruta_reporte):
        wb = openpyxl.Workbook()
        wb.save(ruta_reporte)
        
    try:
        wb = openpyxl.load_workbook(ruta_reporte)
        if mes_hoja not in wb.sheetnames:
            ws = wb.create_sheet(title=mes_hoja)
            ws.append(["Fecha", "Semana (Rango)", "Código/Nombre", "Tipo", "Cantidad", "Efectivo", "Tarjeta", "Total Descuento/Modif", "Total Dinero"])
        else:
            ws = wb[mes_hoja]
            
        semana_str = f"{lunes.strftime('%d/%m')} al {domingo.strftime('%d/%m')}"
        
        ws.append([
            venta_data["Fecha"],
            semana_str,
            venta_data["Identificador"],
            venta_data["Tipo"],
            venta_data["Cantidad"],
            venta_data["Efectivo"],
            venta_data["Tarjeta"],
            venta_data["Descuento_Aplicado"],
            venta_data["Total"]
        ])
        
        # Eliminar hoja default si se creó vacía al inicio
        if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
            wb.remove(wb["Sheet"])
            
        wb.save(ruta_reporte)
        return True
    except Exception as e:
        messagebox.showerror("Error al reportar venta", f"Asegúrese de que el archivo de reportes no esté abierto.\n{e}")
        return False

# ============================================
# SECCIÓN 4: CONTROLADORES DE EVENTOS Y UI
# ============================================

class AppGaleria(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Gestión de Galería de Arte v1.0")
        self.geometry("900x650")
        self.configure(bg=COLOR_BG)
        
        self.ruta_maestro = tk.StringVar(value=DB_MAESTRO_DEFAULT)
        self.ruta_reportes = tk.StringVar(value=DB_REPORTES_DEFAULT)
        
        aplicar_estilos_modernos()
        self.crear_interfaz_seleccion_archivo()
        self.crear_cuerpo_principal()
        
        # Inicialización de bases de datos locales por defecto
        inicializar_excel_maestro(self.ruta_maestro.get())

    def crear_interfaz_seleccion_archivo(self):
        """Genera el banner superior para la configuración y enlace de archivos de datos."""
        frame_top = ttk.LabelFrame(self, text=" Configuración de Archivos Base (BBDD Excel) ", padding=10)
        frame_top.pack(fill="x", padx=15, pady=10)
        
        # Maestro
        ttk.Label(frame_top, text="Excel Maestro (Inventario):").grid(row=0, column=0, sticky="w", pady=2)
        entry_m = ttk.Entry(frame_top, textvariable=self.ruta_maestro, width=60)
        entry_m.grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(frame_top, text="Examinar", command=self.examinar_maestro).grid(row=0, column=2, padx=2, pady=2)
        
        # Reportes
        ttk.Label(frame_top, text="Excel Reportes Mensuales:").grid(row=1, column=0, sticky="w", pady=2)
        entry_r = ttk.Entry(frame_top, textvariable=self.ruta_reportes, width=60)
        entry_r.grid(row=1, column=1, padx=5, pady=2)
        ttk.Button(frame_top, text="Examinar", command=self.examinar_reportes).grid(row=1, column=2, padx=2, pady=2)

    def examinar_maestro(self):
        ruta = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
        if ruta:
            self.ruta_maestro.set(ruta)
            inicializar_excel_maestro(ruta)

    def examinar_reportes(self):
        ruta = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
        if ruta:
            self.ruta_reportes.set(ruta)

    def crear_cuerpo_principal(self):
        """Instancia el contenedor de pestañas (Notebook) de la aplicación."""
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=15, pady=5)
        
        self.tab_ventas = ttk.Frame(self.notebook, padding=10)
        self.tab_inventario = ttk.Frame(self.notebook, padding=10)
        self.tab_reportes = ttk.Frame(self.notebook, padding=10)
        
        self.notebook.add(self.tab_ventas, text="Salida por Ventas")
        self.notebook.add(self.tab_inventario, text="Ajustes de Inventario")
        self.notebook.add(self.tab_reportes, text="Reportes y Ganancias")
        
        self.disenar_tab_ventas()
        self.disenar_tab_inventario()
        self.disenar_tab_reportes()

    # --------------------------------------------------------------------------
    # MÓDULO 1: SALIDA DE PIEZAS POR VENTAS
    # --------------------------------------------------------------------------
    def disenar_tab_ventas(self):
        # Selección de tipo de pieza
        frame_tipo = ttk.LabelFrame(self.tab_ventas, text=" Tipo de Producto ", padding=10)
        frame_tipo.pack(fill="x", pady=5)
        
        self.var_tipo_pieza = tk.StringVar(value="Registrada")
        ttk.Radiobutton(frame_tipo, text="Pieza Registrada (Con Código)", variable=self.var_tipo_pieza, value="Registrada", command=self.alternar_modo_venta).grid(row=0, column=0, padx=10)
        ttk.Radiobutton(frame_tipo, text="Pieza NO Registrada (Venta Rápida)", variable=self.var_tipo_pieza, value="NoRegistrada", command=self.alternar_modo_venta).grid(row=0, column=1, padx=10)
        
        # Formulario de venta
        self.frame_form_ventas = ttk.LabelFrame(self.tab_ventas, text=" Datos de la Transacción ", padding=15)
        self.frame_form_ventas.pack(fill="both", expand=True, pady=10)
        
        # Elementos Dinámicos
        self.lbl_identificador = ttk.Label(self.frame_form_ventas, text="Código de Pieza:")
        self.lbl_identificador.grid(row=0, column=0, sticky="w", pady=5)
        self.ent_identificador = ttk.Entry(self.frame_form_ventas, width=25)
        self.ent_identificador.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        # Botón para buscar producto registrado al vuelo
        self.btn_buscar_p = ttk.Button(self.frame_form_ventas, text="Buscar Código", command=self.evento_buscar_codigo_venta)
        self.btn_buscar_p.grid(row=0, column=2, sticky="w", padx=5)
        
        self.lbl_info_prod = ttk.Label(self.frame_form_ventas, text="", font=("Segoe UI", 9, "italic"), foreground=COLOR_SECONDARY)
        self.lbl_info_prod.grid(row=0, column=3, columnspan=2, sticky="w")

        # Cantidades por método de pago
        ttk.Label(self.frame_form_ventas, text="Cant. Vendida en Efectivo:").grid(row=1, column=0, sticky="w", pady=5)
        self.ent_cant_efectivo = ttk.Entry(self.frame_form_ventas, width=10)
        self.ent_cant_efectivo.insert(0, "0")
        self.ent_cant_efectivo.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(self.frame_form_ventas, text="Cant. Vendida con Tarjeta:").grid(row=1, column=2, sticky="w", pady=5)
        self.ent_cant_tarjeta = ttk.Entry(self.frame_form_ventas, width=10)
        self.ent_cant_tarjeta.insert(0, "0")
        self.ent_cant_tarjeta.grid(row=1, column=3, sticky="w", padx=5, pady=5)

        # Campos específicos para piezas no registradas (inicialmente ocultos o bloqueados)
        self.lbl_precio_manual = ttk.Label(self.frame_form_ventas, text="Precio Unitario ($):")
        self.lbl_precio_manual.grid(row=2, column=0, sticky="w", pady=5)
        self.ent_precio_manual = ttk.Entry(self.frame_form_ventas, width=15)
        self.ent_precio_manual.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        # Descuentos / Modificaciones para registradas
        self.var_tiene_desc = tk.BooleanVar(value=False)
        self.chk_desc = ttk.Checkbutton(self.frame_form_ventas, text="¿Modificar precio o aplicar descuento?", variable=self.var_tiene_desc, command=self.alternar_campos_descuento)
        self.chk_desc.grid(row=3, column=0, columnspan=2, sticky="w", pady=8)
        
        self.lbl_desc = ttk.Label(self.frame_form_ventas, text="Descuento (%) O Nuevo Precio ($):")
        self.lbl_desc.grid(row=4, column=0, sticky="w", pady=5)
        self.ent_desc = ttk.Entry(self.frame_form_ventas, width=15)
        self.ent_desc.grid(row=4, column=1, sticky="w", padx=5, pady=5)
        
        # Inicializar UI de ventas
        self.alternar_modo_venta()
        
        # Botón de Procesar
        btn_procesar = ttk.Button(self.tab_ventas, text="PROCESAR Y REGISTRAR VENTA", style="Accent.TButton", command=self.evento_procesar_venta)
        btn_procesar.pack(fill="x", pady=5)

    def alternar_modo_venta(self):
        modo = self.var_tipo_pieza.get()
        if modo == "Registrada":
            self.lbl_identificador.config(text="Código de Pieza:")
            self.btn_buscar_p.grid()
            self.lbl_precio_manual.grid_remove()
            self.ent_precio_manual.grid_remove()
            self.chk_desc.grid()
            self.alternar_campos_descuento()
        else:
            self.lbl_identificador.config(text="Nombre del Producto:")
            self.btn_buscar_p.grid_remove()
            self.lbl_info_prod.config(text="")
            self.lbl_precio_manual.grid()
            self.ent_precio_manual.grid()
            self.chk_desc.grid_remove()
            self.lbl_desc.grid_remove()
            self.ent_desc.grid_remove()

    def alternar_campos_descuento(self):
        if self.var_tiene_desc.get() and self.var_tipo_pieza.get() == "Registrada":
            self.lbl_desc.grid()
            self.ent_desc.grid()
        else:
            self.lbl_desc.grid_remove()
            self.ent_desc.grid_remove()

    def evento_buscar_codigo_venta(self):
        cod = self.ent_identificador.get().strip()
        if not cod:
            messagebox.showwarning("Atención", "Escriba un código primero.")
            return
        prod = buscar_producto_maestro(self.ruta_maestro.get(), cod)
        if prod:
            self.lbl_info_prod.config(text=f"Encontrado: {prod['Nombre']} | Stock: {prod['Cantidad']} | Precio: ${prod['Precio Venta']}")
            return prod
        else:
            self.lbl_info_prod.config(text="Producto no localizado en el Maestro.", foreground=COLOR_ACCENT)
            return None

    def evento_procesar_venta(self):
        modo = self.var_tipo_pieza.get()
        identificador = self.ent_identificador.get().strip()
        
        try:
            cant_ef = int(self.ent_cant_efectivo.get())
            cant_tj = int(self.ent_cant_tarjeta.get())
        except ValueError:
            messagebox.showerror("Error", "Las cantidades deben ser números enteros.")
            return
            
        cant_total = cant_ef + cant_tj
        if cant_total <= 0:
            messagebox.showerror("Error", "La cantidad total vendida debe ser mayor a 0.")
            return

        precio_unitario = 0.0
        descuento_nota = "Ninguno"
        
        if modo == "Registrada":
            prod = buscar_producto_maestro(self.ruta_maestro.get(), identificador)
            if not prod:
                messagebox.showerror("Error", "Debe ser un código válido y existente en el maestro.")
                return
                
            if prod["Cantidad"] < cant_total:
                messagebox.showerror("Error", f"Stock insuficiente en maestro. Disponible: {prod['Cantidad']}")
                return
                
            precio_unitario = float(prod["Precio Venta"])
            
            # Verificar si aplica descuento/modificación
            if self.var_tiene_desc.get():
                val_desc = self.ent_desc.get().strip()
                if not val_desc:
                    messagebox.showerror("Error", "Escriba el valor de la modificación o descuento.")
                    return
                try:
                    # Si contiene '%' se asume porcentaje de descuento, si no, precio directo modificado
                    if "%" in val_desc:
                        porcentaje = float(val_desc.replace("%", ""))
                        precio_unitario = precio_unitario * (1 - (porcentaje / 100))
                        descuento_nota = f"Desc {porcentaje}%"
                    else:
                        precio_unitario = float(val_desc)
                        descuento_nota = f"Precio Modif a ${precio_unitario}"
                except ValueError:
                    messagebox.showerror("Error", "Formato de descuento incorrecto. Ejemplos: '15%' o '450'")
                    return
                    
            # Actualizar base de datos del maestro (Resta del Stock)
            modificar_stock_maestro(self.ruta_maestro.get(), identificador, -cant_total)
            
        else: # No Registrada
            if not identificador:
                messagebox.showerror("Error", "Escriba el nombre del producto.")
                return
            try:
                precio_unitario = float(self.ent_precio_manual.get())
            except ValueError:
                messagebox.showerror("Error", "El precio unitario debe ser un número válido.")
                return
            descuento_nota = "Venta Rápida (No Reg)"

        # Cálculos de totales
        tot_efectivo = cant_ef * precio_unitario
        tot_tarjeta = cant_tj * precio_unitario
        total_dinero = tot_efectivo + tot_tarjeta
        
        # Data de la Transacción
        venta_obj = {
            "Fecha": datetime.date.today().strftime("%Y-%m-%d"),
            "Identificador": identificador,
            "Tipo": modo,
            "Cantidad": cant_total,
            "Efectivo": tot_efectivo,
            "Tarjeta": tot_tarjeta,
            "Descuento_Aplicado": descuento_nota,
            "Total": total_dinero
        }
        
        # Guardar en Reportes del Año Vigente
        exito = registrar_venta_en_reporte(self.ruta_reportes.get(), venta_obj)
        if exito:
            messagebox.showinfo("Éxito", f"Venta procesada correctamente.\nTotal: ${total_dinero:.2f}")
            # Limpieza de campos
            self.ent_identificador.delete(0, tk.END)
            self.ent_cant_efectivo.delete(0, tk.END)
            self.ent_cant_efectivo.insert(0, "0")
            self.ent_cant_tarjeta.delete(0, tk.END)
            self.ent_cant_tarjeta.insert(0, "0")
            self.ent_precio_manual.delete(0, tk.END)
            self.ent_desc.delete(0, tk.END)
            self.lbl_info_prod.config(text="")

    # --------------------------------------------------------------------------
    # MÓDULO 2: AJUSTES DE INVENTARIO
    # --------------------------------------------------------------------------
    def disenar_tab_inventario(self):
        # Subsección 1: Ajuste de existentes
        frame_existente = ttk.LabelFrame(self.tab_inventario, text=" Ajustar Stock de Pieza Existente ", padding=10)
        frame_existente.pack(fill="x", pady=5)
        
        ttk.Label(frame_existente, text="Código de Pieza:").grid(row=0, column=0, sticky="w", pady=5)
        self.ent_ajuste_cod = ttk.Entry(frame_existente, width=20)
        self.ent_ajuste_cod.grid(row=0, column=1, padx=5)
        
        ttk.Label(frame_existente, text="Cantidad (+/-):").grid(row=0, column=2, sticky="w", pady=5)
        self.ent_ajuste_cant = ttk.Entry(frame_existente, width=10)
        self.ent_ajuste_cant.grid(row=0, column=3, padx=5)
        
        ttk.Button(frame_existente, text="Aplicar Ajuste", command=self.evento_ajustar_stock_existente).grid(row=0, column=4, padx=5)
        
        # Subsección 2: Alta de piezas nuevas
        frame_nueva = ttk.LabelFrame(self.tab_inventario, text=" Registrar Alta de Nueva Pieza ", padding=10)
        frame_nueva.pack(fill="both", expand=True, pady=5)
        
        fields = [
            ("Código de Pieza:", "cod"),
            ("Nombre del Producto:", "nom"),
            ("Departamento / Categoría:", "dep"),
            ("Precio de Venta ($):", "pre"),
            ("Cantidad Inicial:", "can")
        ]
        
        self.dict_alta_entries = {}
        for idx, (label_text, key) in enumerate(fields):
            ttk.Label(frame_nueva, text=label_text).grid(row=idx, column=0, sticky="w", pady=4, padx=5)
            entry = ttk.Entry(frame_nueva, width=30)
            entry.grid(row=idx, column=1, sticky="w", pady=4, padx=5)
            self.dict_alta_entries[key] = entry
            
        ttk.Button(frame_nueva, text="GUARDAR NUEVA PIEZA EN MAESTRO", style="Accent.TButton", command=self.evento_alta_nueva_pieza).grid(row=len(fields), column=0, columnspan=2, pady=15)

    def evento_ajustar_stock_existente(self):
        cod = self.ent_ajuste_cod.get().strip()
        try:
            cambio = int(self.ent_ajuste_cant.get())
        except ValueError:
            messagebox.showerror("Error", "La cantidad debe ser un entero válido (positivo o negativo).")
            return
            
        if not cod:
            messagebox.showerror("Error", "Ingrese un código válido.")
            return
            
        exito, msg = modificar_stock_maestro(self.ruta_maestro.get(), cod, cambio)
        if exito:
            messagebox.showinfo("Éxito", f"Inventario modificado. Nuevo stock: {msg}")
            self.ent_ajuste_cod.delete(0, tk.END)
            self.ent_ajuste_cant.delete(0, tk.END)
        else:
            messagebox.showerror("Error", f"No se pudo realizar el ajuste:\n{msg}")

    def evento_alta_nueva_pieza(self):
        # Captura de datos
        codigo = self.dict_alta_entries["cod"].get().strip()
        nombre = self.dict_alta_entries["nom"].get().strip()
        depto = self.dict_alta_entries["dep"].get().strip().upper()
        
        if not (codigo and nombre and depto):
            messagebox.showerror("Error", "Código, Nombre y Departamento son obligatorios.")
            return
            
        try:
            precio = float(self.dict_alta_entries["pre"].get())
            cantidad = int(self.dict_alta_entries["can"].get())
        except ValueError:
            messagebox.showerror("Error", "Precio y Cantidad deben contener valores numéricos apropiados.")
            return
            
        # Comprobar duplicados redundantes
        if buscar_producto_maestro(self.ruta_maestro.get(), codigo):
            messagebox.showerror("Error", f"El código '{codigo}' ya está asignado a otro producto.")
            return
            
        item_dict = {
            "Código": codigo,
            "Nombre": nombre,
            "Departamento": depto,
            "Precio Venta": precio,
            "Cantidad": cantidad
        }
        
        if guardar_o_actualizar_producto(self.ruta_maestro.get(), item_dict):
            messagebox.showinfo("Éxito", f"Producto '{nombre}' añadido bajo el departamento '{depto}' correctamente.")
            # Limpiar entradas
            for entry in self.dict_alta_entries.values():
                entry.delete(0, tk.END)

    # --------------------------------------------------------------------------
    # MÓDULO 3: REPORTES DE VENTAS Y RENDIMIENTOS
    # --------------------------------------------------------------------------
    def disenar_tab_reportes(self):
        # Filtro por Fecha de Referencia
        frame_filtro = ttk.LabelFrame(self.tab_reportes, text=" Selección de Semana para Análisis ", padding=10)
        frame_filtro.pack(fill="x", pady=5)
        
        ttk.Label(frame_filtro, text="Seleccione un día de la semana de interés (AAAA-MM-DD):").grid(row=0, column=0, sticky="w")
        self.ent_fecha_reporte = ttk.Entry(frame_filtro, width=15)
        self.ent_fecha_reporte.insert(0, datetime.date.today().strftime("%Y-%m-%d"))
        self.ent_fecha_reporte.grid(row=0, column=1, padx=5)
        
        ttk.Button(frame_filtro, text="Generar Reporte Integral", command=self.evento_calcular_reporte_semanal).grid(row=0, column=2, padx=5)
        
        self.lbl_semana_activa = ttk.Label(frame_filtro, text="Semana Evaluada: -", font=("Segoe UI", 10, "bold"), foreground=COLOR_SECONDARY)
        self.lbl_semana_activa.grid(row=1, column=0, columnspan=3, pady=5, sticky="w")

        # Layout del Reporte de Resultados en Pantalla con Scroll
        self.canvas_rep = tk.Canvas(self.tab_reportes, bg=COLOR_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.tab_reportes, orient="vertical", command=self.canvas_rep.yview)
        self.scrollable_frame = ttk.Frame(self.canvas_rep)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas_rep.configure(scrollregion=self.canvas_rep.bbox("all"))
        )
        self.canvas_rep.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas_rep.configure(yscrollcommand=scrollbar.set)
        
        self.canvas_rep.pack(side="left", fill="both", expand=True, pady=5)
        scrollbar.pack(side="right", fill="y")
        
        self.construir_esqueleto_tablas_reporte()

    def construir_esqueleto_tablas_reporte(self):
        # Contenedores para las tablas del informe
        self.frame_t1 = ttk.LabelFrame(self.scrollable_frame, text=" Tabla 1: Ventas de Piezas Registradas ", padding=5)
        self.frame_t1.pack(fill="x", expand=True, pady=5)
        
        self.frame_t2 = ttk.LabelFrame(self.scrollable_frame, text=" Tabla 2: Ventas de Piezas NO Registradas ", padding=5)
        self.frame_t2.pack(fill="x", expand=True, pady=5)
        
        self.frame_totales = ttk.LabelFrame(self.scrollable_frame, text=" Resumen Financiero y Comisión de Ventas ", padding=10)
        self.frame_totales.pack(fill="x", expand=True, pady=5)

    def evento_calcular_reporte_semanal(self):
        try:
            fecha_ref = datetime.datetime.strptime(self.ent_fecha_reporte.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Error", "El formato de fecha debe ser AAAA-MM-DD")
            return
            
        mes_hoja, lunes, domingo = obtener_semana_y_mes_asignado(fecha_ref)
        semana_str = f"{lunes.strftime('%d/%m')} al {domingo.strftime('%d/%m')}"
        self.lbl_semana_activa.config(text=f"Semana Evaluada: Del Lunes {lunes.strftime('%d-%m-%Y')} al Domingo {domingo.strftime('%d-%m-%Y')} (Asignada a: {mes_hoja})")
        
        ruta_rep = self.ruta_reportes.get()
        if not os.path.exists(ruta_rep):
            messagebox.showwarning("Sin Datos", "No existe archivo de reportes registrado aún.")
            return
            
        try:
            # Leer la hoja correspondiente al mes asignado
            df = pd.read_excel(ruta_rep, sheet_name=mes_hoja)
        except Exception:
            # Si no existe la hoja es porque no hay transacciones en ese mes
            df = pd.DataFrame()
            
        if df.empty:
            messagebox.showinfo("Reporte Vacío", "No se encontraron registros de ventas para este período/mes.")
            return
            
        # Filtrar el dataframe por la semana calculada
        # El campo 'Semana (Rango)' coincide exactamente con semana_str
        df_semana = df[df["Semana (Rango)"] == semana_str]
        
        if df_semana.empty:
            messagebox.showinfo("Reporte Vacío", f"No existen ventas en la semana del {semana_str}.")
            return
            
        # Limpiar frames de tablas previos
        for widget in self.frame_t1.winfo_children(): widget.destroy()
        for widget in self.frame_t2.winfo_children(): widget.destroy()
        for widget in self.frame_totales.winfo_children(): widget.destroy()
        
        # --- TABLA 1: PIEZAS REGISTRADAS ---
        df_reg = df_semana[df_semana["Tipo"] == "Registrada"]
        ttk.Label(self.frame_t1, text="Código | Cantidad Vendida | Total ($)", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        if not df_reg.empty:
            # Agrupar por código de pieza
            t1_grouped = df_reg.groupby("Código/Nombre").agg({"Cantidad": "sum", "Total Dinero": "sum"}).reset_index()
            for _, row in t1_grouped.iterrows():
                ttk.Label(self.frame_t1, text=f"• {row['Código/Nombre']}  -->  {row['Cantidad']} u.  -->  ${row['Total Dinero']:.2f}").pack(anchor="w", padx=10)
        else:
            ttk.Label(self.frame_t1, text="No hay ventas de piezas registradas en esta semana.", font=("Segoe UI", 9, "italic")).pack(anchor="w", padx=10)
            
        # --- TABLA 2: PIEZAS NO REGISTRADAS ---
        df_noreg = df_semana[df_semana["Tipo"] == "NoRegistrada"]
        ttk.Label(self.frame_t2, text="Nombre Producto | Cantidad Vendida | Total ($)", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        if not df_noreg.empty:
            t2_grouped = df_noreg.groupby("Código/Nombre").agg({"Cantidad": "sum", "Total Dinero": "sum"}).reset_index()
            for _, row in t2_grouped.iterrows():
                ttk.Label(self.frame_t2, text=f"• {row['Código/Nombre']}  -->  {row['Cantidad']} u.  -->  ${row['Total Dinero']:.2f}").pack(anchor="w", padx=10)
        else:
            ttk.Label(self.frame_t2, text="No hay ventas de piezas sin registro esta semana.", font=("Segoe UI", 9, "italic")).pack(anchor="w", padx=10)
            
        # --- TOTALES Y COMISIONES ---
        gran_total = df_semana["Total Dinero"].sum()
        
        # Distribución contractual
        ganancia_duena = gran_total * 0.50
        ganancia_guias = gran_total * 0.25
        ganancia_encargada = gran_total * 0.25
        
        # TOP 10 Más Vendidos de la semana
        top10_df = df_semana.groupby("Código/Nombre").agg({"Cantidad": "sum"}).sort_values(by="Cantidad", ascending=False).head(10)
        
        # Renderizado Financiero
        ttk.Label(self.frame_totales, text=f"TOTAL GENERAL DE VENTAS: ${gran_total:.2f}", font=("Segoe UI", 12, "bold"), foreground=COLOR_SECONDARY).grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
        
        ttk.Label(self.frame_totales, text="Distribución de Utilidades:", font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky="w", pady=3)
        ttk.Label(self.frame_totales, text=f"- Dueña del Negocio (50%): ${ganancia_duena:.2f}").grid(row=2, column=0, sticky="w", padx=15)
        ttk.Label(self.frame_totales, text=f"- Guías de Turistas (25%): ${ganancia_guias:.2f}").grid(row=3, column=0, sticky="w", padx=15)
        ttk.Label(self.frame_totales, text=f"- Encargada de Galería (25%): ${ganancia_encargada:.2f}").grid(row=4, column=0, sticky="w", padx=15)
        
        # Renderizado TOP 10
        ttk.Label(self.frame_totales, text="Top 10 Productos Más Vendidos:", font=("Segoe UI", 10, "bold")).grid(row=1, column=1, sticky="w", padx=30, pady=3)
        row_i = 2
        if not top10_df.empty:
            for idx, (name, row) in enumerate(top10_df.iterrows(), start=1):
                ttk.Label(self.frame_totales, text=f"{idx}. {name} ({row['Cantidad']} uds)").grid(row=row_i, column=1, sticky="w", padx=45)
                row_i += 1
        else:
            ttk.Label(self.frame_totales, text="Sin datos.").grid(row=2, column=1, sticky="w", padx=45)

# ============================================
# SECCIÓN 5: INICIALIZACIÓN Y MAIN
# ============================================
if __name__ == "__main__":
    # Asegura compatibilidad con resoluciones de alta densidad de pixeles (HiDPI)
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
        
    app = AppGaleria()
    app.mainloop()