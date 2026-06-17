# 🎨 Sistema de Gestión de Galería de Arte v1.0

¡Bienvenido al sistema de control de inventario y reportes de ventas para la Galería! Esta aplicación de escritorio está desarrollada en **Python** utilizando **Tkinter** para una interfaz nativa, ligera y moderna, y **Pandas/OpenPyXL** para el manejo estructurado de bases de datos directamente en archivos de Excel (`.xlsx`).

---

## 🚀 Características Principales

* **🛒 Salida por Ventas:** Registra transacciones de piezas en sistema (descontando stock) o ventas rápidas de piezas no registradas. Permite aplicar descuentos porcentuales o modificaciones de precio al vuelo.
* **📦 Ajustes de Inventario:** Da de alta piezas nuevas con la estructura real del catálogo o ajusta las existencias de productos actuales (+/-) por mermas o roturas.
* **📊 Reportes y Comisiones:** Genera reportes semanales automáticos basados en fechas. Calcula el **Top 10** de productos más vendidos y realiza la división exacta de ganancias (50% Dueña, 25% Guías, 25% Encargada).
* **💾 Recálculo Financiero:** Cada movimiento de inventario actualiza automáticamente el valor totalizado del activo en tu Excel maestro.

---

## 📂 Estructura de la Base de Datos (Excel)

El sistema se acopla perfectamente a tu formato de inventario actual, respetando las siguientes columnas:
* `Codigo de producto` | `Departamento` | `Origen` | `Nombre` | `Descripcion` | `Precio por unidad` | `Cantidad en existencias` | `Valor de inventario` | `¿Suspendido?`

> ⚠️ **Nota Importante:** El sistema automatiza la columna *Valor de inventario* mediante la operación: 
> $$\text{Precio por unidad} \times \text{Cantidad en existencias}$$

---

## 🛠️ Requisitos e Instalación

### 1. Clonar o guardar el archivo
Asegúrate de tener el archivo de la aplicación (por ejemplo, `app.py`) en una carpeta limpia.

### 2. Instalar Dependencias
Abre tu terminal o consola de comandos (CMD) dentro de la carpeta del proyecto e instala las librerías necesarias ejecutando:

```bash
pip install openpyxl pandas
