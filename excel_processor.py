import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os
import subprocess
import platform

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de Excel - BD BNN y COFACO")
        self.root.geometry("500x300")
        
        # Variables para almacenar las rutas de los archivos
        self.bd_bnn_path = None
        self.bd_cofaco_path = None
        
        # Configurar la interfaz
        self.setup_ui()
        
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        title_label = ttk.Label(main_frame, text="Procesador de Excel", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Botón BD BNN
        self.btn_bd_bnn = ttk.Button(main_frame, text="BD BNN", 
                                    command=self.load_bd_bnn, width=20)
        self.btn_bd_bnn.grid(row=1, column=0, pady=10, padx=5)
        
        self.label_bd_bnn = ttk.Label(main_frame, text="No seleccionado", 
                                     foreground="red")
        self.label_bd_bnn.grid(row=1, column=1, pady=10, padx=5, sticky=tk.W)
        
        # Botón BD COFACO
        self.btn_bd_cofaco = ttk.Button(main_frame, text="BD COFACO", 
                                       command=self.load_bd_cofaco, width=20)
        self.btn_bd_cofaco.grid(row=2, column=0, pady=10, padx=5)
        
        self.label_bd_cofaco = ttk.Label(main_frame, text="No seleccionado", 
                                        foreground="red")
        self.label_bd_cofaco.grid(row=2, column=1, pady=10, padx=5, sticky=tk.W)
        
        # Botón PROCESAR
        self.btn_procesar = ttk.Button(main_frame, text="PROCESAR", 
                                      command=self.procesar, width=20)
        self.btn_procesar.grid(row=3, column=0, columnspan=2, pady=30)
        
        # Área de estado
        self.status_text = tk.Text(main_frame, height=8, width=60)
        self.status_text.grid(row=4, column=0, columnspan=2, pady=(20, 0))
        
        # Scrollbar para el área de estado
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=4, column=2, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
    def log_status(self, message):
        """Agregar mensaje al área de estado"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
        
    def load_bd_bnn(self):
        """Cargar archivo BD BNN"""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo BD BNN",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if file_path:
            self.bd_bnn_path = file_path
            filename = os.path.basename(file_path)
            self.label_bd_bnn.config(text=filename, foreground="green")
            self.log_status(f"BD BNN cargado: {filename}")
        
    def load_bd_cofaco(self):
        """Cargar archivo BD COFACO"""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo BD COFACO",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if file_path:
            self.bd_cofaco_path = file_path
            filename = os.path.basename(file_path)
            self.label_bd_cofaco.config(text=filename, foreground="green")
            self.log_status(f"BD COFACO cargado: {filename}")
    
    def procesar(self):
        """Procesar los archivos Excel"""
        if not self.bd_bnn_path:
            messagebox.showerror("Error", "Debe seleccionar el archivo BD BNN")
            return
            
        try:
            self.log_status("Iniciando procesamiento...")
            
            # Cargar BD BNN (primera hoja)
            self.log_status("Cargando BD BNN...")
            df_bnn = pd.read_excel(self.bd_bnn_path, sheet_name=0)
            self.log_status(f"BD BNN cargado: {df_bnn.shape[0]} filas, {df_bnn.shape[1]} columnas")
            
            # Procesar datos
            df_procesado = self.procesar_bd_bnn(df_bnn)
            
            # Obtener la carpeta donde está el archivo BD BNN
            carpeta_bd_bnn = os.path.dirname(self.bd_bnn_path)
            
            # Guardar archivo procesado en la misma carpeta
            fecha_hoy = datetime.now().strftime("%Y%m%d")
            nombre_archivo = f"BNN_procesado_{fecha_hoy}.xlsx"
            output_path = os.path.join(carpeta_bd_bnn, nombre_archivo)
            
            self.log_status(f"Guardando archivo en: {output_path}")
            df_procesado.to_excel(output_path, index=False)
            
            self.log_status("¡Procesamiento completado exitosamente!")
            
            # Abrir el archivo automáticamente
            self.abrir_archivo(output_path)
            
            messagebox.showinfo("Éxito", f"Archivo guardado y abierto:\n{nombre_archivo}")
            
        except Exception as e:
            error_msg = f"Error durante el procesamiento: {str(e)}"
            self.log_status(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def abrir_archivo(self, ruta_archivo):
        """Abrir archivo Excel automáticamente según el sistema operativo"""
        try:
            sistema = platform.system()
            
            if sistema == "Windows":
                # En Windows, usar start para abrir con la aplicación predeterminada
                os.startfile(ruta_archivo)
            elif sistema == "Darwin":  # macOS
                subprocess.run(["open", ruta_archivo])
            elif sistema == "Linux":
                subprocess.run(["xdg-open", ruta_archivo])
            
            self.log_status(f"Archivo abierto automáticamente: {os.path.basename(ruta_archivo)}")
            
        except Exception as e:
            self.log_status(f"No se pudo abrir automáticamente el archivo: {str(e)}")
            self.log_status("Puede abrir el archivo manualmente desde la carpeta de BD BNN")
    
    def procesar_bd_bnn(self, df):
        """Procesar el DataFrame de BD BNN según las especificaciones"""
        self.log_status("Procesando estructura de datos...")
        
        # Limpiar nombres de columnas (quitar espacios extra)
        df.columns = df.columns.str.strip()
        
        # Definir grupos de tallas según especificaciones
        grupos_tallas = {
            'Grupo1': ['XXS', 'XS', 'S', 'M', 'L', 'XL', 'XXL', '1X'],
            'Grupo2': ['XXSP', 'XSP', 'SP', 'MP', 'LP'],
            'Grupo3': ['XST', 'ST', 'MT', 'LT', 'XLT'],
            'Grupo4': ['M/T', 'L/T', 'XL/T', 'XXL/T']
        }
        
        # Columnas fijas iniciales
        columnas_fijas = [
            'Vendor', 'Order Class', 'PO #', 'Market PO No.', 
            'Buyer Item #', 'Style Description', 'Color Desc', 'Color Code'
        ]
        
        # Columnas adicionales al final
        columnas_finales = [
            'Gap Sku', 'Destination Country', 'PO Channel Desc', 
            'Retail Cost Currency', 'Retail Cost', 'Final Destination'
        ]
        
        # Agrupar por PO # y Color Code
        grupos = df.groupby(['PO #', 'Color Code'])
        
        resultado_filas = []
        
        self.log_status("Agrupando datos por PO # y Color Code...")
        
        for (po_num, color_code), grupo in grupos:
            # Para cada grupo de tallas, crear una fila
            for nombre_grupo, tallas_grupo in grupos_tallas.items():
                # Filtrar el grupo actual por las tallas del grupo
                datos_grupo = grupo[grupo['Size'].isin(tallas_grupo)]
                
                if not datos_grupo.empty:
                    # Crear nueva fila
                    nueva_fila = {}
                    
                    # Agregar columnas fijas (tomar valores del primer registro)
                    primera_fila = datos_grupo.iloc[0]
                    for col in columnas_fijas:
                        nueva_fila[col] = primera_fila[col] if col in df.columns else ''
                    
                    # Agregar todas las tallas horizontalmente (mantener orden completo)
                    todas_tallas = []
                    for tallas in grupos_tallas.values():
                        todas_tallas.extend(tallas)
                    
                    total_tallas = 0  # Para calcular el total de todas las tallas
                    
                    for talla in todas_tallas:
                        if talla in tallas_grupo:
                            # Solo agregar cantidades para las tallas del grupo actual
                            datos_talla = datos_grupo[datos_grupo['Size'] == talla]
                            if not datos_talla.empty:
                                cantidad = datos_talla['Ordered Item Quantity'].sum()
                                if cantidad > 0:
                                    nueva_fila[talla] = cantidad
                                    total_tallas += cantidad
                                else:
                                    nueva_fila[talla] = ''
                            else:
                                nueva_fila[talla] = ''
                        else:
                            # Para tallas de otros grupos, poner en blanco
                            nueva_fila[talla] = ''
                    
                    # Agregar columna TT con el total de tallas
                    nueva_fila['TT'] = total_tallas if total_tallas > 0 else ''
                    
                    # Agregar columnas finales
                    for col in columnas_finales:
                        nueva_fila[col] = primera_fila[col] if col in df.columns else ''
                    
                    resultado_filas.append(nueva_fila)
        
        # Crear DataFrame resultado
        if resultado_filas:
            df_resultado = pd.DataFrame(resultado_filas)
            
            # Reordenar columnas
            todas_tallas = []
            for tallas in grupos_tallas.values():
                todas_tallas.extend(tallas)
            
            # Agregar columna TT después de las tallas y antes de las columnas finales
            orden_columnas = columnas_fijas + todas_tallas + ['TT'] + columnas_finales
            
            # Filtrar columnas que realmente existen
            orden_columnas = [col for col in orden_columnas if col in df_resultado.columns]
            df_resultado = df_resultado[orden_columnas]
            
            self.log_status(f"Procesamiento completado: {len(df_resultado)} filas generadas")
            return df_resultado
        else:
            self.log_status("No se generaron datos procesados")
            return pd.DataFrame()

def main():
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()