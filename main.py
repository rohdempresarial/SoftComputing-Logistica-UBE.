import tkinter as tk
from tkinter import messagebox, ttk
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import random
import pandas as pd  # Añadido para manejo de Excel Premium
import os

class SistemaLogisticoMaestroUBE:
    def __init__(self, root):
        self.root = root
        self.root.title("SoftComputing UBE - Proyecto Final (Grupo #6)")
        self.root.geometry("1350x850")
        
        # --- Colores Corporativos ---
        self.bg_side = "#2c3e50"
        self.accent = "#3498db"

        # --- Panel de Control Izquierdo ---
        self.panel = tk.Frame(self.root, width=400, bg=self.bg_side)
        self.panel.pack(side=tk.LEFT, fill=tk.Y)

        tk.Label(self.panel, text="SOFTEMPRESARIAL IA", fg="white", bg=self.bg_side, font=("Arial", 16, "bold")).pack(pady=15)

        # Entradas de Parámetros
        tk.Label(self.panel, text="Día Operativo (Carga):", fg=self.accent, bg=self.bg_side).pack()
        self.ent_dia = tk.Entry(self.panel, justify='center', font=("Arial", 10))
        self.ent_dia.insert(0, "4")
        self.ent_dia.pack(pady=5)

        tk.Label(self.panel, text="Población Evolutiva:", fg="#27ae60", bg=self.bg_side).pack()
        self.ent_pob = tk.Entry(self.panel, justify='center', font=("Arial", 10))
        self.ent_pob.insert(0, "50")
        self.ent_pob.pack(pady=5)

        # BOTÓN GENERAR
        self.btn_run = tk.Button(self.panel, text="🚀 GENERAR RUTA ÓPTIMA", command=self.ejecutar_sistema, 
                                 bg="#e67e22", fg="white", font=("Arial", 11, "bold"), cursor="hand2")
        self.btn_run.pack(pady=15, padx=20, fill=tk.X)

        # --- REPORTE SOFTCOMPUTING ---
        tk.Label(self.panel, text="REPORTE SOFTCOMPUTING", fg="white", bg=self.bg_side, font=("Arial", 9, "bold")).pack()
        self.txt_reporte = tk.Text(self.panel, height=8, width=42, font=("Consolas", 9), bg="#ecf0f1", padx=5, pady=5)
        self.txt_reporte.pack(pady=5, padx=15)
        
        self.progress = ttk.Progressbar(self.panel, orient=tk.HORIZONTAL, length=250, mode='determinate')
        self.progress.pack(pady=10, padx=30, fill=tk.X)

        # Tabla de Itinerario
        tk.Label(self.panel, text="ITINERARIO DE VISITAS", fg="white", bg=self.bg_side, font=("Arial", 9, "bold")).pack()
        self.tabla = ttk.Treeview(self.panel, columns=("Orden", "X", "Y"), show='headings', height=8)
        self.tabla.heading("Orden", text="Visita")
        self.tabla.heading("X", text="Lat (X)")
        self.tabla.heading("Y", text="Lon (Y)")
        self.tabla.column("Orden", width=60, anchor='center')
        self.tabla.column("X", width=95, anchor='center')
        self.tabla.column("Y", width=95, anchor='center')
        self.tabla.pack(pady=10, padx=15)

        # Botones de Exportación
        self.frame_btns = tk.Frame(self.panel, bg=self.bg_side)
        self.frame_btns.pack(pady=10)
        
        # Botón actualizado para Excel Premium
        self.btn_excel = tk.Button(self.frame_btns, text="📊 EXCEL PREMIUM", command=self.exportar_excel, bg="#27ae60", fg="white", width=15)
        self.btn_excel.pack(side=tk.LEFT, padx=5)
        
        self.btn_log = tk.Button(self.frame_btns, text="📝 GUARDAR LOG", command=self.guardar_log, bg="#95a5a6", fg="white", width=15)
        self.btn_log.pack(side=tk.LEFT, padx=5)

        # --- Área Gráfica (Derecha) ---
        self.fig, self.ax = plt.subplots(figsize=(8, 7))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.root)
        self.canvas.get_tk_widget().pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.itinerario_data = []
        self.texto_log = ""

    def ejecutar_sistema(self):
        try:
            # 1. Fase ANN (Predicción)
            dia = float(self.ent_dia.get())
            n_puntos = int(dia * 4 + 5)
            tam_pob = int(self.ent_pob.get())
            coords = np.random.rand(n_puntos, 2) * 100
            
            # 2. Fase GA (Algoritmo Genético)
            poblacion = [random.sample(range(n_puntos), n_puntos) for _ in range(tam_pob)]
            dist_inicial = self.calcular_distancia(poblacion[0], coords)
            
            for gen in range(101):
                poblacion = sorted(poblacion, key=lambda r: self.calcular_distancia(r, coords))
                # Cruce y Mutación
                padres = poblacion[:max(2, tam_pob//2)]
                hijos = []
                while len(hijos) < tam_pob//2:
                    p1, p2 = random.sample(padres, 2)
                    corte = n_puntos // 2
                    hijo = p1[:corte] + [x for x in p2 if x not in p1[:corte]]
                    if random.random() < 0.2:
                        idx = random.sample(range(n_puntos), 2)
                        hijo[idx[0]], hijo[idx[1]] = hijo[idx[1]], hijo[idx[0]]
                    hijos.append(hijo)
                poblacion = padres + hijos

                if gen % 10 == 0:
                    self.progress['value'] = gen
                    self.graficar(poblacion[0], coords, gen)
                    self.root.update()

            mejor_ruta = poblacion[0]
            dist_final = self.calcular_distancia(mejor_ruta, coords)
            ahorro = ((dist_inicial - dist_final) / dist_inicial) * 100

            # Actualizar Reporte y Tabla
            self.itinerario_data = [(i+1, coords[idx][0], coords[idx][1]) for i, idx in enumerate(mejor_ruta)]
            
            self.texto_log = (f"--- REPORTE TÉCNICO SOFTCOMPUTING ---\n"
                              f"Responsable: Grupo #6\n"
                              f"Empresa: SoftEmpresarial\n"
                              f"{'-'*35}\n"
                              f"Clientes detectados: {n_puntos}\n"
                              f"Distancia Inicial: {dist_inicial:.2f} km\n"
                              f"Distancia Óptima: {dist_final:.2f} km\n"
                              f"Mejora de Eficiencia: {ahorro:.2f}%\n"
                              f"Estado: RUTA OPTIMIZADA")

            self.txt_reporte.delete(1.0, tk.END)
            self.txt_reporte.insert(tk.END, self.texto_log)
            self.actualizar_tabla()
            
            messagebox.showinfo("Éxito", "Cálculo Híbrido Finalizado")

        except Exception as e:
            messagebox.showerror("Error", f"Error en procesamiento: {e}")

    def graficar(self, ruta, coords, gen):
        self.ax.clear()
        r_plot = ruta + [ruta[0]]
        self.ax.plot(coords[r_plot, 0], coords[r_plot, 1], color='#3498db', linewidth=2, alpha=0.7, zorder=1)
        self.ax.scatter(coords[:,0], coords[:,1], color='#e74c3c', s=100, edgecolors='black', zorder=2, label="Clientes")
        
        inicio = coords[ruta[0]]
        self.ax.scatter(inicio[0], inicio[1], marker='>', color='#f1c40f', s=450, edgecolors='black', zorder=3, label="Vehículo")
        
        self.ax.legend(loc='upper right', shadow=True)
        self.ax.set_title(f"Visualización Evolutiva - Gen: {gen}")
        self.canvas.draw()

    def actualizar_tabla(self):
        for item in self.tabla.get_children(): self.tabla.delete(item)
        for f in self.itinerario_data:
            self.tabla.insert("", tk.END, values=(f"#{f[0]}", f"{f[1]:.2f}", f"{f[2]:.2f}"))

    def exportar_excel(self):
        if not self.itinerario_data:
            messagebox.showwarning("Atención", "Genere una ruta antes de exportar.")
            return
        
        try:
            nombre_archivo = "Reporte_Logistico_SoftEmpresarial.xlsx"
            # Crear DataFrame
            df = pd.DataFrame(self.itinerario_data, columns=['Visita', 'Latitud_X', 'Longitud_Y'])
            
            # Crear el archivo Excel con motor XlsxWriter
            writer = pd.ExcelWriter(nombre_archivo, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Itinerario', index=False)
            
            workbook  = writer.book
            worksheet = writer.sheets['Itinerario']
            
            # --- CREAR GRÁFICO DENTRO DE EXCEL ---
            chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
            
            # Configurar serie de datos (X = Latitud, Y = Longitud)
            max_row = len(self.itinerario_data) + 1
            chart.add_series({
                'name':       'Ruta Optimizada',
                'categories': ['Itinerario', 1, 1, max_row - 1, 1],
                'values':     ['Itinerario', 1, 2, max_row - 1, 2],
                'line':       {'color': '#3498db', 'width': 2},
                'marker':     {'type': 'circle', 'size': 6, 'border': {'color': 'red'}, 'fill': {'color': 'red'}},
            })
            
            chart.set_title({'name': 'Logística: Visualización de Ruta'})
            chart.set_x_axis({'name': 'Coordenada X', 'major_gridlines': {'visible': True}})
            chart.set_y_axis({'name': 'Coordenada Y', 'major_gridlines': {'visible': True}})
            chart.set_style(10)
            
            # Insertar gráfico en la hoja
            worksheet.insert_chart('E2', chart)
            
            writer.close()
            messagebox.showinfo("Exportar", f"Archivo '{nombre_archivo}' generado exitosamente con gráfico.")
            
        except Exception as e:
            messagebox.showerror("Error Excel", f"No se pudo generar el Excel: {e}")

    def guardar_log(self):
        if not self.texto_log: return
        with open("log_proyecto_final.txt", "w") as f:
            f.write(self.texto_log)
            f.write("\n\nDETALLE DE RUTA:\n")
            for f_data in self.itinerario_data:
                f.write(f"Visita {f_data[0]}: ({f_data[1]:.2f}, {f_data[2]:.2f})\n")
        messagebox.showinfo("Log", "Log de reporte guardado como 'log_proyecto_final.txt'")

    def calcular_distancia(self, r, c):
        d = 0
        for i in range(len(r)-1): d += np.linalg.norm(c[r[i]] - c[r[i+1]])
        d += np.linalg.norm(c[r[-1]] - c[r[0]])
        return d

if __name__ == "__main__":
    root = tk.Tk()
    app = SistemaLogisticoMaestroUBE(root)
    root.mainloop()