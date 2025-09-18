import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class ComparadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Productos")
        self.root.state('zoomed')

        self.archivo_cargados = None
        self.archivo_nuevos = None

        
        self.root.configure(bg="black")

        
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Custom.Treeview",
            background="black",
            foreground="white",
            rowheight=25,
            fieldbackground="black",
            bordercolor="blue",
            borderwidth=1,
        )
        style.configure(
            "Custom.Treeview.Heading",
            background="blue",
            foreground="white",
            font=("Segoe UI", 10, "bold")
        )
        style.map("Custom.Treeview", background=[("selected", "darkblue")])

        
        self.frame_botones = tk.Frame(root, bg="black")
        self.frame_botones.pack(pady=20)

        tk.Button(self.frame_botones, text="Seleccionar Productos en Presea", 
                  command=self.seleccionar_presea, width=50, bg="blue", fg="white").pack(pady=5)
        tk.Button(self.frame_botones, text="Seleccionar Productos Ingresados", 
                  command=self.seleccionar_ingresados, width=50, bg="blue", fg="white").pack(pady=5)
        tk.Button(self.frame_botones, text="Generar archivo de faltantes", 
                  command=self.generar_faltantes, width=50, bg="blue", fg="white").pack(pady=15)
        tk.Button(self.frame_botones, text="Stock Turturici", 
                  command=self.generar_stock_turturici, width=50, bg="green", fg="white").pack(pady=5)

        
        self.frame_info = tk.Frame(root, bg="black")
        self.frame_info.pack(pady=5)

        self.label_presea = tk.Label(self.frame_info, text="Productos en Presea: ❌ No cargado", 
                                     bg="black", fg="white", font=("Segoe UI", 9))
        self.label_presea.pack(anchor="w")

        self.label_ingresados = tk.Label(self.frame_info, text="Productos Ingresados: ❌ No cargado", 
                                         bg="black", fg="white", font=("Segoe UI", 9))
        self.label_ingresados.pack(anchor="w")

        
        self.frame_tabla = tk.Frame(root, bg="black")
        self.frame_tabla.pack(fill="both", expand=True)

        self.df_matcheo = None  

    def seleccionar_presea(self):
        self.archivo_cargados = filedialog.askopenfilename(
            title="Selecciona el Excel de productos en Presea",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.archivo_cargados:
            self.label_presea.config(text=f"Productos en Presea: {self.archivo_cargados}", fg="cyan")
            self.verificar_carga_completa()

    def seleccionar_ingresados(self):
        self.archivo_nuevos = filedialog.askopenfilename(
            title="Selecciona el Excel de productos ingresados",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.archivo_nuevos:
            self.label_ingresados.config(text=f"Productos Ingresados: {self.archivo_nuevos}", fg="cyan")
            self.verificar_carga_completa()

    def verificar_carga_completa(self):
        """Si ya se cargaron ambos archivos, mostrar la tabla automáticamente"""
        if self.archivo_cargados and self.archivo_nuevos:
            self.generar_tabla()

    def generar_tabla(self):
        try:
            productos_cargados = pd.read_excel(self.archivo_cargados)
            productos_nuevos = pd.read_excel(self.archivo_nuevos, header=None)

            
            productos_cargados["COD_ALFA"] = productos_cargados["COD_ALFA"].astype(str).str.strip().str.upper()
            productos_nuevos[0] = productos_nuevos[0].astype(str).str.strip().str.upper()

            mapa_detalle = dict(zip(productos_cargados["COD_ALFA"], productos_cargados["DETALLE"]))
            mapa_codigo_real = dict(zip(productos_cargados["COD_ALFA"], productos_cargados["CODIGO"]))

            matcheo = []
            for _, row in productos_nuevos.iterrows():
                cod_alfa = row[0]
                detalle_mercado = row[1]
                detalle_presea = mapa_detalle.get(cod_alfa, "❌ NO encontrado")
                codigo_real = mapa_codigo_real.get(cod_alfa, "❌ NO encontrado")
                matcheo.append([codigo_real, cod_alfa, detalle_presea, detalle_mercado])

            
            self.df_matcheo = pd.DataFrame(matcheo, columns=["CODIGO", "COD_ALFA", "DETALLE_PRESEA", "TURTURICI"])

            self.mostrar_tabla(self.df_matcheo)

        except Exception as e:
            messagebox.showerror("Error", str(e))
            print("ERROR:", e)


    def generar_faltantes(self):
        """Genera el archivo de faltantes solo si hay DataFrame cargado"""
        if self.df_matcheo is None:
            messagebox.showerror("Error", "Primero debes cargar ambos archivos para ver la tabla.")
            return

        try:
            productos_cargados = pd.read_excel(self.archivo_cargados)
            productos_nuevos = pd.read_excel(self.archivo_nuevos, header=None)
            codigos_cargados = productos_cargados["COD_ALFA"].astype(str).str.strip().str.upper()
            productos_nuevos[0] = productos_nuevos[0].astype(str).str.strip().str.upper()

            faltantes = productos_nuevos[~productos_nuevos[0].isin(codigos_cargados)].copy()
            if not faltantes.empty:
                faltantes.columns = ["COD_ALFA", "DETALLE", "STOCK", "PRECIO"]
                archivo_salida = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Guardar archivo de faltantes como..."
                )
                if archivo_salida:
                    faltantes.to_excel(archivo_salida, index=False)
                    messagebox.showinfo("Éxito", f"Archivo generado con {len(faltantes)} faltantes.")
            else:
                messagebox.showinfo("Info", "No se encontraron faltantes.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            print("ERROR:", e)


    def generar_stock_turturici(self):
        """Genera un archivo Excel con CODIGO (Presea), COD_ALFA y STOCK (Turturici)"""
        if not self.archivo_cargados or not self.archivo_nuevos:
            messagebox.showerror("Error", "Debes cargar ambos archivos primero.")
            return

        try:
            presea = pd.read_excel(self.archivo_cargados)
            mercado = pd.read_excel(self.archivo_nuevos, header=None)

            presea["COD_ALFA"] = presea["COD_ALFA"].astype(str).str.strip().str.upper()
            mercado[0] = mercado[0].astype(str).str.strip().str.upper()

            df_merge = pd.merge(
                presea[["CODIGO", "COD_ALFA"]],
                mercado[[0, 2]],  # 0 = COD_ALFA, 2 = STOCK
                left_on="COD_ALFA",
                right_on=0,
                how="inner"
            ).drop(columns=[0])  # eliminar columna duplicada

            df_merge.rename(columns={2: "STOCK"}, inplace=True)

            df_merge = df_merge.astype({
                "CODIGO": "Int64",    
                "COD_ALFA": "string",  
                "STOCK": "Int64"       
            })

            df_merge = df_merge[df_merge["STOCK"] > 0]

            archivo_salida = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Guardar archivo de Stock Turturici como..."
            )
            if archivo_salida:
                df_merge.to_excel(archivo_salida, index=False)
                messagebox.showinfo("Éxito", f"Archivo generado con {len(df_merge)} registros.")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            print("ERROR:", e)


    def mostrar_tabla(self, df):
        for widget in self.frame_tabla.winfo_children():
            widget.destroy()

        tabla = ttk.Treeview(self.frame_tabla, columns=list(df.columns), show="headings", style="Custom.Treeview")

        tabla.heading("CODIGO", text="CODIGO")
        tabla.column("CODIGO", width=90, anchor="center")

        tabla.heading("COD_ALFA", text="COD_ALFA")
        tabla.column("COD_ALFA", width=90, anchor="center")

        tabla.heading("DETALLE_PRESEA", text="DETALLE PRESEA")
        tabla.column("DETALLE_PRESEA", width=400, anchor="w")

        tabla.heading("TURTURICI", text="TURTURICI")
        tabla.column("TURTURICI", width=400, anchor="w")

        for i, fila in df.iterrows():
            if fila["DETALLE_PRESEA"] == "❌ NO encontrado":
                tag = "notfound"
            elif str(fila["DETALLE_PRESEA"]).strip().upper() != str(fila["TURTURICI"]).strip().upper():
                tag = "mismatch"
            else:
                tag = "evenrow" if i % 2 == 0 else "oddrow"

            tabla.insert("", "end", values=list(fila), tags=(tag,))

        tabla.tag_configure("evenrow", background="black", foreground="white")
        tabla.tag_configure("oddrow", background="#111111", foreground="white")
        tabla.tag_configure("notfound", background="black", foreground="red", font=("Segoe UI", 9, "bold"))
        tabla.tag_configure("mismatch", background="black", foreground="lime", font=("Segoe UI", 9, "bold"))

        tabla.pack(side="left", fill="both", expand=True)

        scrollbar_y = ttk.Scrollbar(self.frame_tabla, orient="vertical", command=tabla.yview)
        tabla.configure(yscroll=scrollbar_y.set)
        scrollbar_y.pack(side="right", fill="y")


if __name__ == "__main__":
    root = tk.Tk()
    app = ComparadorApp(root)
    root.mainloop()
