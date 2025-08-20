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

        # 🔹 Fondo general negro
        self.root.configure(bg="black")

        # 🔹 Estilo Treeview personalizado
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

        # Frame botones
        self.frame_botones = tk.Frame(root, bg="black")
        self.frame_botones.pack(pady=20)

        tk.Button(self.frame_botones, text="Seleccionar Productos en Presea", 
                  command=self.seleccionar_presea, width=50, bg="blue", fg="white").pack(pady=5)
        tk.Button(self.frame_botones, text="Seleccionar Productos Ingresados", 
                  command=self.seleccionar_ingresados, width=50, bg="blue", fg="white").pack(pady=5)
        tk.Button(self.frame_botones, text="Generar archivo de faltantes", 
                  command=self.generar_faltantes, width=50, bg="blue", fg="white").pack(pady=15)

        # Frame para info de archivos seleccionados
        self.frame_info = tk.Frame(root, bg="black")
        self.frame_info.pack(pady=5)

        self.label_presea = tk.Label(self.frame_info, text="Productos en Presea: ❌ No cargado", 
                                     bg="black", fg="white", font=("Segoe UI", 9))
        self.label_presea.pack(anchor="w")

        self.label_ingresados = tk.Label(self.frame_info, text="Productos Ingresados: ❌ No cargado", 
                                         bg="black", fg="white", font=("Segoe UI", 9))
        self.label_ingresados.pack(anchor="w")

        # Frame para la tabla
        self.frame_tabla = tk.Frame(root, bg="black")
        self.frame_tabla.pack(fill="both", expand=True)

        self.df_matcheo = None  # se guarda el último DataFrame mostrado

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

            codigos_cargados = productos_cargados["COD_ALFA"].astype(str).str.strip().str.upper()
            productos_nuevos[0] = productos_nuevos[0].astype(str).str.strip().str.upper()

            mapa_presea = dict(zip(codigos_cargados, productos_cargados["DETALLE"]))

            matcheo = []
            for idx, row in productos_nuevos.iterrows():
                codigo = row[0]
                detalle_mercado = row[1]
                detalle_presea = mapa_presea.get(codigo, "❌ NO encontrado")
                matcheo.append([idx + 1, codigo, detalle_presea, detalle_mercado])

            self.df_matcheo = pd.DataFrame(matcheo, columns=["NRO", "CODIGO", "DETALLE_PRESEA", "TURTURICI"])

            # Mostrar tabla
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

    def mostrar_tabla(self, df):
        for widget in self.frame_tabla.winfo_children():
            widget.destroy()

        tabla = ttk.Treeview(self.frame_tabla, columns=list(df.columns), show="headings", style="Custom.Treeview")

        # 🔹 Ajuste de anchos
        tabla.heading("NRO", text="NRO")
        tabla.column("NRO", width=10, anchor="center")

        tabla.heading("CODIGO", text="CODIGO")
        tabla.column("CODIGO", width=90, anchor="center")

        tabla.heading("DETALLE_PRESEA", text="DETALLE PRESEA")
        tabla.column("DETALLE_PRESEA", width=400, anchor="w")

        tabla.heading("TURTURICI", text="TURTURICI")
        tabla.column("TURTURICI", width=400, anchor="w")

        # 🔹 Insertar filas con estilos
        for i, fila in df.iterrows():
            if fila["DETALLE_PRESEA"] == "❌ NO encontrado":
                tag = "notfound"
            elif fila["DETALLE_PRESEA"].strip().upper() != str(fila["TURTURICI"]).strip().upper():
                tag = "mismatch"
            else:
                tag = "evenrow" if i % 2 == 0 else "oddrow"

            tabla.insert("", "end", values=list(fila), tags=(tag,))

        # 🔹 Configuración de colores
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
