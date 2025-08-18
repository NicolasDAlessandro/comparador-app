import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

class ComparadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Productos")
        self.root.geometry("400x200")

        self.archivo_cargados = None
        self.archivo_nuevos = None

        # Botón para seleccionar productos en Presea
        self.btn_presea = tk.Button(root, text="Seleccionar Productos en Presea", command=self.seleccionar_presea, width=40)
        self.btn_presea.pack(pady=10)

        # Botón para seleccionar productos ingresados
        self.btn_ingresados = tk.Button(root, text="Seleccionar Productos Ingresados", command=self.seleccionar_ingresados, width=40)
        self.btn_ingresados.pack(pady=10)

        # Botón para guardar resultado
        self.btn_guardar = tk.Button(root, text="Generar archivo de faltantes", command=self.generar_resultado, width=40)
        self.btn_guardar.pack(pady=20)

    def seleccionar_presea(self):
        self.archivo_cargados = filedialog.askopenfilename(
            title="Selecciona el Excel de productos en Presea",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.archivo_cargados:
            messagebox.showinfo("Archivo seleccionado", f"Productos en Presea:\n{self.archivo_cargados}")

    def seleccionar_ingresados(self):
        self.archivo_nuevos = filedialog.askopenfilename(
            title="Selecciona el Excel de productos ingresados",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.archivo_nuevos:
            messagebox.showinfo("Archivo seleccionado", f"Productos ingresados:\n{self.archivo_nuevos}")

    def generar_resultado(self):
        if not self.archivo_cargados or not self.archivo_nuevos:
            messagebox.showerror("Error", "Debes seleccionar ambos archivos antes de continuar.")
            return

        try:
            # Leer archivos
            productos_cargados = pd.read_excel(self.archivo_cargados)
            productos_nuevos = pd.read_excel(self.archivo_nuevos, header=None)

            # Tomar las columnas relevantes
            codigos_cargados = productos_cargados["COD_ALFA"].astype(str).str.strip()
            codigos_nuevos = productos_nuevos[0].astype(str).str.strip()

            # Filtrar los que no estén en los cargados
            faltantes = productos_nuevos[~productos_nuevos[0].isin(codigos_cargados)]

            # Renombrar columnas MercadoLibre
            faltantes.columns = ["COD_ALFA", "DETALLE", "STOCK", "PRECIO"]

            # Establecer un nombre de archivo por defecto
            nombre_por_defecto = "diferencias.xlsx"

            # Seleccionar dónde guardar el archivo de salida
            archivo_salida = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=nombre_por_defecto,  # Establecer el nombre por defecto
                title="Guardar archivo de faltantes como..."
            )

            if archivo_salida:
                faltantes.to_excel(archivo_salida, index=False)
                messagebox.showinfo("Éxito", f"Archivo generado con éxito:\n{archivo_salida}")

        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = ComparadorApp(root)
    root.mainloop()
