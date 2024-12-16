import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import re
import threading

class ExcelProcessorApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de Excel")
        self.root.geometry("600x500")

        # Inicializar variables de archivos
        self.file_paths = {}

        # Crear botones e interfaz
        self.select_button = tk.Button(
            self.root, text="Seleccionar Archivos", command=self.load_files
        )
        self.select_button.pack(pady=10)

        self.tree = ttk.Treeview(self.root, columns=('Archivo', 'Ruta'), show='headings')
        self.tree.heading('Archivo', text='Archivo')
        self.tree.heading('Ruta', text='Ruta')
        self.tree.column('Archivo', width=200)
        self.tree.column('Ruta', width=380)
        self.tree.pack(pady=10)

        self.process_button = tk.Button(
            self.root, text="Procesar Archivos", command=self.process_file, state=tk.DISABLED
        )
        self.process_button.pack(pady=10)

        self.progress = ttk.Progressbar(self.root, orient='horizontal', length=500, mode='determinate')
        self.progress.pack(pady=10)
        self.progress.pack_forget()

        self.status_label = tk.Label(self.root, text="")
        self.status_label.pack(pady=10)
        self.status_label.pack_forget()

    def load_files(self):
        files_to_load = [
            ('FieldServiceReport', 'file_path1', "Seleccione el archivo FieldServiceReport."),
            ('BDD MOPS', 'file_path2', "Seleccione el archivo BDD MOPS."),
            ("'sites report'", 'file_path4', "Seleccione el archivo 'SitesReport'."),
            ("Archivo '0050'", 'file_path5', "Seleccione el archivo '0050'.")
        ]

        for file_label, file_attr, message in files_to_load:
            messagebox.showinfo("Seleccionar Archivo", message)
            file_path = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
            if not file_path:
                messagebox.showwarning("Advertencia", f"¡No se ha seleccionado {file_label}!")
                return
            setattr(self, file_attr, file_path)
            self.tree.insert('', 'end', values=(file_label, os.path.basename(file_path)))
        self.process_button.config(state=tk.NORMAL)

    def process_file(self):
        threading.Thread(target=self.process_file_thread).start()

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()

    def process_file_thread(self):
        try:
            self.progress.pack(pady=10)
            self.status_label.pack(pady=10)

            self.update_status("Leyendo documentos...")
            df = pd.read_excel(self.file_path1)
            df_bdd_mops = pd.read_excel(self.file_path2, header=1)
            df_sites_report = pd.read_excel(self.file_path4)
            df_0050 = pd.read_excel(self.file_path5, header=0)

            self.update_status("Preprocesando datos...")
            df_bdd_mops['FOLIO'] = df_bdd_mops['FOLIO'].astype(str).str.strip()
            df_sites_report['SITE'] = df_sites_report['SITE'].astype(str).str.strip()
            df_0050['service_order'] = df_0050['service_order'].astype(str).str.strip()

            df['Fecha Inicio'] = pd.to_datetime(df['Fecha Inicio'], errors='coerce')
            df['Mes'] = df['Fecha Inicio'].dt.month.apply(lambda x: self.month_name_spanish(x) if not pd.isna(x) else 'Desconocido')
            df['Q'] = df['Fecha Inicio'].dt.month.apply(lambda x: self.calculate_quarter(x) if not pd.isna(x) else 'Desconocido')

            df['ID_EVIDENCIA'] = df['Evidencia'].apply(self.extract_id_panda).fillna(-1).astype(int)

            df = df.merge(df_bdd_mops[['FOLIO', 'FORM ID', 'STATUS']], how='left', left_on='Folio', right_on='FOLIO')
            df.rename(columns={'FORM ID': 'ID_MOP', 'STATUS': 'STATUS_MOP'}, inplace=True)
            df.drop('FOLIO', axis=1, inplace=True)
            df['ID_MOP'] = pd.to_numeric(df['ID_MOP'], errors='coerce').fillna(-1).astype(int)

            df = df.merge(df_sites_report[['SITE', 'CELL REGION', 'CELLOWNER']], how='left', left_on='Predio', right_on='SITE')
            df.drop('SITE', axis=1, inplace=True)

            df = df.merge(df_0050[['service_order', 'id', 'status']], how='left', left_on='Folio', right_on='service_order')
            df.rename(columns={'id': 'ID_request0050', 'status': 'STATUS_request0050'}, inplace=True)
            df.drop('service_order', axis=1, inplace=True)

            self.update_status("Calculando confirmaciones y estatus finales...")
            df['Confirmacion'] = df['ID_MOP'] == df['ID_EVIDENCIA']

            df['Estatus Actividad'] = df['Estatus Actividad'].astype(str).str.strip().str.lower()
            df['STATUS_MOP'] = df['STATUS_MOP'].astype(str).str.strip().str.lower()

            conditions_odk = [
                df['Confirmacion'] & (df['ID_MOP'] != -1),
                (df['Estatus Actividad'] == 'accepted') & (df['STATUS_MOP'] != 'accepted'),
                (df['STATUS_MOP'] == 'accepted') & (df['Estatus Actividad'] != 'accepted'),
                (df['Estatus Actividad'] == 'accepted') & (df['STATUS_MOP'] == 'accepted') & ~df['Confirmacion']
            ]

            choices_odk = [
                df['ID_MOP'].astype(str),
                df['ID_EVIDENCIA'].astype(str),
                df['ID_MOP'].astype(str),
                'NA'
            ]

            choices_status = [
                df['STATUS_MOP'].str.capitalize(),
                df['Estatus Actividad'].str.capitalize(),
                df['STATUS_MOP'].str.capitalize(),
                'NA'
            ]

            df['ODK Final'] = np.select(conditions_odk, choices_odk, default='Cancelado')
            df['STATUS FINAL'] = np.select(conditions_odk, choices_status, default='Cancelado')

            # Eliminar columnas innecesarias
            columns_to_drop = [
                'Tipo de Actividad', 'Tipo de Servicio', 'C4i', 'Fecha Fin', 'Nombre Sitio',
                'Recepcionado', 'Check In Ing.', 'Check Out Ing.', 'Local Time Check In Ing.',
                'Local Time Check Out Ing.', 'Zona Horaria', 'Penalizacion', 'SLA Description',
                'Duration', 'TimeoutSLA', 'Nombre de la regla', 'Descripción de la regla',
                'Área del proveedor', 'Subservicio', 'Comentario'
            ]
            df.drop(columns=[col for col in columns_to_drop if col in df.columns], inplace=True)

            # Ordenar Columnas...
            desired_order = [
                'Fecha Inicio', 'Folio', 'Proveedor', 'Dominio', 'MOP', 'Prioridad', 'Predio',
                'Evidencia', 'Usuario Solicitante', 'Coordinacion', 'Gerencia', 'Region',
                'Cumplimiento (SLA)', 'Mes', 'Q', 'ID_EVIDENCIA', 'Estatus Actividad', 'ID_MOP',
                'STATUS_MOP', 'Confirmacion', 'ODK Final', 'STATUS FINAL', 'CELL REGION',
                'CELLOWNER', 'ID_request0050', 'STATUS_request0050'
            ]

            # Reorganiza las columnas según el orden deseado
            df = df[desired_order]

            # Reemplazar valores -1 y NaN por ''
            df.replace(-1, '', inplace=True)       # Reemplaza -1 con ''
            #df.fillna('', inplace=True)            # Reemplaza NaN con ''


            self.update_status("Guardando archivo procesado...")
            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")])
            if output_file:
                df.to_excel(output_file, index=False, engine='openpyxl')
                messagebox.showinfo("Éxito", f"Archivo guardado exitosamente: {os.path.basename(output_file)}")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")
        finally:
            self.progress.pack_forget()

    def calculate_quarter(self, month):
        return {1: 'Q1', 2: 'Q1', 3: 'Q1', 4: 'Q2', 5: 'Q2', 6: 'Q2', 7: 'Q3', 8: 'Q3', 9: 'Q3', 10: 'Q4', 11: 'Q4', 12: 'Q4'}.get(month, 'Desconocido')

    def month_name_spanish(self, month):
        return {1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'}.get(month, 'Desconocido')

    def extract_id_panda(self, evidencia):
        match = re.search(r'id=(\d+)', str(evidencia))
        return float(match.group(1)) if match else np.nan

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()



        # Filtrar solo filas con STATUS Accepted
#df_bdd_mops = df_bdd_mops[df_bdd_mops['STATUS'].str.lower() == 'accepted']