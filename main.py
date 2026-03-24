import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import sys
import threading
import os
import subprocess
import pandas as pd

# Add src to python path
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from src import logic
from config import settings

class ConsoleRedirector:
    """Redirects stdout/stderr to a Tkinter widget"""
    def __init__(self, start_idx, text_widget):
        self.text_widget = text_widget
        self.start_idx = start_idx
        self.buffer = ""

    def write(self, string):
        self.text_widget.after(0, self._write, string)

    def _write(self, string):
        self.text_widget.configure(state='normal')
        self.text_widget.insert('end', string)
        self.text_widget.see('end')
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Unificación Henry")
        self.root.geometry("600x700")
        
        # Style
        style = ttk.Style()
        style.theme_use('clam')
        
        # --- Class variables for enrichment tool ---
        self.enrich_base_file = tk.StringVar()
        self.enrich_side_file = tk.StringVar()
        self.enrich_base_key_col = tk.StringVar()
        self.enrich_side_key_col = tk.StringVar()
        self.enrich_side_file_headers = []
        self.enrich_col_search_var = tk.StringVar()

        # Main container for tabs
        notebook = ttk.Notebook(root)
        notebook.pack(pady=10, padx=10, fill="x", expand=False)

        # Tab 1: Main Process
        tab1 = ttk.Frame(notebook)
        notebook.add(tab1, text='  Proceso Principal  ')

        # Tab 2: Enrichment Tool
        tab2 = ttk.Frame(notebook)
        notebook.add(tab2, text='  Herramienta de Cruce  ')
        
        # --- Populate Tabs ---
        self._create_main_process_tab(tab1)
        self._create_enrich_tab(tab2)

        # --- Console Output (Shared) ---
        frame_console = ttk.LabelFrame(root, text=" Registro de Actividad ", padding=10)
        frame_console.pack(fill="both", expand=True, padx=10, pady=5)

        self.console_text = scrolledtext.ScrolledText(frame_console, height=10, state='disabled', font=("Consolas", 9))
        self.console_text.pack(fill="both", expand=True)

        # Redirect stdout/stderr
        sys.stdout = ConsoleRedirector(0, self.console_text)
        sys.stderr = ConsoleRedirector(0, self.console_text)

        print("Sistema listo. Esperando instrucciones...")

    def _create_main_process_tab(self, parent_tab):
        # Frame and widgets for the main process (Stages 1 & 2)
        frame1 = ttk.LabelFrame(parent_tab, text=" ETAPA 1: Procesar Lotes por Carpetas ", padding=10)
        frame1.pack(fill="x", padx=10, pady=5)
        ttk.Label(frame1, text="Coloque los archivos en subcarpetas dentro de 'data/stage1_raw'.").pack(anchor="w")
        self.clean_folder_var = tk.BooleanVar()
        ttk.Checkbutton(frame1, text="Limpiar subcarpetas de entrada al terminar", variable=self.clean_folder_var).pack(anchor="w", pady=5)
        ttk.Button(frame1, text="PROCESAR LOTES POR CARPETAS", command=self.run_stage1).pack(fill="x", pady=5)

        frame2 = ttk.LabelFrame(parent_tab, text=" ETAPA 2: Unificación Final ", padding=10)
        frame2.pack(fill="x", padx=10, pady=5)
        ttk.Label(frame2, text="Une todos los lotes procesados y cruza con la Maestra.").pack(anchor="w")
        ttk.Button(frame2, text="GENERAR MASTER FINAL", command=self.run_stage2).pack(fill="x", pady=10)

        frame3 = ttk.LabelFrame(parent_tab, text=" Accesos Rápidos ", padding=10)
        frame3.pack(fill="x", padx=10, pady=5)
        ttk.Button(frame3, text="Abrir Carpeta Entrada (Stage 1)", command=lambda: os.startfile(settings.STAGE1_RAW_DIR)).pack(side="left", expand=True, fill="x", padx=2)
        ttk.Button(frame3, text="Abrir Carpeta Salida Final", command=lambda: os.startfile(settings.OUTPUT_DIR)).pack(side="left", expand=True, fill="x", padx=2)

    def _create_enrich_tab(self, parent_tab):
        # Frame and widgets for the enrichment tool
        f_files = ttk.LabelFrame(parent_tab, text=" 1. Selección de Archivos ", padding=10)
        f_files.pack(fill="x", padx=10, pady=5)

        ttk.Button(f_files, text="Seleccionar Archivo Base", command=self._select_base_file).pack(fill="x", pady=(0, 2))
        ttk.Label(f_files, textvariable=self.enrich_base_file, foreground="blue", wraplength=550).pack(fill="x", padx=5, pady=(0, 5))
        
        ttk.Button(f_files, text="Seleccionar Archivo de Enriquecimiento", command=self._select_enrich_file).pack(fill="x", pady=(5, 2))
        ttk.Label(f_files, textvariable=self.enrich_side_file, foreground="blue", wraplength=550).pack(fill="x", padx=5)

        f_keys = ttk.LabelFrame(parent_tab, text=" 2. Configuración del Cruce ", padding=10)
        f_keys.pack(fill="x", padx=10, pady=5)

        ttk.Label(f_keys, text="Columna Clave (Archivo Base):").pack(anchor="w")
        self.combo_base_key = ttk.Combobox(f_keys, textvariable=self.enrich_base_key_col, state="readonly")
        self.combo_base_key.pack(fill="x", pady=2)

        ttk.Label(f_keys, text="Columna Clave (Archivo Enriquecimiento):").pack(anchor="w")
        self.combo_side_key = ttk.Combobox(f_keys, textvariable=self.enrich_side_key_col, state="readonly")
        self.combo_side_key.pack(fill="x", pady=2)

        f_cols = ttk.LabelFrame(parent_tab, text=" 3. Columnas a Añadir ", padding=10)
        f_cols.pack(fill="x", padx=10, pady=5)
        
        search_frame = ttk.Frame(f_cols)
        search_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(search_frame, text="Buscar:").pack(side="left", padx=(0, 5))
        search_entry = ttk.Entry(search_frame, textvariable=self.enrich_col_search_var)
        search_entry.pack(fill="x", expand=True)
        search_entry.bind("<KeyRelease>", self._filter_cols_listbox)

        list_frame = ttk.Frame(f_cols)
        list_frame.pack(fill="both", expand=True)
        self.listbox_cols = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, height=5)
        self.listbox_cols.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox_cols.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox_cols.config(yscrollcommand=scrollbar.set)

        f_run = ttk.Frame(parent_tab, padding=10)
        f_run.pack(fill="x", padx=10, pady=5)
        ttk.Button(f_run, text="ENRIQUECER Y GUARDAR ARCHIVO", command=self._run_enrich_process).pack(fill="x")

    def _get_file_headers(self, file_path):
        if not file_path or not os.path.exists(file_path):
            return []
        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            if file_ext in ['.xlsx', '.xls']:
                columns = pd.read_excel(file_path, nrows=0).columns.tolist()
                return [str(c).strip() for c in columns]
            elif file_ext == '.csv':
                columns = pd.read_csv(file_path, nrows=0, sep=',', encoding='utf-8').columns.tolist()
                return [str(c).strip() for c in columns]
        except Exception as e:
            print(f"No se pudieron leer las columnas de {os.path.basename(file_path)}: {e}")
            messagebox.showwarning("Error de Lectura", f"No se pudieron leer las columnas del archivo: {os.path.basename(file_path)}\n\nAsegúrese de que sea un archivo Excel o CSV válido.")
            return []

    def _select_base_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar Archivo Base",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return
        
        self.enrich_base_file.set(file_path)
        headers = self._get_file_headers(file_path)
        self.combo_base_key['values'] = headers
        if headers:
            self.enrich_base_key_col.set(headers[0])
        else:
            self.enrich_base_key_col.set("")


    def _select_enrich_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar Archivo de Enriquecimiento",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        self.enrich_side_file.set(file_path)
        headers = self._get_file_headers(file_path)
        self.enrich_side_file_headers = headers  # Store original headers
        self.combo_side_key['values'] = headers
        self.listbox_cols.delete(0, tk.END)
        self.enrich_col_search_var.set("") # Clear search
        
        if headers:
            self.enrich_side_key_col.set(headers[0])
            for header in headers:
                self.listbox_cols.insert(tk.END, header)
        else:
            self.enrich_side_key_col.set("")

    def _filter_cols_listbox(self, event=None):
        search_term = self.enrich_col_search_var.get().lower()
        
        # Get current selections (the actual strings) before clearing
        selected_items = {self.listbox_cols.get(i) for i in self.listbox_cols.curselection()}

        self.listbox_cols.delete(0, tk.END)
        
        # Repopulate listbox based on search and re-select
        for header in self.enrich_side_file_headers:
            if search_term in header.lower():
                self.listbox_cols.insert(tk.END, header)
                if header in selected_items:
                    self.listbox_cols.selection_set(self.listbox_cols.size() - 1)

    def _run_enrich_process(self):
        # Gather all data from UI
        base_path = self.enrich_base_file.get()
        side_path = self.enrich_side_file.get()
        base_key = self.enrich_base_key_col.get()
        side_key = self.enrich_side_key_col.get()
        
        selected_indices = self.listbox_cols.curselection()
        cols_to_add = [self.listbox_cols.get(i) for i in selected_indices]

        # Validation
        if not all([base_path, side_path, base_key, side_key, cols_to_add]):
            messagebox.showerror("Error", "Todos los campos son obligatorios. Por favor, seleccione archivos, columnas clave y al menos una columna para añadir.")
            return

        # Ask for output file path
        output_path = filedialog.asksaveasfilename(
            title="Guardar Archivo Enriquecido Como...",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not output_path:
            return
        
        print("--- Iniciando proceso de enriquecimiento en segundo plano... ---")
        # Run in thread
        threading.Thread(
            target=self._run_enrich_thread, 
            args=(base_path, side_path, base_key, side_key, cols_to_add, output_path),
            daemon=True
        ).start()

    def _run_enrich_thread(self, base_path, side_path, base_key, side_key, cols_to_add, output_path):
        success = logic.enrich_file(base_path, side_path, base_key, side_key, cols_to_add, output_path)
        if success:
            messagebox.showinfo("Éxito", f"Archivo enriquecido guardado correctamente en:\n{output_path}")
        else:
            messagebox.showerror("Error", "Falló el proceso de enriquecimiento. Revise el registro de actividad para más detalles.")

    def run_stage1(self):
        print(f"\n--- INICIANDO PROCESO DE LOTES POR CARPETAS ---")
        threading.Thread(target=self._process_stage1_thread, daemon=True).start()

    def _process_stage1_thread(self):
        processed_folders, success = logic.process_stage1_by_subfolders()
        
        if self.clean_folder_var.get() and processed_folders:
            print("\n--- Limpiando subcarpetas procesadas ---")
            for folder_path in processed_folders:
                print(f"Limpiando carpeta: {os.path.basename(folder_path)}")
                try:
                    import glob
                    files_to_delete = glob.glob(os.path.join(folder_path, "*"))
                    for f in files_to_delete:
                        os.remove(f)
                    print(f"Carpeta '{os.path.basename(folder_path)}' limpiada.")
                except Exception as e:
                    print(f"Error general limpiando la carpeta '{os.path.basename(folder_path)}': {e}")
        
        if success:
            messagebox.showinfo("Éxito", "Proceso de lotes finalizado. Revise el registro para más detalles.")
        else:
            messagebox.showerror("Error", "No se procesó ningún lote con éxito o no se encontraron datos.")
        
        print("--- Fin del proceso ---")

    def run_stage2(self):
        print(f"\n--- INICIANDO GENERACIÓN MASTER FINAL ---")
        threading.Thread(target=self._process_stage2_thread, daemon=True).start()

    def _process_stage2_thread(self):
        success = logic.process_stage2_consolidation()
        if success:
            messagebox.showinfo("Éxito", "Master Final generado y enriquecido.")
        else:
            messagebox.showerror("Error", "No se pudo generar el Master Final.")
        print("--- Fin del proceso ---")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
