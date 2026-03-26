import glob
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd

# Add src to python path
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from src import logic


class ConsoleRedirector:
    """Redirect stdout/stderr to a Tkinter widget."""

    def __init__(self, start_idx, text_widget):
        self.text_widget = text_widget
        self.start_idx = start_idx

    def write(self, string):
        self.text_widget.after(0, self._write, string)

    def _write(self, string):
        self.text_widget.configure(state="normal")
        self.text_widget.insert("end", string)
        self.text_widget.see("end")
        self.text_widget.configure(state="disabled")

    def flush(self):
        pass


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Unificacion Henry")
        self.root.geometry("900x820")
        self.root.rowconfigure(0, weight=3)
        self.root.rowconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("clam")

        # Enrichment state
        self.enrich_base_file = tk.StringVar()
        self.enrich_side_file = tk.StringVar()
        self.enrich_base_key_col = tk.StringVar()
        self.enrich_side_key_col = tk.StringVar()
        self.enrich_col_search_var = tk.StringVar()
        self.enrich_base_headers = []
        self.enrich_side_headers = []
        self.enrich_selected_add_cols = []
        self.enrich_selected_drop_cols = []

        # Main process state
        self.stage1_input_dir = tk.StringVar(value="")
        self.stage1_output_dir = tk.StringVar(value="")
        self.stage2_input_dir = tk.StringVar(value="")
        self.stage2_enable_filter_var = tk.BooleanVar(value=False)
        self.stage2_filter_column_var = tk.StringVar(value="Tipo")
        self.stage2_filter_value_var = tk.StringVar()
        self.stage2_filter_values = []
        self.stage2_available_headers = []
        self.output_numeric_format_var = tk.StringVar(value="excel")

        notebook = ttk.Notebook(root)
        notebook.grid(row=0, column=0, pady=(10, 5), padx=10, sticky="nsew")

        tab1 = ttk.Frame(notebook)
        tab2 = ttk.Frame(notebook)
        notebook.add(tab1, text="  Proceso Principal  ")
        notebook.add(tab2, text="  Herramienta de Cruce  ")

        self._create_main_process_tab(tab1)
        self._create_enrich_tab(tab2)

        frame_console = ttk.LabelFrame(root, text=" Registro de Actividad ", padding=10)
        frame_console.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        frame_console.rowconfigure(0, weight=1)
        frame_console.columnconfigure(0, weight=1)

        self.console_text = scrolledtext.ScrolledText(
            frame_console,
            height=10,
            state="disabled",
            font=("Consolas", 9),
        )
        self.console_text.grid(row=0, column=0, sticky="nsew")

        sys.stdout = ConsoleRedirector(0, self.console_text)
        sys.stderr = ConsoleRedirector(0, self.console_text)

        print("Sistema listo. Esperando instrucciones...")

    def _create_main_process_tab(self, parent_tab):
        frame1 = ttk.LabelFrame(parent_tab, text=" ETAPA 1: Procesar Lotes por Carpetas ", padding=10)
        frame1.pack(fill="x", padx=10, pady=5)

        ttk.Button(
            frame1,
            text="Seleccionar Carpeta de Entrada (Lotes)",
            command=lambda: self._select_dir(self.stage1_input_dir),
        ).pack(fill="x", pady=(0, 2))
        ttk.Label(frame1, textvariable=self.stage1_input_dir, foreground="blue", wraplength=820).pack(
            fill="x", padx=5, pady=(0, 5)
        )

        ttk.Button(
            frame1,
            text="Seleccionar Carpeta de Destino (Lotes Procesados)",
            command=lambda: self._select_dir(self.stage1_output_dir),
        ).pack(fill="x", pady=(5, 2))
        ttk.Label(frame1, textvariable=self.stage1_output_dir, foreground="blue", wraplength=820).pack(
            fill="x", padx=5, pady=(0, 10)
        )

        self.clean_folder_var = tk.BooleanVar()
        ttk.Checkbutton(
            frame1,
            text="Limpiar subcarpetas de entrada al terminar",
            variable=self.clean_folder_var,
        ).pack(anchor="w", pady=5)
        self._create_output_format_selector(frame1)
        ttk.Button(frame1, text="PROCESAR LOTES POR CARPETAS", command=self.run_stage1).pack(fill="x", pady=5)

        frame2 = ttk.LabelFrame(parent_tab, text=" ETAPA 2: Unificacion Final ", padding=10)
        frame2.pack(fill="both", expand=True, padx=10, pady=5)

        ttk.Button(
            frame2,
            text="Seleccionar Carpeta de Lotes Procesados",
            command=self._select_stage2_dir,
        ).pack(fill="x", pady=(0, 2))
        ttk.Label(frame2, textvariable=self.stage2_input_dir, foreground="blue", wraplength=820).pack(
            fill="x", padx=5, pady=(0, 8)
        )

        filter_frame = ttk.LabelFrame(frame2, text=" Filtro Opcional del Consolidado ", padding=10)
        filter_frame.pack(fill="x", pady=5)

        ttk.Checkbutton(
            filter_frame,
            text="Aplicar filtro por columna y lista de valores",
            variable=self.stage2_enable_filter_var,
        ).pack(anchor="w", pady=(0, 8))

        top_filter_row = ttk.Frame(filter_frame)
        top_filter_row.pack(fill="x", pady=(0, 8))

        ttk.Label(top_filter_row, text="Columna a filtrar:").pack(side="left")
        self.stage2_filter_column_combo = ttk.Combobox(
            top_filter_row,
            textvariable=self.stage2_filter_column_var,
            state="readonly",
            width=35,
        )
        self.stage2_filter_column_combo.pack(side="left", padx=(8, 8))
        ttk.Button(top_filter_row, text="Recargar columnas", command=self._load_stage2_headers).pack(side="left")

        ttk.Label(
            filter_frame,
            text="Valores permitidos. Solo se conservaran filas donde la columna tenga alguno de estos valores.",
        ).pack(anchor="w")

        values_row = ttk.Frame(filter_frame)
        values_row.pack(fill="x", pady=(5, 5))

        ttk.Entry(values_row, textvariable=self.stage2_filter_value_var).pack(side="left", fill="x", expand=True)
        ttk.Button(values_row, text="Agregar valor", command=self._add_stage2_filter_value).pack(side="left", padx=(8, 0))

        self.stage2_filter_values_listbox = tk.Listbox(filter_frame, height=5, exportselection=False)
        self.stage2_filter_values_listbox.pack(fill="x", pady=(0, 5))

        values_actions = ttk.Frame(filter_frame)
        values_actions.pack(fill="x")
        ttk.Button(values_actions, text="Quitar seleccionados", command=self._remove_stage2_filter_values).pack(
            side="left"
        )
        ttk.Button(values_actions, text="Limpiar lista", command=self._clear_stage2_filter_values).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(values_actions, text="Cargar 'Materiales'", command=lambda: self._add_stage2_filter_value("Materiales")).pack(
            side="left", padx=(8, 0)
        )

        ttk.Button(frame2, text="UNIR LOTES Y GUARDAR", command=self.run_stage2).pack(fill="x", pady=10)

    def _create_enrich_tab(self, parent_tab):
        f_files = ttk.LabelFrame(parent_tab, text=" 1. Seleccion de Archivos ", padding=10)
        f_files.pack(fill="x", padx=10, pady=5)

        ttk.Button(f_files, text="Seleccionar Archivo Base", command=self._select_base_file).pack(fill="x", pady=(0, 2))
        ttk.Label(f_files, textvariable=self.enrich_base_file, foreground="blue", wraplength=820).pack(
            fill="x", padx=5, pady=(0, 5)
        )

        ttk.Button(
            f_files,
            text="Seleccionar Archivo de Enriquecimiento",
            command=self._select_enrich_file,
        ).pack(fill="x", pady=(5, 2))
        ttk.Label(f_files, textvariable=self.enrich_side_file, foreground="blue", wraplength=820).pack(fill="x", padx=5)

        f_keys = ttk.LabelFrame(parent_tab, text=" 2. Configuracion del Cruce ", padding=10)
        f_keys.pack(fill="x", padx=10, pady=5)

        ttk.Label(f_keys, text="Columna Clave (Archivo Base):").pack(anchor="w")
        self.combo_base_key = ttk.Combobox(f_keys, textvariable=self.enrich_base_key_col, state="readonly")
        self.combo_base_key.pack(fill="x", pady=2)

        ttk.Label(f_keys, text="Columna Clave (Archivo Enriquecimiento):").pack(anchor="w")
        self.combo_side_key = ttk.Combobox(f_keys, textvariable=self.enrich_side_key_col, state="readonly")
        self.combo_side_key.pack(fill="x", pady=2)

        f_add = ttk.LabelFrame(parent_tab, text=" 3. Columnas a Agregar ", padding=10)
        f_add.pack(fill="both", expand=True, padx=10, pady=5)

        search_frame = ttk.Frame(f_add)
        search_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(search_frame, text="Buscar en archivo de enriquecimiento:").pack(side="left", padx=(0, 5))
        search_entry = ttk.Entry(search_frame, textvariable=self.enrich_col_search_var)
        search_entry.pack(fill="x", expand=True)
        search_entry.bind("<KeyRelease>", self._filter_cols_listbox)

        add_lists = ttk.Frame(f_add)
        add_lists.pack(fill="both", expand=True)

        available_add_frame = ttk.Frame(add_lists)
        available_add_frame.pack(side="left", fill="both", expand=True)
        ttk.Label(available_add_frame, text="Disponibles").pack(anchor="w")
        self.listbox_available_add = tk.Listbox(
            available_add_frame,
            selectmode=tk.EXTENDED,
            height=8,
            exportselection=False,
        )
        self.listbox_available_add.pack(fill="both", expand=True)

        add_actions = ttk.Frame(add_lists)
        add_actions.pack(side="left", fill="y", padx=8)
        ttk.Button(add_actions, text="Agregar >", command=self._add_selected_enrich_columns).pack(fill="x", pady=(26, 4))
        ttk.Button(add_actions, text="< Quitar", command=self._remove_selected_enrich_columns).pack(fill="x")

        selected_add_frame = ttk.Frame(add_lists)
        selected_add_frame.pack(side="left", fill="both", expand=True)
        ttk.Label(selected_add_frame, text="Se agregaran").pack(anchor="w")
        self.listbox_selected_add = tk.Listbox(
            selected_add_frame,
            selectmode=tk.EXTENDED,
            height=8,
            exportselection=False,
        )
        self.listbox_selected_add.pack(fill="both", expand=True)

        f_drop = ttk.LabelFrame(parent_tab, text=" 4. Columnas a Quitar del Resultado Final ", padding=10)
        f_drop.pack(fill="both", expand=True, padx=10, pady=5)

        drop_lists = ttk.Frame(f_drop)
        drop_lists.pack(fill="both", expand=True)

        available_drop_frame = ttk.Frame(drop_lists)
        available_drop_frame.pack(side="left", fill="both", expand=True)
        ttk.Label(available_drop_frame, text="Columnas base disponibles").pack(anchor="w")
        self.listbox_available_drop = tk.Listbox(
            available_drop_frame,
            selectmode=tk.EXTENDED,
            height=6,
            exportselection=False,
        )
        self.listbox_available_drop.pack(fill="both", expand=True)

        drop_actions = ttk.Frame(drop_lists)
        drop_actions.pack(side="left", fill="y", padx=8)
        ttk.Button(drop_actions, text="Quitar >", command=self._add_drop_columns).pack(fill="x", pady=(26, 4))
        ttk.Button(drop_actions, text="< Dejar", command=self._remove_drop_columns).pack(fill="x")

        selected_drop_frame = ttk.Frame(drop_lists)
        selected_drop_frame.pack(side="left", fill="both", expand=True)
        ttk.Label(selected_drop_frame, text="Se quitaran").pack(anchor="w")
        self.listbox_selected_drop = tk.Listbox(
            selected_drop_frame,
            selectmode=tk.EXTENDED,
            height=6,
            exportselection=False,
        )
        self.listbox_selected_drop.pack(fill="both", expand=True)

        f_run = ttk.Frame(parent_tab, padding=10)
        f_run.pack(fill="x", padx=10, pady=5)
        self._create_output_format_selector(f_run)
        ttk.Button(f_run, text="ENRIQUECER Y GUARDAR ARCHIVO", command=self._run_enrich_process).pack(fill="x")

    def _create_output_format_selector(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(6, 4))
        ttk.Label(row, text="Formato final de columnas numericas:").pack(side="left")
        ttk.Radiobutton(
            row,
            text="Numero Excel",
            value="excel",
            variable=self.output_numeric_format_var,
        ).pack(side="left", padx=(8, 4))
        ttk.Radiobutton(
            row,
            text="Texto con coma",
            value="comma_text",
            variable=self.output_numeric_format_var,
        ).pack(side="left", padx=4)

    def _select_dir(self, string_var):
        dir_path = filedialog.askdirectory(title="Seleccionar Carpeta")
        if dir_path:
            string_var.set(dir_path)

    def _select_stage2_dir(self):
        dir_path = filedialog.askdirectory(title="Seleccionar Carpeta de Lotes Procesados")
        if not dir_path:
            return
        self.stage2_input_dir.set(dir_path)
        self._load_stage2_headers()

    def _get_file_headers(self, file_path):
        if not file_path or not os.path.exists(file_path):
            return []
        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            if file_ext in [".xlsx", ".xls"]:
                columns = pd.read_excel(file_path, nrows=0).columns.tolist()
                return [str(c).strip() for c in columns]
            if file_ext == ".csv":
                try:
                    columns = pd.read_csv(file_path, nrows=0, sep=",", encoding="utf-8").columns.tolist()
                except Exception:
                    columns = pd.read_csv(file_path, nrows=0, sep=";", encoding="latin1").columns.tolist()
                return [str(c).strip() for c in columns]
        except Exception as e:
            print(f"No se pudieron leer las columnas de {os.path.basename(file_path)}: {e}")
            messagebox.showwarning(
                "Error de Lectura",
                f"No se pudieron leer las columnas del archivo: {os.path.basename(file_path)}",
            )
        return []

    def _load_stage2_headers(self):
        input_dir = self.stage2_input_dir.get()
        self.stage2_available_headers = []
        self.stage2_filter_column_combo["values"] = []

        if not os.path.isdir(input_dir):
            return

        input_files = sorted(glob.glob(os.path.join(input_dir, "*.xlsx")))
        if not input_files:
            print("No se encontraron archivos .xlsx para detectar columnas en la etapa 2.")
            return

        headers = self._get_file_headers(input_files[0])
        self.stage2_available_headers = headers
        self.stage2_filter_column_combo["values"] = headers

        if "Tipo" in headers:
            self.stage2_filter_column_var.set("Tipo")
        elif headers:
            self.stage2_filter_column_var.set(headers[0])
        else:
            self.stage2_filter_column_var.set("")

    def _populate_listbox(self, listbox, items):
        listbox.delete(0, tk.END)
        for item in items:
            listbox.insert(tk.END, item)

    def _select_base_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar Archivo Base",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not file_path:
            return

        self.enrich_base_file.set(file_path)
        headers = self._get_file_headers(file_path)
        self.enrich_base_headers = headers
        self.combo_base_key["values"] = headers

        if headers:
            self.enrich_base_key_col.set(headers[0])
        else:
            self.enrich_base_key_col.set("")

        self.enrich_selected_drop_cols = [
            col for col in self.enrich_selected_drop_cols if col in self.enrich_base_headers
        ]
        self._refresh_drop_columns_ui()

    def _select_enrich_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar Archivo de Enriquecimiento",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not file_path:
            return

        self.enrich_side_file.set(file_path)
        headers = self._get_file_headers(file_path)
        self.enrich_side_headers = headers
        self.combo_side_key["values"] = headers

        if headers:
            self.enrich_side_key_col.set(headers[0])
        else:
            self.enrich_side_key_col.set("")

        self.enrich_selected_add_cols = [
            col for col in self.enrich_selected_add_cols if col in self.enrich_side_headers
        ]
        self.enrich_col_search_var.set("")
        self._refresh_enrich_columns_ui()

    def _refresh_enrich_columns_ui(self):
        search_term = self.enrich_col_search_var.get().strip().lower()
        available = []
        for header in self.enrich_side_headers:
            if header in self.enrich_selected_add_cols:
                continue
            if search_term and search_term not in header.lower():
                continue
            available.append(header)

        self._populate_listbox(self.listbox_available_add, available)
        self._populate_listbox(self.listbox_selected_add, self.enrich_selected_add_cols)

    def _refresh_drop_columns_ui(self):
        available = [col for col in self.enrich_base_headers if col not in self.enrich_selected_drop_cols]
        self._populate_listbox(self.listbox_available_drop, available)
        self._populate_listbox(self.listbox_selected_drop, self.enrich_selected_drop_cols)

    def _filter_cols_listbox(self, event=None):
        self._refresh_enrich_columns_ui()

    def _add_selected_enrich_columns(self):
        selected = [self.listbox_available_add.get(i) for i in self.listbox_available_add.curselection()]
        for col in selected:
            if col not in self.enrich_selected_add_cols:
                self.enrich_selected_add_cols.append(col)
        self._refresh_enrich_columns_ui()

    def _remove_selected_enrich_columns(self):
        selected = [self.listbox_selected_add.get(i) for i in self.listbox_selected_add.curselection()]
        self.enrich_selected_add_cols = [col for col in self.enrich_selected_add_cols if col not in selected]
        self._refresh_enrich_columns_ui()

    def _add_drop_columns(self):
        selected = [self.listbox_available_drop.get(i) for i in self.listbox_available_drop.curselection()]
        for col in selected:
            if col not in self.enrich_selected_drop_cols:
                self.enrich_selected_drop_cols.append(col)
        self._refresh_drop_columns_ui()

    def _remove_drop_columns(self):
        selected = [self.listbox_selected_drop.get(i) for i in self.listbox_selected_drop.curselection()]
        self.enrich_selected_drop_cols = [col for col in self.enrich_selected_drop_cols if col not in selected]
        self._refresh_drop_columns_ui()

    def _add_stage2_filter_value(self, value=None):
        raw_value = value if value is not None else self.stage2_filter_value_var.get()
        clean_value = str(raw_value).strip()
        if not clean_value:
            return
        if clean_value not in self.stage2_filter_values:
            self.stage2_filter_values.append(clean_value)
            self._populate_listbox(self.stage2_filter_values_listbox, self.stage2_filter_values)
        self.stage2_filter_value_var.set("")

    def _remove_stage2_filter_values(self):
        selected = [self.stage2_filter_values_listbox.get(i) for i in self.stage2_filter_values_listbox.curselection()]
        self.stage2_filter_values = [value for value in self.stage2_filter_values if value not in selected]
        self._populate_listbox(self.stage2_filter_values_listbox, self.stage2_filter_values)

    def _clear_stage2_filter_values(self):
        self.stage2_filter_values = []
        self._populate_listbox(self.stage2_filter_values_listbox, self.stage2_filter_values)

    def _run_enrich_process(self):
        base_path = self.enrich_base_file.get()
        side_path = self.enrich_side_file.get()
        base_key = self.enrich_base_key_col.get()
        side_key = self.enrich_side_key_col.get()
        cols_to_add = list(self.enrich_selected_add_cols)
        cols_to_drop = list(self.enrich_selected_drop_cols)
        output_numeric_format = self.output_numeric_format_var.get()

        if not all([base_path, side_path, base_key, side_key]) or not cols_to_add:
            messagebox.showerror(
                "Error",
                "Seleccione ambos archivos, columnas clave y al menos una columna para agregar.",
            )
            return

        output_path = filedialog.asksaveasfilename(
            title="Guardar Archivo Enriquecido Como...",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            return

        print("--- Iniciando proceso de enriquecimiento en segundo plano... ---")
        threading.Thread(
            target=self._run_enrich_thread,
            args=(
                base_path,
                side_path,
                base_key,
                side_key,
                cols_to_add,
                cols_to_drop,
                output_path,
                output_numeric_format,
            ),
            daemon=True,
        ).start()

    def _run_enrich_thread(
        self,
        base_path,
        side_path,
        base_key,
        side_key,
        cols_to_add,
        cols_to_drop,
        output_path,
        output_numeric_format,
    ):
        success = logic.enrich_file(
            base_path,
            side_path,
            base_key,
            side_key,
            cols_to_add,
            cols_to_drop,
            output_path,
            output_numeric_format,
        )
        if success:
            messagebox.showinfo("Exito", f"Archivo enriquecido guardado correctamente en:\n{output_path}")
        else:
            messagebox.showerror("Error", "Fallo el proceso de enriquecimiento. Revise el registro.")

    def run_stage1(self):
        input_dir = self.stage1_input_dir.get()
        output_dir = self.stage1_output_dir.get()
        output_numeric_format = self.output_numeric_format_var.get()

        if not os.path.isdir(input_dir) or not os.path.isdir(output_dir):
            messagebox.showerror(
                "Error de Ruta",
                "Seleccione una carpeta de entrada y de destino validas para la Etapa 1.",
            )
            return

        print("\n--- INICIANDO PROCESO DE LOTES POR CARPETAS ---")
        threading.Thread(
            target=self._process_stage1_thread,
            args=(input_dir, output_dir, output_numeric_format),
            daemon=True,
        ).start()

    def _process_stage1_thread(self, input_dir, output_dir, output_numeric_format):
        processed_folders, success = logic.process_stage1_by_subfolders(
            input_dir,
            output_dir,
            output_numeric_format,
        )

        if self.clean_folder_var.get() and processed_folders:
            print("\n--- Limpiando subcarpetas procesadas ---")
            for folder_path in processed_folders:
                print(f"Limpiando carpeta: {os.path.basename(folder_path)}")
                try:
                    files_to_delete = glob.glob(os.path.join(folder_path, "*"))
                    for file_path in files_to_delete:
                        os.remove(file_path)
                    print(f"Carpeta '{os.path.basename(folder_path)}' limpiada.")
                except Exception as e:
                    print(f"Error general limpiando la carpeta '{os.path.basename(folder_path)}': {e}")

        if success:
            messagebox.showinfo("Exito", "Proceso de lotes finalizado. Revise el registro para mas detalles.")
        else:
            messagebox.showerror("Error", "No se proceso ningun lote con exito o no se encontraron datos.")

        print("--- Fin del proceso ---")

    def run_stage2(self):
        input_dir = self.stage2_input_dir.get()
        output_numeric_format = self.output_numeric_format_var.get()

        if not os.path.isdir(input_dir):
            messagebox.showerror("Error de Ruta", "Seleccione una carpeta valida de lotes procesados.")
            return

        filter_column = None
        allowed_values = None

        if self.stage2_enable_filter_var.get():
            filter_column = self.stage2_filter_column_var.get().strip()
            allowed_values = [value.strip() for value in self.stage2_filter_values if value.strip()]

            if not filter_column:
                messagebox.showerror("Error", "Seleccione una columna para aplicar el filtro.")
                return

            if not allowed_values:
                messagebox.showerror("Error", "Agregue al menos un valor permitido para el filtro.")
                return

        output_path = filedialog.asksaveasfilename(
            title="Guardar Archivo Unificado Como...",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            return

        print("\n--- INICIANDO UNIFICACION DE LOTES ---")
        threading.Thread(
            target=self._process_stage2_thread,
            args=(input_dir, output_path, filter_column, allowed_values, output_numeric_format),
            daemon=True,
        ).start()

    def _process_stage2_thread(self, input_dir, output_path, filter_column, allowed_values, output_numeric_format):
        success = logic.process_stage2_consolidation(
            input_dir,
            output_path,
            filter_column,
            allowed_values,
            output_numeric_format,
        )
        if success:
            messagebox.showinfo("Exito", f"Archivo unificado guardado en:\n{output_path}")
        else:
            messagebox.showerror("Error", "No se pudo generar el archivo unificado.")
        print("--- Fin del proceso ---")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
