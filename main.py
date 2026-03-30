import glob
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd

# Add src to python path
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from src import logic


CORPORATE_COLORS = {
    "navy": "#1B3F66",
    "blue": "#2E86AB",
    "green": "#2ECC71",
    "amber": "#F39C12",
    "bg": "#F4F7FB",
    "surface": "#FFFFFF",
    "surface_alt": "#EAF1F8",
    "border": "#C8D8E6",
    "text": "#17324D",
    "muted": "#5F7892",
}


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class ConsoleRedirector:
    """Redirect stdout/stderr to a text widget."""

    def __init__(self, text_widget):
        self.text_widget = text_widget

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
        self.root.title("Gestor de Unificacion")
        self.root.geometry("1280x860")
        self.root.minsize(1100, 760)
        self.root.configure(fg_color=CORPORATE_COLORS["bg"])
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        self.enrich_base_file = tk.StringVar()
        self.enrich_side_file = tk.StringVar()
        self.enrich_base_key_col = tk.StringVar()
        self.enrich_side_key_col = tk.StringVar()
        self.enrich_col_search_var = tk.StringVar()
        self.drop_col_search_var = tk.StringVar()
        self.enrich_base_headers = []
        self.enrich_side_headers = []
        self.enrich_selected_add_cols = []
        self.enrich_selected_drop_cols = []

        self.stage1_input_dir = tk.StringVar(value="")
        self.stage1_output_dir = tk.StringVar(value="")
        self.stage2_input_dir = tk.StringVar(value="")
        self.stage2_enable_filter_var = tk.BooleanVar(value=False)
        self.stage2_filter_column_var = tk.StringVar(value="Tipo")
        self.stage2_filter_value_var = tk.StringVar()
        self.stage2_filter_values = []
        self.stage2_available_headers = []
        self.output_numeric_format_var = tk.StringVar(value="excel")
        self.clean_folder_var = tk.BooleanVar(value=False)
        self.unified_enable_enrich_var = tk.BooleanVar(value=False)

        self._build_shell()
        self._sync_unified_base_headers()

        sys.stdout = ConsoleRedirector(self.console_text)
        sys.stderr = ConsoleRedirector(self.console_text)

        print("Sistema listo. Esperando instrucciones...")

    def _build_shell(self):
        self.content = ctk.CTkFrame(
            self.root,
            fg_color=CORPORATE_COLORS["bg"],
            corner_radius=0,
        )
        self.content.grid(row=0, column=0, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(1, weight=1)
        self.content.grid_rowconfigure(2, weight=0)

        self._build_header()
        self._build_tabs()
        self._build_console()

    def _build_header(self):
        header = ctk.CTkFrame(
            self.content,
            fg_color=CORPORATE_COLORS["surface"],
            corner_radius=14,
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
        )
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 8))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="Gestor Excel",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=CORPORATE_COLORS["text"],
        ).grid(row=0, column=0, sticky="w", padx=16, pady=12)

        ctk.CTkButton(
            header,
            text="Limpiar consola",
            command=self._clear_console,
            fg_color=CORPORATE_COLORS["surface_alt"],
            hover_color="#DCE9F4",
            text_color=CORPORATE_COLORS["navy"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            width=120,
            height=34,
            corner_radius=10,
        ).grid(row=0, column=1, sticky="e", padx=(0, 10))

        ctk.CTkButton(
            header,
            text="Salir",
            command=self.root.destroy,
            fg_color=CORPORATE_COLORS["navy"],
            hover_color=CORPORATE_COLORS["blue"],
            text_color="white",
            width=80,
            height=34,
            corner_radius=10,
        ).grid(row=0, column=2, sticky="e", padx=(0, 16))

    def _build_tabs(self):
        self.tabview = ctk.CTkTabview(
            self.content,
            fg_color=CORPORATE_COLORS["surface"],
            segmented_button_fg_color=CORPORATE_COLORS["surface_alt"],
            segmented_button_selected_color=CORPORATE_COLORS["navy"],
            segmented_button_selected_hover_color=CORPORATE_COLORS["blue"],
            segmented_button_unselected_color=CORPORATE_COLORS["surface_alt"],
            segmented_button_unselected_hover_color="#DCE9F4",
            text_color=CORPORATE_COLORS["text"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            corner_radius=24,
        )
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 10))

        main_tab = self.tabview.add("Proceso unificado")
        main_tab.grid_columnconfigure(0, weight=1)
        main_tab.grid_rowconfigure(0, weight=1)
        self._create_main_process_tab(main_tab)

    def _build_console(self):
        console_card = ctk.CTkFrame(
            self.content,
            fg_color=CORPORATE_COLORS["surface"],
            corner_radius=24,
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
        )
        console_card.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 24))
        console_card.grid_columnconfigure(0, weight=1)

        top_row = ctk.CTkFrame(console_card, fg_color="transparent")
        top_row.grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 10))
        top_row.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            top_row,
            text="Registro de actividad",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=CORPORATE_COLORS["text"],
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            top_row,
            text="Salida operativa en tiempo real",
            font=ctk.CTkFont(size=12),
            text_color=CORPORATE_COLORS["muted"],
        ).grid(row=1, column=0, sticky="w")

        self.console_text = ctk.CTkTextbox(
            console_card,
            height=120,
            fg_color="#0F2740",
            text_color="#D7E5F2",
            border_width=0,
            corner_radius=18,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word",
        )
        self.console_text.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 20))
        self.console_text.configure(state="disabled")

    def _create_main_process_tab(self, parent_tab):
        scroll = ctk.CTkScrollableFrame(parent_tab, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=8, pady=8)

        source_card = self._create_section_card(scroll, "1 - Configuracion")
        self._file_picker_row(
            source_card,
            "Carpeta de entrada",
            self.stage1_input_dir,
            lambda: self._select_dir(self.stage1_input_dir),
            "Seleccionar carpeta",
        )

        ctk.CTkCheckBox(
            source_card,
            text="Limpiar subcarpetas de entrada al terminar",
            variable=self.clean_folder_var,
            fg_color=CORPORATE_COLORS["green"],
            hover_color="#27AE60",
            text_color=CORPORATE_COLORS["text"],
        ).pack(anchor="w", padx=20, pady=(0, 10))

        filter_toggle = ctk.CTkFrame(source_card, fg_color="transparent")
        filter_toggle.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkCheckBox(
            filter_toggle,
            text="Usar filtro",
            variable=self.stage2_enable_filter_var,
            fg_color=CORPORATE_COLORS["amber"],
            hover_color="#D68910",
            text_color=CORPORATE_COLORS["text"],
            command=self._toggle_filter_section,
        ).pack(anchor="w")

        self.filter_card = ctk.CTkFrame(
            source_card,
            fg_color=CORPORATE_COLORS["surface_alt"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            corner_radius=14,
        )
        self.filter_card.pack(fill="x", padx=20, pady=(0, 12))

        top_filter_row = ctk.CTkFrame(self.filter_card, fg_color="transparent")
        top_filter_row.pack(fill="x", padx=16, pady=(0, 10))
        top_filter_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            top_filter_row,
            text="Columna a filtrar",
            text_color=CORPORATE_COLORS["text"],
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=(0, 10))

        self.stage2_filter_column_combo = ctk.CTkComboBox(
            top_filter_row,
            variable=self.stage2_filter_column_var,
            values=[],
            button_color=CORPORATE_COLORS["blue"],
            button_hover_color=CORPORATE_COLORS["navy"],
            border_color=CORPORATE_COLORS["border"],
            fg_color=CORPORATE_COLORS["surface"],
            text_color=CORPORATE_COLORS["text"],
            dropdown_fg_color=CORPORATE_COLORS["surface"],
            dropdown_hover_color=CORPORATE_COLORS["surface_alt"],
            dropdown_text_color=CORPORATE_COLORS["text"],
        )
        self.stage2_filter_column_combo.grid(row=0, column=1, sticky="ew", padx=(0, 10))

        ctk.CTkButton(
            top_filter_row,
            text="Recargar",
            command=self._load_stage2_headers,
            width=110,
            fg_color=CORPORATE_COLORS["surface"],
            text_color=CORPORATE_COLORS["navy"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            hover_color="#DCE9F4",
        ).grid(row=0, column=2, sticky="e")

        ctk.CTkLabel(
            self.filter_card,
            text="Valores permitidos",
            text_color=CORPORATE_COLORS["muted"],
            wraplength=860,
            justify="left",
        ).pack(anchor="w", padx=16)

        values_row = ctk.CTkFrame(self.filter_card, fg_color="transparent")
        values_row.pack(fill="x", padx=16, pady=(10, 10))
        values_row.grid_columnconfigure(0, weight=1)

        ctk.CTkEntry(
            values_row,
            textvariable=self.stage2_filter_value_var,
            placeholder_text="Escribe un valor y agregalo a la lista",
            border_color=CORPORATE_COLORS["border"],
            fg_color=CORPORATE_COLORS["surface"],
            text_color=CORPORATE_COLORS["text"],
        ).grid(row=0, column=0, sticky="ew", padx=(0, 10))

        ctk.CTkButton(
            values_row,
            text="Agregar valor",
            command=self._add_stage2_filter_value,
            fg_color=CORPORATE_COLORS["green"],
            hover_color="#27AE60",
            width=140,
        ).grid(row=0, column=1)

        self.stage2_filter_values_listbox = self._create_listbox(self.filter_card, height=4)
        self.stage2_filter_values_listbox.pack(fill="x", padx=16, pady=(0, 10))

        values_actions = ctk.CTkFrame(self.filter_card, fg_color="transparent")
        values_actions.pack(fill="x", padx=16, pady=(0, 16))

        ctk.CTkButton(
            values_actions,
            text="Quitar seleccionados",
            command=self._remove_stage2_filter_values,
            fg_color=CORPORATE_COLORS["surface"],
            text_color=CORPORATE_COLORS["navy"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            hover_color="#DCE9F4",
        ).pack(side="left")

        ctk.CTkButton(
            values_actions,
            text="Limpiar lista",
            command=self._clear_stage2_filter_values,
            fg_color=CORPORATE_COLORS["surface"],
            text_color=CORPORATE_COLORS["navy"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            hover_color="#DCE9F4",
        ).pack(side="left", padx=(10, 10))

        ctk.CTkButton(
            values_actions,
            text="Cargar Materiales",
            command=lambda: self._add_stage2_filter_value("Materiales"),
            fg_color=CORPORATE_COLORS["amber"],
            hover_color="#D68910",
        ).pack(side="left")

        enrich_card = self._create_section_card(scroll, "2 - Enriquecimiento")
        enrich_toggle = ctk.CTkFrame(enrich_card, fg_color="transparent")
        enrich_toggle.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkCheckBox(
            enrich_toggle,
            text="Usar enriquecimiento",
            variable=self.unified_enable_enrich_var,
            fg_color=CORPORATE_COLORS["green"],
            hover_color="#27AE60",
            text_color=CORPORATE_COLORS["text"],
            command=self._toggle_enrich_section,
        ).pack(anchor="w")

        self.enrich_details = ctk.CTkFrame(enrich_card, fg_color="transparent")

        enrich_files_card = ctk.CTkFrame(
            self.enrich_details,
            fg_color=CORPORATE_COLORS["surface_alt"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            corner_radius=14,
        )
        enrich_files_card.pack(fill="x", padx=20, pady=(0, 14))

        self._file_picker_row(
            enrich_files_card,
            "Archivo de enriquecimiento",
            self.enrich_side_file,
            self._select_enrich_file,
            "Seleccionar archivo",
        )

        keys_card = ctk.CTkFrame(
            self.enrich_details,
            fg_color=CORPORATE_COLORS["surface_alt"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            corner_radius=14,
        )
        keys_card.pack(fill="x", padx=20, pady=(0, 14))

        self.combo_base_key = self._create_labeled_combo(
            keys_card,
            "Columna clave del consolidado",
            self.enrich_base_key_col,
        )
        self.combo_side_key = self._create_labeled_combo(
            keys_card,
            "Columna clave del archivo de enriquecimiento",
            self.enrich_side_key_col,
        )

        add_card = ctk.CTkFrame(
            self.enrich_details,
            fg_color=CORPORATE_COLORS["surface_alt"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            corner_radius=14,
        )
        add_card.pack(fill="x", padx=20, pady=(0, 14))

        search_row = ctk.CTkFrame(add_card, fg_color="transparent")
        search_row.pack(fill="x", padx=16, pady=(14, 10))

        ctk.CTkLabel(
            search_row,
            text="Buscar columnas para agregar",
            text_color=CORPORATE_COLORS["text"],
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(anchor="w")

        search_entry = ctk.CTkEntry(
            search_row,
            textvariable=self.enrich_col_search_var,
            placeholder_text="Filtra por nombre de columna",
            border_color=CORPORATE_COLORS["border"],
            fg_color=CORPORATE_COLORS["surface"],
            text_color=CORPORATE_COLORS["text"],
        )
        search_entry.pack(fill="x", pady=(8, 0))
        search_entry.bind("<KeyRelease>", self._filter_cols_listbox)

        add_lists = ctk.CTkFrame(add_card, fg_color="transparent")
        add_lists.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        add_lists.grid_columnconfigure((0, 2), weight=1)
        add_lists.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(add_lists, text="Disponibles", text_color=CORPORATE_COLORS["text"]).grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        ctk.CTkLabel(add_lists, text="Se agregaran", text_color=CORPORATE_COLORS["text"]).grid(
            row=0, column=2, sticky="w", pady=(0, 6)
        )

        self.listbox_available_add = self._create_listbox(add_lists, height=6)
        self.listbox_available_add.grid(row=1, column=0, sticky="nsew")

        self.listbox_selected_add = self._create_listbox(add_lists, height=6)
        self.listbox_selected_add.grid(row=1, column=2, sticky="nsew")

        add_actions = ctk.CTkFrame(add_lists, fg_color="transparent")
        add_actions.grid(row=1, column=1, sticky="ns", padx=12)

        ctk.CTkButton(
            add_actions,
            text="Agregar >",
            command=self._add_selected_enrich_columns,
            fg_color=CORPORATE_COLORS["green"],
            hover_color="#27AE60",
            width=110,
        ).pack(pady=(30, 10))

        ctk.CTkButton(
            add_actions,
            text="< Quitar",
            command=self._remove_selected_enrich_columns,
            fg_color=CORPORATE_COLORS["surface"],
            text_color=CORPORATE_COLORS["navy"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            hover_color="#DCE9F4",
            width=110,
        ).pack()

        drop_card = ctk.CTkFrame(
            self.enrich_details,
            fg_color=CORPORATE_COLORS["surface_alt"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            corner_radius=14,
        )
        drop_card.pack(fill="x", padx=20, pady=(0, 14))

        drop_search_row = ctk.CTkFrame(drop_card, fg_color="transparent")
        drop_search_row.pack(fill="x", padx=16, pady=(14, 10))

        ctk.CTkLabel(
            drop_search_row,
            text="Buscar columnas para quitar",
            text_color=CORPORATE_COLORS["text"],
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(anchor="w")

        drop_search_entry = ctk.CTkEntry(
            drop_search_row,
            textvariable=self.drop_col_search_var,
            placeholder_text="Filtra por nombre de columna",
            border_color=CORPORATE_COLORS["border"],
            fg_color=CORPORATE_COLORS["surface"],
            text_color=CORPORATE_COLORS["text"],
        )
        drop_search_entry.pack(fill="x", pady=(8, 0))
        drop_search_entry.bind("<KeyRelease>", self._filter_drop_cols_listbox)

        drop_lists = ctk.CTkFrame(drop_card, fg_color="transparent")
        drop_lists.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        drop_lists.grid_columnconfigure((0, 2), weight=1)
        drop_lists.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(drop_lists, text="Columnas del consolidado", text_color=CORPORATE_COLORS["text"]).grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        ctk.CTkLabel(drop_lists, text="Se quitaran", text_color=CORPORATE_COLORS["text"]).grid(
            row=0, column=2, sticky="w", pady=(0, 6)
        )

        self.listbox_available_drop = self._create_listbox(drop_lists, height=5)
        self.listbox_available_drop.grid(row=1, column=0, sticky="nsew")

        self.listbox_selected_drop = self._create_listbox(drop_lists, height=5)
        self.listbox_selected_drop.grid(row=1, column=2, sticky="nsew")

        drop_actions = ctk.CTkFrame(drop_lists, fg_color="transparent")
        drop_actions.grid(row=1, column=1, sticky="ns", padx=12)

        ctk.CTkButton(
            drop_actions,
            text="Quitar >",
            command=self._add_drop_columns,
            fg_color=CORPORATE_COLORS["amber"],
            hover_color="#D68910",
            width=110,
        ).pack(pady=(22, 10))

        ctk.CTkButton(
            drop_actions,
            text="< Dejar",
            command=self._remove_drop_columns,
            fg_color=CORPORATE_COLORS["surface"],
            text_color=CORPORATE_COLORS["navy"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            hover_color="#DCE9F4",
            width=110,
        ).pack()

        run_card = self._create_section_card(scroll, "3 - Ejecutar")
        self._create_output_format_selector(run_card, compact=True)
        ctk.CTkButton(
            run_card,
            text="Ejecutar proceso unificado",
            command=self.run_unified_process,
            fg_color=CORPORATE_COLORS["green"],
            hover_color="#27AE60",
            height=48,
            corner_radius=14,
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(fill="x", padx=20, pady=(0, 20))

        self._toggle_filter_section()
        self._toggle_enrich_section()

    def _create_section_card(self, parent, title):
        card = ctk.CTkFrame(
            parent,
            fg_color=CORPORATE_COLORS["surface"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            corner_radius=18,
        )
        card.pack(fill="x", padx=8, pady=(0, 10))

        ctk.CTkLabel(
            card,
            text=title,
            text_color=CORPORATE_COLORS["text"],
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(anchor="w", padx=18, pady=(14, 10))
        return card

    def _file_picker_row(self, parent, label, string_var, command, button_text):
        wrapper = ctk.CTkFrame(parent, fg_color="transparent")
        wrapper.pack(fill="x", padx=20, pady=(0, 12))
        wrapper.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            wrapper,
            text=label,
            text_color=CORPORATE_COLORS["text"],
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        entry = ctk.CTkEntry(
            wrapper,
            textvariable=string_var,
            border_color=CORPORATE_COLORS["border"],
            fg_color=CORPORATE_COLORS["surface_alt"],
            text_color=CORPORATE_COLORS["text"],
            height=40,
        )
        entry.grid(row=1, column=0, sticky="ew", padx=(0, 12))

        ctk.CTkButton(
            wrapper,
            text=button_text,
            command=command,
            fg_color=CORPORATE_COLORS["surface"],
            text_color=CORPORATE_COLORS["navy"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            hover_color="#DCE9F4",
            width=170,
            height=40,
            corner_radius=12,
        ).grid(row=1, column=1)

    def _create_labeled_combo(self, parent, label, variable):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=20, pady=(0, 12))

        ctk.CTkLabel(
            row,
            text=label,
            text_color=CORPORATE_COLORS["text"],
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(anchor="w", pady=(0, 8))

        combo = ctk.CTkComboBox(
            row,
            variable=variable,
            values=[],
            button_color=CORPORATE_COLORS["blue"],
            button_hover_color=CORPORATE_COLORS["navy"],
            border_color=CORPORATE_COLORS["border"],
            fg_color=CORPORATE_COLORS["surface_alt"],
            text_color=CORPORATE_COLORS["text"],
            dropdown_fg_color=CORPORATE_COLORS["surface"],
            dropdown_hover_color=CORPORATE_COLORS["surface_alt"],
            dropdown_text_color=CORPORATE_COLORS["text"],
            height=40,
        )
        combo.pack(fill="x")
        return combo

    def _create_output_format_selector(self, parent, compact=False):
        card = ctk.CTkFrame(
            parent,
            fg_color=CORPORATE_COLORS["surface_alt"],
            border_width=1,
            border_color=CORPORATE_COLORS["border"],
            corner_radius=14,
        )
        card.pack(fill="x", padx=20, pady=(0, 14))

        radio_row = ctk.CTkFrame(card, fg_color="transparent")
        if compact:
            radio_row.pack(fill="x", padx=12, pady=10)
            ctk.CTkLabel(
                radio_row,
                text="Formato",
                text_color=CORPORATE_COLORS["text"],
                font=ctk.CTkFont(size=13, weight="bold"),
            ).pack(side="left", padx=(4, 16))
        else:
            ctk.CTkLabel(
                card,
                text="Formato final de columnas numericas",
                text_color=CORPORATE_COLORS["text"],
                font=ctk.CTkFont(size=13, weight="bold"),
            ).pack(anchor="w", padx=16, pady=(14, 6))
            radio_row.pack(anchor="w", padx=12, pady=(0, 12))

        ctk.CTkRadioButton(
            radio_row,
            text="Numero Excel",
            value="excel",
            variable=self.output_numeric_format_var,
            fg_color=CORPORATE_COLORS["blue"],
            hover_color=CORPORATE_COLORS["navy"],
            text_color=CORPORATE_COLORS["text"],
        ).pack(side="left", padx=(0, 16))

        ctk.CTkRadioButton(
            radio_row,
            text="Texto con coma",
            value="comma_text",
            variable=self.output_numeric_format_var,
            fg_color=CORPORATE_COLORS["green"],
            hover_color="#27AE60",
            text_color=CORPORATE_COLORS["text"],
        ).pack(side="left")

    def _toggle_filter_section(self):
        if getattr(self, "filter_card", None) is None:
            return
        if self.stage2_enable_filter_var.get():
            self.filter_card.pack(fill="x", padx=20, pady=(0, 12))
        else:
            self.filter_card.pack_forget()

    def _toggle_enrich_section(self):
        if getattr(self, "enrich_details", None) is None:
            return
        if self.unified_enable_enrich_var.get():
            self.enrich_details.pack(fill="x")
        else:
            self.enrich_details.pack_forget()

    def _create_listbox(self, parent, height):
        return tk.Listbox(
            parent,
            height=height,
            selectmode=tk.EXTENDED,
            exportselection=False,
            relief="flat",
            activestyle="none",
            bg=CORPORATE_COLORS["surface_alt"],
            fg=CORPORATE_COLORS["text"],
            selectbackground=CORPORATE_COLORS["blue"],
            selectforeground="white",
            highlightthickness=1,
            highlightbackground=CORPORATE_COLORS["border"],
            highlightcolor=CORPORATE_COLORS["blue"],
            font=("Segoe UI", 11),
        )

    def _clear_console(self):
        self.console_text.configure(state="normal")
        self.console_text.delete("1.0", "end")
        self.console_text.configure(state="disabled")

    def _show_message(self, level, title, message):
        def _callback():
            if level == "info":
                messagebox.showinfo(title, message)
            elif level == "warning":
                messagebox.showwarning(title, message)
            else:
                messagebox.showerror(title, message)

        self.root.after(0, _callback)

    def _sync_unified_base_headers(self):
        consolidated_headers = [
            "Tipo_Obra",
            "LCL_Origen",
            "Tipo",
            "Contador",
            "Mat./Prest.",
            "Descripcion Mat./Serv.",
            "Cantidad",
            "Unidad medida base",
            "Imputacion",
            "Precio unitario eD",
            "Precio unitario cliente",
            "Precio total eD",
            "Precio total cliente",
        ]
        self.enrich_base_headers = consolidated_headers
        if hasattr(self, "combo_base_key"):
            self.combo_base_key.configure(values=consolidated_headers)
            if self.enrich_base_key_col.get() not in consolidated_headers:
                self.enrich_base_key_col.set("LCL_Origen")
        if hasattr(self, "listbox_available_drop"):
            self.enrich_selected_drop_cols = [
                col for col in self.enrich_selected_drop_cols if col in self.enrich_base_headers
            ]
            self._refresh_drop_columns_ui()

    def _select_dir(self, string_var):
        dir_path = filedialog.askdirectory(title="Seleccionar Carpeta")
        if dir_path:
            string_var.set(dir_path)
            if string_var is self.stage1_input_dir:
                self._sync_unified_base_headers()

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
            self._show_message(
                "warning",
                "Error de Lectura",
                f"No se pudieron leer las columnas del archivo: {os.path.basename(file_path)}",
            )
        return []

    def _load_stage2_headers(self):
        input_dir = self.stage2_input_dir.get()
        self.stage2_available_headers = []
        self.stage2_filter_column_combo.configure(values=[""])

        if not os.path.isdir(input_dir):
            return

        input_files = sorted(glob.glob(os.path.join(input_dir, "*.xlsx")))
        if not input_files:
            print("No se encontraron archivos .xlsx para detectar columnas en la etapa 2.")
            return

        headers = self._get_file_headers(input_files[0])
        self.stage2_available_headers = headers
        self.stage2_filter_column_combo.configure(values=headers or [""])
        enriched_headers = ["Tipo_Obra"] + headers if headers else ["Tipo_Obra"]
        self.enrich_base_headers = enriched_headers
        if hasattr(self, "combo_base_key"):
            self.combo_base_key.configure(values=enriched_headers)
            if self.enrich_base_key_col.get() not in enriched_headers:
                if "LCL_Origen" in enriched_headers:
                    self.enrich_base_key_col.set("LCL_Origen")
                elif enriched_headers:
                    self.enrich_base_key_col.set(enriched_headers[0])
            self.enrich_selected_drop_cols = [
                col for col in self.enrich_selected_drop_cols if col in self.enrich_base_headers
            ]
            self._refresh_drop_columns_ui()

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
        self.combo_base_key.configure(values=headers or [""])
        self.enrich_base_key_col.set(headers[0] if headers else "")

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
        self.combo_side_key.configure(values=headers or [""])
        self.enrich_side_key_col.set(headers[0] if headers else "")

        self.enrich_selected_add_cols = [col for col in self.enrich_selected_add_cols if col in self.enrich_side_headers]
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
        search_term = self.drop_col_search_var.get().strip().lower()
        available = []
        for col in self.enrich_base_headers:
            if col in self.enrich_selected_drop_cols:
                continue
            if search_term and search_term not in col.lower():
                continue
            available.append(col)
        self._populate_listbox(self.listbox_available_drop, available)
        self._populate_listbox(self.listbox_selected_drop, self.enrich_selected_drop_cols)

    def _filter_cols_listbox(self, event=None):
        self._refresh_enrich_columns_ui()

    def _filter_drop_cols_listbox(self, event=None):
        self._refresh_drop_columns_ui()

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
            self._show_message(
                "error",
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
            self._show_message("info", "Exito", f"Archivo enriquecido guardado correctamente en:\n{output_path}")
        else:
            self._show_message("error", "Error", "Fallo el proceso de enriquecimiento. Revise el registro.")

    def run_unified_process(self):
        input_dir = self.stage1_input_dir.get().strip()
        output_numeric_format = self.output_numeric_format_var.get()

        if not os.path.isdir(input_dir):
            self._show_message("error", "Error de Ruta", "Seleccione una carpeta valida de entrada para el proceso unificado.")
            return

        filter_column = None
        allowed_values = None
        if self.stage2_enable_filter_var.get():
            filter_column = self.stage2_filter_column_var.get().strip()
            allowed_values = [value.strip() for value in self.stage2_filter_values if value.strip()]

            if not filter_column:
                self._show_message("error", "Error", "Seleccione una columna para aplicar el filtro.")
                return

            if not allowed_values:
                self._show_message("error", "Error", "Agregue al menos un valor permitido para el filtro.")
                return

        enrich_config = None
        if self.unified_enable_enrich_var.get():
            side_path = self.enrich_side_file.get().strip()
            base_key = self.enrich_base_key_col.get().strip()
            side_key = self.enrich_side_key_col.get().strip()
            cols_to_add = list(self.enrich_selected_add_cols)
            cols_to_drop = list(self.enrich_selected_drop_cols)

            if not side_path or not os.path.isfile(side_path):
                self._show_message("error", "Error", "Active el enriquecimiento solo si ya selecciono un archivo de enriquecimiento valido.")
                return

            if not base_key or not side_key or not cols_to_add:
                self._show_message(
                    "error",
                    "Error",
                    "Complete la configuracion de la pestana de cruce: clave base, clave de enriquecimiento y al menos una columna a agregar.",
                )
                return

            enrich_config = {
                "enabled": True,
                "side_path": side_path,
                "base_key": base_key,
                "side_key": side_key,
                "cols_to_add": cols_to_add,
                "cols_to_drop": cols_to_drop,
            }

        output_path = filedialog.asksaveasfilename(
            title="Guardar Archivo Final Como...",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            return

        print("\n--- INICIANDO PROCESO UNIFICADO ---")
        threading.Thread(
            target=self._run_unified_thread,
            args=(input_dir, output_path, filter_column, allowed_values, output_numeric_format, enrich_config),
            daemon=True,
        ).start()

    def _run_unified_thread(
        self,
        input_dir,
        output_path,
        filter_column,
        allowed_values,
        output_numeric_format,
        enrich_config,
    ):
        success = logic.run_unified_process(
            input_dir,
            output_path,
            output_numeric_format=output_numeric_format,
            filter_column=filter_column,
            allowed_values=allowed_values,
            enrich_config=enrich_config,
        )
        if success:
            self._show_message("info", "Exito", f"Proceso unificado completado. Archivo guardado en:\n{output_path}")
        else:
            self._show_message("error", "Error", "No se pudo completar el proceso unificado.")
        print("--- Fin del proceso unificado ---")

    def run_stage1(self):
        input_dir = self.stage1_input_dir.get()
        output_dir = self.stage1_output_dir.get()
        output_numeric_format = self.output_numeric_format_var.get()

        if not os.path.isdir(input_dir) or not os.path.isdir(output_dir):
            self._show_message(
                "error",
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
        processed_folders, success = logic.process_stage1_by_subfolders(input_dir, output_dir, output_numeric_format)

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
            self._show_message("info", "Exito", "Proceso de lotes finalizado. Revise el registro para mas detalles.")
        else:
            self._show_message("error", "Error", "No se proceso ningun lote con exito o no se encontraron datos.")

        print("--- Fin del proceso ---")

    def run_stage2(self):
        input_dir = self.stage2_input_dir.get()
        output_numeric_format = self.output_numeric_format_var.get()

        if not os.path.isdir(input_dir):
            self._show_message("error", "Error de Ruta", "Seleccione una carpeta valida de lotes procesados.")
            return

        filter_column = None
        allowed_values = None

        if self.stage2_enable_filter_var.get():
            filter_column = self.stage2_filter_column_var.get().strip()
            allowed_values = [value.strip() for value in self.stage2_filter_values if value.strip()]

            if not filter_column:
                self._show_message("error", "Error", "Seleccione una columna para aplicar el filtro.")
                return

            if not allowed_values:
                self._show_message("error", "Error", "Agregue al menos un valor permitido para el filtro.")
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
            self._show_message("info", "Exito", f"Archivo unificado guardado en:\n{output_path}")
        else:
            self._show_message("error", "Error", "No se pudo generar el archivo unificado.")
        print("--- Fin del proceso ---")


if __name__ == "__main__":
    root = ctk.CTk()
    app = App(root)
    root.mainloop()
