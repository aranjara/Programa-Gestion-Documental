import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from pathlib import Path

from app.security import validate_new_password
from app.services import (
    initialize_database, authenticate, change_password, verify_current_password,
    list_users, create_user, update_user_status, reset_user_password,
    create_record, update_record, delete_record, list_records, get_record, export_records,
    get_audit_rows, import_excel_data, get_fields, get_visible_fields,
    create_field, update_field, delete_field,
    get_fuid_header_config, save_fuid_header_config,
    get_fuid_detail_mapping, update_fuid_detail_mapping,
    generate_fuid,
    get_rotulo_carpeta_config, save_rotulo_carpeta_config,
    get_rotulo_carpeta_mapping, update_rotulo_carpeta_mapping,
    generate_rotulo_carpeta,
    get_rotulo_caja_config, save_rotulo_caja_config,
    get_rotulo_caja_mapping, update_rotulo_caja_mapping,
    generate_rotulo_caja
)

APP_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = APP_DIR / "templates" / "FUID.xlsx"



def _normalize_role_value(role_text):
    role_text = str(role_text or "").strip().lower()
    if role_text in ["normal", "usuario", "user"]:
        return "normal"
    if role_text in ["admin", "administrador"]:
        return "admin"
    return "normal"

class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.login_window = self
        self.root.title("Archivo Documental - Login")
        self.root.geometry("420x240")
        self.root.resizable(False, False)

        frm = ttk.Frame(root, padding=20)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Sistema de Archivo Documental", font=("Segoe UI", 14, "bold")).pack(pady=(0, 15))

        ttk.Label(frm, text="Usuario").pack(anchor="w")
        self.username_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.username_var).pack(fill="x", pady=(0, 10))

        ttk.Label(frm, text="Clave").pack(anchor="w")
        self.password_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.password_var, show="*").pack(fill="x", pady=(0, 15))

        ttk.Button(frm, text="Ingresar", command=self.login).pack(fill="x", ipady=6)
        self.root.bind("<Return>", lambda _: self.login())

    def reset_form(self):
        self.username_var.set("")
        self.password_var.set("")
        self.root.focus_force()

    def login(self):
        user, err = authenticate(self.username_var.get().strip(), self.password_var.get().strip())
        if err:
            messagebox.showerror("Error", err)
            return
        self.root.withdraw()
        if int(user["must_change_password"]) == 1:
            PasswordChangeWindow(self.root, user, True)
        else:
            MainWindow(self.root, user)


class PasswordChangeWindow(tk.Toplevel):
    def __init__(self, master, user, first_login=False):
        super().__init__(master)
        self.master = master
        self.user = user
        self.first_login = first_login
        self.title("Cambiar clave")
        self.geometry("420x220")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.handle_close)

        frm = ttk.Frame(self, padding=20)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Nueva clave").pack(anchor="w")
        self.p1 = tk.StringVar()
        ttk.Entry(frm, textvariable=self.p1, show="*").pack(fill="x", pady=(0, 10))

        ttk.Label(frm, text="Confirmar clave").pack(anchor="w")
        self.p2 = tk.StringVar()
        ttk.Entry(frm, textvariable=self.p2, show="*").pack(fill="x", pady=(0, 15))

        ttk.Button(frm, text="Guardar", command=self.save).pack(fill="x")

    def save(self):
        ok, msg = validate_new_password(self.p1.get().strip())
        if not ok:
            messagebox.showerror("Error", msg)
            return
        if self.p1.get().strip() != self.p2.get().strip():
            messagebox.showerror("Error", "Las claves no coinciden.")
            return
        change_password(self.user["id"], self.p1.get().strip())
        messagebox.showinfo("Listo", "Clave actualizada correctamente.")
        self.destroy()
        MainWindow(self.master, self.user | {"must_change_password": 0})

    def handle_close(self):
        if self.first_login:
            messagebox.showwarning("Aviso", "Debe cambiar la clave para ingresar.")
        else:
            self.destroy()


class SessionPasswordDialog(tk.Toplevel):
    def __init__(self, parent, user):
        super().__init__(parent)
        self.user = user
        self.title("Cambiar contraseña")
        self.geometry("420x250")
        self.resizable(False, False)
        self.grab_set()

        frm = ttk.Frame(self, padding=20)
        frm.pack(fill="both", expand=True)

        self.cur = tk.StringVar()
        self.p1 = tk.StringVar()
        self.p2 = tk.StringVar()

        ttk.Label(frm, text="Clave actual").pack(anchor="w")
        ttk.Entry(frm, textvariable=self.cur, show="*").pack(fill="x", pady=(0, 10))
        ttk.Label(frm, text="Nueva clave").pack(anchor="w")
        ttk.Entry(frm, textvariable=self.p1, show="*").pack(fill="x", pady=(0, 10))
        ttk.Label(frm, text="Confirmar nueva clave").pack(anchor="w")
        ttk.Entry(frm, textvariable=self.p2, show="*").pack(fill="x", pady=(0, 15))
        ttk.Button(frm, text="Guardar", command=self.save).pack(fill="x")

    def save(self):
        if not verify_current_password(self.user["id"], self.cur.get().strip()):
            messagebox.showerror("Error", "La clave actual no es correcta.")
            return
        ok, msg = validate_new_password(self.p1.get().strip())
        if not ok:
            messagebox.showerror("Error", msg)
            return
        if self.p1.get().strip() != self.p2.get().strip():
            messagebox.showerror("Error", "Las claves no coinciden.")
            return
        change_password(self.user["id"], self.p1.get().strip())
        messagebox.showinfo("Listo", "Contraseña actualizada correctamente.")
        self.destroy()


class DualScrollableForm(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.v_scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.h_scroll = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.inner = ttk.Frame(self.canvas)

        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        self.v_scroll.pack(side="right", fill="y")
        self.h_scroll.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._update)
        self.canvas.bind("<Configure>", self._resize)

    def _update(self, _=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _resize(self, event):
        self.canvas.itemconfig(self.window_id, width=max(event.width, 700))


class FieldDialog(tk.Toplevel):
    def __init__(self, parent, title, initial=None):
        super().__init__(parent)
        self.result = None
        self.title(title)
        self.geometry("520x320")
        self.resizable(False, False)
        self.grab_set()

        initial = initial or {}

        frm = ttk.Frame(self, padding=15)
        frm.pack(fill="both", expand=True)

        self.column_name = tk.StringVar(value=initial.get("column_name", ""))
        self.display_name = tk.StringVar(value=initial.get("display_name", ""))
        self.visible = tk.StringVar(value="Sí" if int(initial.get("visible", 1)) == 1 else "No")
        self.order = tk.StringVar(value=str(initial.get("display_order", "")))
        self.default = tk.StringVar(value=initial.get("default_value", ""))

        ttk.Label(frm, text="Nombre Columna").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.column_name, width=45).grid(row=0, column=1, sticky="w", pady=6)

        ttk.Label(frm, text="Etiqueta").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.display_name, width=45).grid(row=1, column=1, sticky="w", pady=6)

        ttk.Label(frm, text="Visible").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Combobox(frm, textvariable=self.visible, values=["Sí", "No"], state="readonly", width=10).grid(row=2, column=1, sticky="w", pady=6)

        ttk.Label(frm, text="Orden").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.order, width=12).grid(row=3, column=1, sticky="w", pady=6)

        ttk.Label(frm, text="Valor fijo por defecto").grid(row=4, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.default, width=45).grid(row=4, column=1, sticky="w", pady=6)

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(15, 0))
        ttk.Button(btns, text="Guardar", command=self.on_save).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="left")

    def on_save(self):
        try:
            order = int(self.order.get().strip())
        except ValueError:
            messagebox.showerror("Error", "El orden debe ser numérico.")
            return

        col = self.column_name.get().strip()
        if not col:
            messagebox.showerror("Error", "El Nombre Columna es obligatorio.")
            return

        self.result = {
            "column_name": col,
            "display_name": self.display_name.get().strip() or col,
            "visible": 1 if self.visible.get() == "Sí" else 0,
            "display_order": order,
            "default_value": self.default.get()
        }
        self.destroy()


class RecordDialog(tk.Toplevel):
    def __init__(self, parent, title, values=None, record_id=None):
        super().__init__(parent)
        self.parent = parent
        self.user = parent.user
        self.record_id = record_id
        self.result_saved = False

        self.title(title)
        self.geometry("760x620")
        self.resizable(True, True)
        self.grab_set()

        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)

        self.scrollable_form = DualScrollableForm(outer)
        self.scrollable_form.pack(fill="both", expand=True)

        self.form_vars = {}
        self.values = values or {}

        self.render_form()

        btns = ttk.Frame(outer)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Guardar", command=self.save).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="left")

    def render_form(self):
        for child in self.scrollable_form.inner.winfo_children():
            child.destroy()

        fields = get_fields()
        if not fields:
            ttk.Label(
                self.scrollable_form.inner,
                text="No hay campos creados. Primero crea campos o importa un Excel.",
                wraplength=500,
                justify="left"
            ).grid(row=0, column=0, sticky="w", padx=10, pady=10)
            self.scrollable_form._update()
            return

        for i, field in enumerate(fields):
            ttk.Label(self.scrollable_form.inner, text=field["column_name"]).grid(row=i, column=0, sticky="w", padx=(0, 10), pady=5)

            var = tk.StringVar(value=self.values.get(field["column_name"], field.get("default_value", "")))
            self.form_vars[field["column_name"]] = var

            ttk.Entry(
                self.scrollable_form.inner,
                textvariable=var,
                width=70
            ).grid(row=i, column=1, sticky="w", pady=5)

        self.scrollable_form.inner.columnconfigure(1, weight=1)
        self.scrollable_form._update()

    def save(self):
        if not self.form_vars:
            messagebox.showwarning("Aviso", "No hay campos configurados para guardar.")
            return

        payload = {k: v.get().strip() for k, v in self.form_vars.items()}

        if self.record_id:
            update_record(self.user["username"], self.record_id, payload)
            messagebox.showinfo("Listo", "Registro actualizado correctamente.")
        else:
            self.record_id = create_record(self.user["username"], payload)
            messagebox.showinfo("Listo", "Registro creado correctamente.")

        self.result_saved = True
        self.destroy()


class MainWindow(tk.Toplevel):
    def __init__(self, master, user):
        super().__init__(master)
        self.master = master
        self.user = user
        self.record_ids = []
        self.field_id = None
        self.column_widths = {}
        self.fuid_mapping_id = None
        self.fuid_extra_header_fields = []
        self.rotulo_carpeta_mapping_id = None

        self.title(f"Archivo Documental - {user['username']} ({user['role']})")
        self.geometry("1560x880")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.build_top()
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        self.records_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.records_tab, text="Registros")
        self.build_records_tab()

        self.import_tab = ttk.Frame(self.notebook)

# 🔐 SOLO ADMIN ve la pestaña

        if self.user.get("role") == "admin":
           self.notebook.add(self.import_tab, text="Importar Excel")
        self.build_import_tab()

        self.fields_tab = ttk.Frame(self.notebook)
        if self.user.get("role") == "admin":
            self.notebook.add(self.fields_tab, text="Campos")
        self.build_fields_tab()

        self.fuid_config_tab = ttk.Frame(self.notebook)
        if self.user.get("role") == "admin":
            self.notebook.add(self.fuid_config_tab, text="Configuración FUID")
        self.build_fuid_config_tab()

        self.generate_fuid_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.generate_fuid_tab, text="Generar FUID")
        self.build_generate_fuid_tab()

        self.rotulo_carpeta_config_tab = ttk.Frame(self.notebook)
        if self.user.get("role") == "admin":
            self.notebook.add(self.rotulo_carpeta_config_tab, text="Configuración Rótulo Carpeta")
        self.build_rotulo_carpeta_config_tab()

        self.generate_rotulo_carpeta_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.generate_rotulo_carpeta_tab, text="Generar Rótulo Carpeta")
        self.build_generate_rotulo_carpeta_tab()

        self.rotulo_caja_config_tab = ttk.Frame(self.notebook)
        if self.user.get("role") == "admin":
            self.notebook.add(self.rotulo_caja_config_tab, text="Configuración Rótulo Caja")
        self.build_rotulo_caja_config_tab()

        self.generate_rotulo_caja_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.generate_rotulo_caja_tab, text="Generar Rótulo Caja")
        self.build_generate_rotulo_caja_tab()

        if self.user.get("role") == "admin":
            self.users_tab = ttk.Frame(self.notebook)
            self.notebook.add(self.users_tab, text="Usuarios")
            self.build_users_tab()

            self.audit_tab = ttk.Frame(self.notebook)
            self.notebook.add(self.audit_tab, text="Auditoría")
            self.build_audit_tab()

    def build_top(self):
        top = ttk.Frame(self, padding=(8, 6))
        top.pack(fill="x")
        ttk.Label(top, text=f"Usuario: {self.user['username']}").pack(side="left")
        ttk.Button(top, text="Cambiar contraseña", command=self.open_change_password).pack(side="right", padx=(8, 0))
        ttk.Button(top, text="Cerrar sesión", command=self.logout).pack(side="right")

    def open_change_password(self):
        SessionPasswordDialog(self, self.user)

    def logout(self):
        self.destroy()
        if hasattr(self.master, "login_window"):
            self.master.login_window.reset_form()
        self.master.deiconify()

    def on_close(self):
        self.destroy()
        self.master.destroy()

    # REGISTROS
    def build_records_tab(self):
        outer = ttk.Frame(self.records_tab, padding=10)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x", pady=(0, 10))

        self.search_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.search_var).pack(side="left", fill="x", expand=True, padx=(0, 8))

        for txt, cmd in [
            ("Buscar", self.refresh_records),
            ("Nuevo", self.new_record),
            ("Editar", self.edit_selected_record),
            ("Eliminar fila", self.delete_selected_record),
            ("Exportar", self.export_data),
        ]:
            ttk.Button(top, text=txt, command=cmd).pack(side="left", padx=(0, 8))

        tree_frame = ttk.Frame(outer)
        tree_frame.pack(fill="both", expand=True)

        self.records_tree = ttk.Treeview(tree_frame, show="headings")
        y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.records_tree.yview)
        x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.records_tree.xview)

        self.records_tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        self.records_tree.grid(row=0, column=0, sticky="nsew")
        y.grid(row=0, column=1, sticky="ns")
        x.grid(row=1, column=0, sticky="ew")

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.refresh_records()

    def get_current_column_widths(self):
        widths = {}
        try:
            for col in self.records_tree["columns"]:
                widths[col] = self.records_tree.column(col, "width")
        except Exception:
            pass
        return widths

    def refresh_records(self):
        previous_widths = self.get_current_column_widths()
        if previous_widths:
            self.column_widths.update(previous_widths)

        fields = get_visible_fields()

        for item in self.records_tree.get_children():
            self.records_tree.delete(item)

        cols = [f["column_name"] for f in fields]
        self.records_tree["columns"] = cols

        for col in cols:
            self.records_tree.heading(col, text=col)
            self.records_tree.column(col, width=self.column_widths.get(col, 170), anchor="center", stretch=False)

        self.record_ids = []
        for row in list_records(self.search_var.get()):
            self.record_ids.append(row["id"])
            self.records_tree.insert("", "end", values=[row.get(c, "") for c in cols])

    def get_selected_record_id(self):
        sel = self.records_tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Seleccione una fila.")
            return None
        idx = self.records_tree.index(sel[0])
        return self.record_ids[idx] if idx < len(self.record_ids) else None

    def new_record(self):
        dlg = RecordDialog(self, "Nuevo registro")
        self.wait_window(dlg)
        if dlg.result_saved:
            self.refresh_records()

    def edit_selected_record(self):
        rid = self.get_selected_record_id()
        if rid is None:
            return
        data = get_record(rid)
        if not data:
            messagebox.showerror("Error", "No se encontró el registro.")
            return
        dlg = RecordDialog(self, "Editar registro", values=data, record_id=rid)
        self.wait_window(dlg)
        if dlg.result_saved:
            self.refresh_records()

    def delete_selected_record(self):
        rid = self.get_selected_record_id()
        if rid is None:
            return
        if not messagebox.askyesno("Confirmar", "¿Deseas eliminar la fila seleccionada?"):
            return
        delete_record(self.user["username"], rid)
        self.refresh_records()
        messagebox.showinfo("Listo", "Fila eliminada correctamente.")

    def export_data(self):
        path = filedialog.asksaveasfilename(
            title="Exportar registros",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")]
        )
        if not path:
            return
        final_path = export_records(path)
        messagebox.showinfo("Listo", "Datos exportados en:\\n" + str(final_path))

    # IMPORTAR
    def build_import_tab(self):
        outer = ttk.Frame(self.import_tab, padding=20)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Migración inicial desde Excel", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 10))
        ttk.Label(
            outer,
            text="La hoja debe llamarse 'Datos'. Las columnas nuevas del Excel se crean automáticamente. "
                 "Las columnas existentes cargan sus datos sin borrar lo anterior.",
            wraplength=900,
            justify="left"
        ).pack(anchor="w", pady=(0, 15))

        ttk.Button(outer, text="Seleccionar archivo e importar", command=self.run_import).pack(anchor="w")

        self.import_result = tk.StringVar(value="")
        ttk.Label(outer, textvariable=self.import_result, foreground="blue").pack(anchor="w", pady=(15, 0))

    def run_import(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm")]
        )
        if not path:
            return
        if not messagebox.askyesno("Confirmar importación", "Se importarán los registros de la hoja 'Datos'.\\n\\n¿Deseas continuar?"):
            return
        try:
            result = import_excel_data(self.user["username"], path)
            txt = (
                f"Importados: {result['imported']} | Filas vacías: {result['skipped']} | "
                f"Campos creados: {result['created_fields']}"
            )
            self.import_result.set(txt)
            self.refresh_records()
            self.refresh_fields()
            self.refresh_fuid_mapping()
            self.refresh_fuid_field_options()
            messagebox.showinfo("Listo", txt)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # CAMPOS
    def build_fields_tab(self):
        outer = ttk.Frame(self.fields_tab, padding=10)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Campos parametrizables", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 10))

        tree_frame = ttk.Frame(outer)
        tree_frame.pack(fill="both", expand=True, pady=(0, 12))

        self.fields_tree = ttk.Treeview(
            tree_frame,
            columns=("nombre_columna", "etiqueta", "visible", "orden", "valor_fijo"),
            show="headings",
            height=10
        )

        headers = {
            "nombre_columna": "Nombre Columna",
            "etiqueta": "Etiqueta",
            "visible": "Visible",
            "orden": "Orden",
            "valor_fijo": "Valor fijo por defecto"
        }
        widths = {
            "nombre_columna": 260,
            "etiqueta": 260,
            "visible": 100,
            "orden": 90,
            "valor_fijo": 300
        }

        for col in headers:
            self.fields_tree.heading(col, text=headers[col])
            self.fields_tree.column(col, width=widths[col], anchor="center")

        y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.fields_tree.yview)
        x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.fields_tree.xview)
        self.fields_tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)

        self.fields_tree.grid(row=0, column=0, sticky="nsew")
        y.grid(row=0, column=1, sticky="ns")
        x.grid(row=1, column=0, sticky="ew")

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.fields_tree.bind("<<TreeviewSelect>>", self.load_field_into_editor)

        editor = ttk.LabelFrame(outer, text="Editar campo", padding=10)
        editor.pack(fill="x")

        self.field_id = None
        self.column_name_var = tk.StringVar()
        self.display_name_var = tk.StringVar()
        self.visible_var = tk.StringVar(value="Sí")
        self.order_var = tk.StringVar()
        self.default_var = tk.StringVar()

        ttk.Label(editor, text="Nombre Columna").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(editor, textvariable=self.column_name_var, width=35).grid(row=0, column=1, sticky="w", pady=4)
        ttk.Label(editor, text="Etiqueta").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(editor, textvariable=self.display_name_var, width=45).grid(row=1, column=1, sticky="w", pady=4)
        ttk.Label(editor, text="Visible").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Combobox(editor, textvariable=self.visible_var, values=["Sí", "No"], state="readonly", width=10).grid(row=2, column=1, sticky="w", pady=4)
        ttk.Label(editor, text="Orden").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(editor, textvariable=self.order_var, width=12).grid(row=3, column=1, sticky="w", pady=4)
        ttk.Label(editor, text="Valor fijo por defecto").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(editor, textvariable=self.default_var, width=60).grid(row=4, column=1, sticky="w", pady=4)

        btns = ttk.Frame(outer)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Nuevo campo", command=self.new_field_dialog).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Aplicar cambios al campo", command=self.apply_field_changes).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Eliminar campo", command=self.delete_selected_field).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Guardar configuración", command=self.save_fields_refresh).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Recargar", command=self.refresh_fields).pack(side="left")

        self.refresh_fields()

    def refresh_fields(self):
        for item in self.fields_tree.get_children():
            self.fields_tree.delete(item)
        for f in get_fields():
            self.fields_tree.insert(
                "",
                "end",
                iid=str(f["id"]),
                values=(
                    f["column_name"],
                    f["display_name"],
                    "Sí" if int(f["visible"]) == 1 else "No",
                    f["display_order"],
                    f["default_value"]
                )
            )

    def load_field_into_editor(self, _event=None):
        sel = self.fields_tree.selection()
        if not sel:
            return
        iid = sel[0]
        vals = self.fields_tree.item(iid)["values"]
        self.field_id = int(iid)
        self.column_name_var.set(vals[0])
        self.display_name_var.set(vals[1])
        self.visible_var.set(vals[2])
        self.order_var.set(str(vals[3]))
        self.default_var.set(vals[4] if len(vals) > 4 else "")

    def new_field_dialog(self):
        fields = get_fields()
        next_order = max([int(f["display_order"]) for f in fields], default=0) + 1
        dlg = FieldDialog(self, "Nuevo campo", {"display_order": next_order, "visible": 1})
        self.wait_window(dlg)
        if dlg.result:
            try:
                create_field(
                    self.user["username"],
                    dlg.result["column_name"],
                    dlg.result["display_name"],
                    dlg.result["visible"],
                    dlg.result["display_order"],
                    dlg.result["default_value"]
                )
                self.refresh_fields()
                self.refresh_records()
                self.refresh_fuid_field_options()
                messagebox.showinfo("Listo", "Campo creado correctamente.")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def apply_field_changes(self):
        if not self.field_id:
            messagebox.showwarning("Aviso", "Selecciona un campo para editar.")
            return
        try:
            order = int(self.order_var.get().strip())
        except ValueError:
            messagebox.showerror("Error", "El orden debe ser numérico.")
            return
        try:
            update_field(
                self.user["username"],
                self.field_id,
                self.column_name_var.get().strip(),
                self.display_name_var.get().strip(),
                1 if self.visible_var.get() == "Sí" else 0,
                order,
                self.default_var.get()
            )
            self.refresh_fields()
            self.refresh_records()
            self.refresh_fuid_field_options()
            messagebox.showinfo("Listo", "Campo actualizado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def save_fields_refresh(self):
        self.refresh_fields()
        self.refresh_records()
        self.refresh_fuid_field_options()
        messagebox.showinfo("Listo", "Configuración actualizada en Registros.")

    def delete_selected_field(self):
        if not self.field_id:
            messagebox.showwarning("Aviso", "Selecciona un campo para eliminar.")
            return
        if not messagebox.askyesno("Confirmar", "¿Deseas eliminar el campo seleccionado?"):
            return
        try:
            delete_field(self.user["username"], self.field_id)
            self.field_id = None
            self.column_name_var.set("")
            self.display_name_var.set("")
            self.visible_var.set("Sí")
            self.order_var.set("")
            self.default_var.set("")
            self.refresh_fields()
            self.refresh_records()
            self.refresh_fuid_field_options()
            messagebox.showinfo("Listo", "Campo eliminado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # CONFIG FUID
    def build_fuid_config_tab(self):
        outer = ttk.Frame(self.fuid_config_tab, padding=10)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x", pady=(0, 10))
        ttk.Button(top, text="Guardar encabezado FUID", command=self.save_fuid_header).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Nuevo campo encabezado", command=self.new_fuid_header_field).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Recargar", command=self.reload_fuid_config).pack(side="left")

        body = ttk.Frame(outer)
        body.pack(fill="both", expand=True)

        left = ttk.LabelFrame(body, text="Encabezado FUID", padding=10)
        left.pack(side="left", fill="y", padx=(0, 10))

        right = ttk.LabelFrame(body, text="Detalle FUID", padding=10)
        right.pack(side="left", fill="both", expand=True)

        self.fuid_header_vars = {}
        self.header_editor_frame = left
        self.render_fuid_header_fields()

        tree_frame = ttk.Frame(right)
        tree_frame.pack(fill="both", expand=True, pady=(0, 10))
        self.fuid_map_tree = ttk.Treeview(
            tree_frame,
            columns=("campo_fuid", "tipo", "origen"),
            show="headings",
            height=12
        )
        for col, txt, width in [
            ("campo_fuid", "Campo FUID", 250),
            ("tipo", "Tipo de origen", 120),
            ("origen", "Valor / Configuración", 350),
        ]:
            self.fuid_map_tree.heading(col, text=txt)
            self.fuid_map_tree.column(col, width=width, anchor="w")
        y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.fuid_map_tree.yview)
        x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.fuid_map_tree.xview)
        self.fuid_map_tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        self.fuid_map_tree.grid(row=0, column=0, sticky="nsew")
        y.grid(row=0, column=1, sticky="ns")
        x.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        self.fuid_map_tree.bind("<<TreeviewSelect>>", self.load_fuid_mapping)

        editor = ttk.LabelFrame(right, text="Editar mapeo", padding=10)
        editor.pack(fill="x")

        self.fuid_field_var = tk.StringVar()
        self.fuid_type_var = tk.StringVar(value="field")
        self.fuid_value_var = tk.StringVar()

        ttk.Label(editor, text="Campo FUID").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(editor, textvariable=self.fuid_field_var, state="readonly", width=35).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(editor, text="Tipo de origen").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(editor, textvariable=self.fuid_type_var, values=["field", "fixed", "template"], state="readonly", width=15).grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(editor, text="Valor / Configuración").grid(row=2, column=0, sticky="w", pady=4)
        self.fuid_value_combo = ttk.Combobox(editor, textvariable=self.fuid_value_var, width=60)
        self.fuid_value_combo.grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(
            editor,
            text="En template puedes usar algo como: {Serie} / {Subserie} / {Asunto}",
            foreground="gray"
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))

        btns = ttk.Frame(right)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Guardar mapeo", command=self.save_fuid_mapping).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Recargar mapeos", command=self.reload_fuid_config).pack(side="left")

        self.refresh_fuid_mapping()
        self.refresh_fuid_field_options()

    def render_fuid_header_fields(self):
        for child in self.header_editor_frame.winfo_children():
            child.destroy()

        self.fuid_header_vars = {}
        base_labels = [
            ("entidad_remitente", "Entidad remitente"),
            ("entidad_productora", "Entidad productora"),
            ("unidad_administrativa", "Unidad administrativa"),
            ("oficina_productora", "Oficina productora"),
            ("objeto", "Objeto"),
            ("anio", "Año"),
            ("mes", "Mes"),
            ("dia", "Día"),
            ("n_transferencia", "N. Transferencia"),
        ]
        header_data = get_fuid_header_config()

        all_labels = base_labels + self.fuid_extra_header_fields

        for i, (key, label) in enumerate(all_labels):
            ttk.Label(self.header_editor_frame, text=label).grid(row=i, column=0, sticky="w", pady=4)
            var = tk.StringVar(value=header_data.get(key, ""))
            self.fuid_header_vars[key] = var
            ttk.Entry(self.header_editor_frame, textvariable=var, width=35).grid(row=i, column=1, sticky="w", pady=4)

    def new_fuid_header_field(self):
        key = simpledialog.askstring("Nuevo campo encabezado", "Nombre interno del campo (sin espacios o usando guion bajo):", parent=self)
        if not key:
            return
        label = simpledialog.askstring("Nuevo campo encabezado", "Etiqueta visible del campo:", parent=self)
        if not label:
            return
        key = key.strip()
        label = label.strip()
        if key in self.fuid_header_vars:
            messagebox.showwarning("Aviso", "Ese campo ya existe en el encabezado.")
            return
        self.fuid_extra_header_fields.append((key, label))
        self.render_fuid_header_fields()
        messagebox.showinfo("Listo", "Campo extra de encabezado agregado en esta sesión.")

    def reload_fuid_config(self):
        self.render_fuid_header_fields()
        self.refresh_fuid_mapping()
        self.refresh_fuid_field_options()

    def save_fuid_header(self):
        data = {k: v.get().strip() for k, v in self.fuid_header_vars.items()}
        save_fuid_header_config(self.user["username"], data)
        messagebox.showinfo("Listo", "Encabezado FUID guardado correctamente.")

    def refresh_fuid_mapping(self):
        for item in self.fuid_map_tree.get_children():
            self.fuid_map_tree.delete(item)
        for row in get_fuid_detail_mapping():
            self.fuid_map_tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(row["fuid_field"], row["mapping_type"], row["mapping_value"])
            )

    def refresh_fuid_field_options(self):
        field_names = [f["column_name"] for f in get_fields()]
        self.fuid_value_combo["values"] = field_names

    def load_fuid_mapping(self, _event=None):
        sel = self.fuid_map_tree.selection()
        if not sel:
            return
        iid = sel[0]
        vals = self.fuid_map_tree.item(iid)["values"]
        self.fuid_mapping_id = int(iid)
        self.fuid_field_var.set(vals[0])
        self.fuid_type_var.set(vals[1])
        self.fuid_value_var.set(vals[2])

    def save_fuid_mapping(self):
        if not self.fuid_mapping_id:
            messagebox.showwarning("Aviso", "Selecciona una fila del detalle FUID.")
            return
        try:
            update_fuid_detail_mapping(
                self.user["username"],
                self.fuid_mapping_id,
                self.fuid_type_var.get(),
                self.fuid_value_var.get()
            )
            self.refresh_fuid_mapping()
            messagebox.showinfo("Listo", "Mapeo FUID actualizado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # GENERAR FUID
    def build_generate_fuid_tab(self):
        outer = ttk.Frame(self.generate_fuid_tab, padding=20)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Generar FUID", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 12))

        form = ttk.Frame(outer)
        form.pack(fill="x", pady=(0, 15))

        ttk.Label(form, text="Filtro de búsqueda (opcional)").grid(row=0, column=0, sticky="w", pady=6)
        self.fuid_search_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.fuid_search_var, width=50).grid(row=0, column=1, sticky="w", pady=6)

        ttk.Label(form, text="Formato de salida").grid(row=1, column=0, sticky="w", pady=6)
        self.fuid_format_var = tk.StringVar(value="Excel")
        ttk.Combobox(form, textvariable=self.fuid_format_var, values=["Excel", "PDF", "Word"], state="readonly", width=15).grid(row=1, column=1, sticky="w", pady=6)

        ttk.Label(form, text="Carpeta de salida").grid(row=2, column=0, sticky="w", pady=6)
        self.fuid_output_dir_var = tk.StringVar(value=str(APP_DIR / "salidas"))
        ttk.Entry(form, textvariable=self.fuid_output_dir_var, width=60).grid(row=2, column=1, sticky="w", pady=6)
        ttk.Button(form, text="Elegir carpeta", command=self.choose_output_dir).grid(row=2, column=2, sticky="w", padx=(8, 0), pady=6)

        ttk.Button(outer, text="Generar FUID", command=self.run_generate_fuid).pack(anchor="w")

        self.fuid_result_var = tk.StringVar(value="")
        ttk.Label(outer, textvariable=self.fuid_result_var, foreground="blue", wraplength=1000, justify="left").pack(anchor="w", pady=(15, 0))

    def choose_output_dir(self):
        folder = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if folder:
            self.fuid_output_dir_var.set(folder)

    def choose_template_file(self):
        file_path = filedialog.askopenfilename(title="Seleccionar plantilla FUID", filetypes=[("Excel", "*.xlsx")])
        if file_path:
            self.fuid_template_var.set(file_path)

    def run_generate_fuid(self):
        try:
            out = generate_fuid(
                self.user["username"],
                self.fuid_search_var.get().strip(),
                self.fuid_format_var.get(),
                self.fuid_output_dir_var.get().strip(),
                None
            )
            self.fuid_result_var.set(f"FUID generado correctamente:\\n{out}")
            messagebox.showinfo("Listo", "FUID generado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", str(e))


    # CONFIGURACIÓN RÓTULO CARPETA
    def build_rotulo_carpeta_config_tab(self):
        outer = ttk.Frame(self.rotulo_carpeta_config_tab, padding=10)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x", pady=(0, 10))
        ttk.Button(top, text="Guardar configuración", command=self.save_rotulo_carpeta_config).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Recargar", command=self.reload_rotulo_carpeta_config).pack(side="left")

        config_frame = ttk.LabelFrame(outer, text="Encabezado Rótulo Carpeta", padding=10)
        config_frame.pack(fill="x", pady=(0, 10))

        self.rotulo_carpeta_search_field_var = tk.StringVar()
        self.rotulo_carpeta_title_var = tk.StringVar()
        self.rotulo_carpeta_fondo_var = tk.StringVar()
        self.rotulo_carpeta_seccion_var = tk.StringVar()
        self.rotulo_carpeta_subseccion_var = tk.StringVar()
        self.rotulo_carpeta_serie_var = tk.StringVar()
        self.rotulo_carpeta_subserie_var = tk.StringVar()

        ttk.Label(config_frame, text="Título superior").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(config_frame, textvariable=self.rotulo_carpeta_title_var, width=55).grid(row=0, column=1, sticky="w", pady=3)

        ttk.Label(config_frame, text="Fondo").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(config_frame, textvariable=self.rotulo_carpeta_fondo_var, width=55).grid(row=1, column=1, sticky="w", pady=3)

        ttk.Label(config_frame, text="Sección").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Entry(config_frame, textvariable=self.rotulo_carpeta_seccion_var, width=55).grid(row=2, column=1, sticky="w", pady=3)

        ttk.Label(config_frame, text="Subsección").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Entry(config_frame, textvariable=self.rotulo_carpeta_subseccion_var, width=55).grid(row=3, column=1, sticky="w", pady=3)

        ttk.Label(config_frame, text="Serie").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Entry(config_frame, textvariable=self.rotulo_carpeta_serie_var, width=55).grid(row=4, column=1, sticky="w", pady=3)

        ttk.Label(config_frame, text="Subserie").grid(row=5, column=0, sticky="w", pady=3)
        ttk.Entry(config_frame, textvariable=self.rotulo_carpeta_subserie_var, width=55).grid(row=5, column=1, sticky="w", pady=3)

        ttk.Label(config_frame, text="Campo para buscar rango").grid(row=6, column=0, sticky="w", pady=3)
        self.rotulo_carpeta_search_combo = ttk.Combobox(config_frame, textvariable=self.rotulo_carpeta_search_field_var, width=52)
        self.rotulo_carpeta_search_combo.grid(row=6, column=1, sticky="w", pady=3)

        ttk.Label(
            config_frame,
            text="Estos campos son texto fijo del rótulo. Los demás datos se configuran en el mapeo inferior.",
            foreground="gray"
        ).grid(row=7, column=0, columnspan=2, sticky="w", pady=(6, 0))

        body = ttk.LabelFrame(outer, text="Mapeo de campos del Rótulo de Carpeta", padding=10)
        body.pack(fill="both", expand=True)

        tree_frame = ttk.Frame(body)
        tree_frame.pack(fill="both", expand=True, pady=(0, 10))

        self.rotulo_carpeta_tree = ttk.Treeview(
            tree_frame,
            columns=("campo", "tipo", "origen"),
            show="headings",
            height=10
        )
        for col, txt, width in [
            ("campo", "Campo Rótulo", 250),
            ("tipo", "Tipo de origen", 120),
            ("origen", "Valor / Configuración", 350),
        ]:
            self.rotulo_carpeta_tree.heading(col, text=txt)
            self.rotulo_carpeta_tree.column(col, width=width, anchor="w")

        y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.rotulo_carpeta_tree.yview)
        x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.rotulo_carpeta_tree.xview)
        self.rotulo_carpeta_tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        self.rotulo_carpeta_tree.grid(row=0, column=0, sticky="nsew")
        y.grid(row=0, column=1, sticky="ns")
        x.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        self.rotulo_carpeta_tree.bind("<<TreeviewSelect>>", self.load_rotulo_carpeta_mapping)

        editor = ttk.LabelFrame(body, text="Editar mapeo", padding=10)
        editor.pack(fill="x")

        self.rotulo_carpeta_field_var = tk.StringVar()
        self.rotulo_carpeta_type_var = tk.StringVar(value="field")
        self.rotulo_carpeta_value_var = tk.StringVar()

        ttk.Label(editor, text="Campo Rótulo").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(editor, textvariable=self.rotulo_carpeta_field_var, state="readonly", width=35).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(editor, text="Tipo de origen").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(editor, textvariable=self.rotulo_carpeta_type_var, values=["field", "fixed", "template"], state="readonly", width=15).grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(editor, text="Valor / Configuración").grid(row=2, column=0, sticky="w", pady=4)
        self.rotulo_carpeta_value_combo = ttk.Combobox(editor, textvariable=self.rotulo_carpeta_value_var, width=60)
        self.rotulo_carpeta_value_combo.grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(
            editor,
            text="En template puedes usar algo como: {Fecha inicial} - {Fecha final}",
            foreground="gray"
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))

        btns = ttk.Frame(body)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Guardar mapeo", command=self.save_rotulo_carpeta_mapping).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Recargar mapeos", command=self.reload_rotulo_carpeta_config).pack(side="left")

        self.reload_rotulo_carpeta_config()

    def reload_rotulo_carpeta_config(self):
        data = get_rotulo_carpeta_config()
        self.rotulo_carpeta_search_field_var.set(data.get("campo_busqueda", "Número de orden"))
        self.rotulo_carpeta_title_var.set(data.get("titulo", "ALCALDIA DE RIONEGRO"))
        self.rotulo_carpeta_fondo_var.set(data.get("fondo", "ALCALDÍA"))
        self.rotulo_carpeta_seccion_var.set(data.get("seccion", "SECRETARÍA DE HACIENDA"))
        self.rotulo_carpeta_subseccion_var.set(data.get("subseccion", "SUBSECRETARÍA DE TESORERÍA"))
        self.rotulo_carpeta_serie_var.set(data.get("serie", "COMPROBANTES CONTABLES"))
        self.rotulo_carpeta_subserie_var.set(data.get("subserie", "COMPROBANTES CONTABLES DE EGRESO"))

        field_names = [f["column_name"] for f in get_fields()]
        self.rotulo_carpeta_search_combo["values"] = field_names
        self.rotulo_carpeta_value_combo["values"] = field_names

        self.refresh_rotulo_carpeta_mapping()

    def save_rotulo_carpeta_config(self):
        save_rotulo_carpeta_config(
            self.user["username"],
            {
                "campo_busqueda": self.rotulo_carpeta_search_field_var.get().strip() or "Número de orden",
                "titulo": self.rotulo_carpeta_title_var.get().strip() or "ALCALDIA DE RIONEGRO",
                "fondo": self.rotulo_carpeta_fondo_var.get().strip(),
                "seccion": self.rotulo_carpeta_seccion_var.get().strip(),
                "subseccion": self.rotulo_carpeta_subseccion_var.get().strip(),
                "serie": self.rotulo_carpeta_serie_var.get().strip(),
                "subserie": self.rotulo_carpeta_subserie_var.get().strip(),
            }
        )
        messagebox.showinfo("Listo", "Configuración del Rótulo de Carpeta guardada correctamente.")

    def refresh_rotulo_carpeta_mapping(self):
        for item in self.rotulo_carpeta_tree.get_children():
            self.rotulo_carpeta_tree.delete(item)

        for row in get_rotulo_carpeta_mapping():
            self.rotulo_carpeta_tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(row["rotulo_field"], row["mapping_type"], row["mapping_value"])
            )

    def load_rotulo_carpeta_mapping(self, _event=None):
        sel = self.rotulo_carpeta_tree.selection()
        if not sel:
            return
        iid = sel[0]
        vals = self.rotulo_carpeta_tree.item(iid)["values"]
        self.rotulo_carpeta_mapping_id = int(iid)
        self.rotulo_carpeta_field_var.set(vals[0])
        self.rotulo_carpeta_type_var.set(vals[1])
        self.rotulo_carpeta_value_var.set(vals[2])

    def save_rotulo_carpeta_mapping(self):
        if not self.rotulo_carpeta_mapping_id:
            messagebox.showwarning("Aviso", "Selecciona una fila del mapeo del rótulo.")
            return

        try:
            update_rotulo_carpeta_mapping(
                self.user["username"],
                self.rotulo_carpeta_mapping_id,
                self.rotulo_carpeta_type_var.get(),
                self.rotulo_carpeta_value_var.get()
            )
            self.refresh_rotulo_carpeta_mapping()
            messagebox.showinfo("Listo", "Mapeo del Rótulo de Carpeta actualizado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # GENERAR RÓTULO CARPETA
    def build_generate_rotulo_carpeta_tab(self):
        outer = ttk.Frame(self.generate_rotulo_carpeta_tab, padding=20)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Generar Rótulo de Carpeta", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 12))

        form = ttk.Frame(outer)
        form.pack(fill="x", pady=(0, 15))

        ttk.Label(form, text="Desde").grid(row=0, column=0, sticky="w", pady=6)
        self.rotulo_carpeta_desde_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.rotulo_carpeta_desde_var, width=25).grid(row=0, column=1, sticky="w", pady=6)

        ttk.Label(form, text="Hasta").grid(row=1, column=0, sticky="w", pady=6)
        self.rotulo_carpeta_hasta_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.rotulo_carpeta_hasta_var, width=25).grid(row=1, column=1, sticky="w", pady=6)

        ttk.Label(form, text="Formato de salida").grid(row=2, column=0, sticky="w", pady=6)
        self.rotulo_carpeta_format_var = tk.StringVar(value="Excel")
        ttk.Combobox(
            form,
            textvariable=self.rotulo_carpeta_format_var,
            values=["Excel", "PDF", "Word"],
            state="readonly",
            width=15
        ).grid(row=2, column=1, sticky="w", pady=6)

        ttk.Label(form, text="Carpeta de salida").grid(row=3, column=0, sticky="w", pady=6)
        self.rotulo_carpeta_output_dir_var = tk.StringVar(value=str(APP_DIR / "salidas"))
        ttk.Entry(form, textvariable=self.rotulo_carpeta_output_dir_var, width=60).grid(row=3, column=1, sticky="w", pady=6)
        ttk.Button(form, text="Elegir carpeta", command=self.choose_rotulo_carpeta_output_dir).grid(row=3, column=2, sticky="w", padx=(8, 0), pady=6)

        ttk.Label(
            outer,
            text="El campo usado para buscar el rango se configura en 'Configuración Rótulo Carpeta'. Ejemplo: Desde 1 / Hasta 100.",
            foreground="gray"
        ).pack(anchor="w", pady=(0, 10))

        ttk.Button(outer, text="Generar Rótulo de Carpeta", command=self.run_generate_rotulo_carpeta).pack(anchor="w")

        self.rotulo_carpeta_result_var = tk.StringVar(value="")
        ttk.Label(
            outer,
            textvariable=self.rotulo_carpeta_result_var,
            foreground="blue",
            wraplength=1000,
            justify="left"
        ).pack(anchor="w", pady=(15, 0))

    def choose_rotulo_carpeta_output_dir(self):
        folder = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if folder:
            self.rotulo_carpeta_output_dir_var.set(folder)

    def run_generate_rotulo_carpeta(self):
        try:
            desde = self.rotulo_carpeta_desde_var.get().strip()
            hasta = self.rotulo_carpeta_hasta_var.get().strip() or desde
            if not desde:
                messagebox.showwarning("Aviso", "Digite el valor inicial en Desde.")
                return

            output_dir = self.rotulo_carpeta_output_dir_var.get().strip()
            if not output_dir:
                output_dir = str(APP_DIR / "salidas")
                self.rotulo_carpeta_output_dir_var.set(output_dir)

            out = generate_rotulo_carpeta(
                self.user["username"],
                desde,
                hasta,
                self.rotulo_carpeta_format_var.get(),
                output_dir
            )
            self.rotulo_carpeta_result_var.set(f"Rótulo de Carpeta generado correctamente:\\n{out}")
            messagebox.showinfo("Listo", "Rótulo de Carpeta generado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", str(e))


    # USUARIOS

    # CONFIGURACIÓN RÓTULO CAJA
    def build_rotulo_caja_config_tab(self):
        outer = ttk.Frame(self.rotulo_caja_config_tab, padding=10)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x", pady=(0, 10))
        ttk.Button(top, text="Guardar configuración", command=self.save_rotulo_caja_config).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Recargar", command=self.reload_rotulo_caja_config).pack(side="left")

        config_frame = ttk.LabelFrame(outer, text="Encabezado Rótulo Caja", padding=10)
        config_frame.pack(fill="x", pady=(0, 10))

        self.rotulo_caja_campo_busqueda_var = tk.StringVar()
        self.rotulo_caja_campo_caja_var = tk.StringVar()
        self.rotulo_caja_campo_unidad_var = tk.StringVar()
        self.rotulo_caja_campo_fecha_inicial_var = tk.StringVar()
        self.rotulo_caja_campo_fecha_final_var = tk.StringVar()
        self.rotulo_caja_titulo_var = tk.StringVar()
        self.rotulo_caja_dependencia_var = tk.StringVar()
        self.rotulo_caja_serie_var = tk.StringVar()
        self.rotulo_caja_subserie_var = tk.StringVar()
        self.rotulo_caja_texto_consecutivo_var = tk.StringVar()
        self.rotulo_caja_texto_correlativo_var = tk.StringVar()
        self.rotulo_caja_observaciones_var = tk.StringVar()

        labels = [
            ("Título superior", self.rotulo_caja_titulo_var),
            ("Dependencia", self.rotulo_caja_dependencia_var),
            ("Serie", self.rotulo_caja_serie_var),
            ("Subserie", self.rotulo_caja_subserie_var),
            ("Texto consecutivo", self.rotulo_caja_texto_consecutivo_var),
            ("Texto correlativo", self.rotulo_caja_texto_correlativo_var),
        ]

        for idx, (label, var) in enumerate(labels):
            ttk.Label(config_frame, text=label).grid(row=idx, column=0, sticky="w", pady=3)
            ttk.Entry(config_frame, textvariable=var, width=70).grid(row=idx, column=1, sticky="w", pady=3)

        ttk.Label(config_frame, text="Observaciones").grid(row=len(labels), column=0, sticky="w", pady=3)
        ttk.Entry(config_frame, textvariable=self.rotulo_caja_observaciones_var, width=70).grid(row=len(labels), column=1, sticky="w", pady=3)

        row = len(labels) + 1
        ttk.Label(config_frame, text="Campo No. orden").grid(row=row, column=0, sticky="w", pady=3)
        self.rotulo_caja_busqueda_combo = ttk.Combobox(config_frame, textvariable=self.rotulo_caja_campo_busqueda_var, width=67)
        self.rotulo_caja_busqueda_combo.grid(row=row, column=1, sticky="w", pady=3)

        ttk.Label(config_frame, text="Campo caja").grid(row=row+1, column=0, sticky="w", pady=3)
        self.rotulo_caja_caja_combo = ttk.Combobox(config_frame, textvariable=self.rotulo_caja_campo_caja_var, width=67)
        self.rotulo_caja_caja_combo.grid(row=row+1, column=1, sticky="w", pady=3)

        ttk.Label(config_frame, text="Campo unidad documental").grid(row=row+2, column=0, sticky="w", pady=3)
        self.rotulo_caja_unidad_combo = ttk.Combobox(config_frame, textvariable=self.rotulo_caja_campo_unidad_var, width=67)
        self.rotulo_caja_unidad_combo.grid(row=row+2, column=1, sticky="w", pady=3)

        ttk.Label(config_frame, text="Campo fecha inicial").grid(row=row+3, column=0, sticky="w", pady=3)
        self.rotulo_caja_fecha_ini_combo = ttk.Combobox(config_frame, textvariable=self.rotulo_caja_campo_fecha_inicial_var, width=67)
        self.rotulo_caja_fecha_ini_combo.grid(row=row+3, column=1, sticky="w", pady=3)

        ttk.Label(config_frame, text="Campo fecha final").grid(row=row+4, column=0, sticky="w", pady=3)
        self.rotulo_caja_fecha_fin_combo = ttk.Combobox(config_frame, textvariable=self.rotulo_caja_campo_fecha_final_var, width=67)
        self.rotulo_caja_fecha_fin_combo.grid(row=row+4, column=1, sticky="w", pady=3)

        self.reload_rotulo_caja_config()

    def reload_rotulo_caja_config(self):
        data = get_rotulo_caja_config()
        self.rotulo_caja_campo_busqueda_var.set(data.get("campo_busqueda", "Número de orden"))
        self.rotulo_caja_campo_caja_var.set(data.get("campo_caja", "Caja"))
        self.rotulo_caja_campo_unidad_var.set(data.get("campo_unidad_documental", "Nombre unidad documental"))
        self.rotulo_caja_campo_fecha_inicial_var.set(data.get("campo_fecha_inicial", "Fecha inicial"))
        self.rotulo_caja_campo_fecha_final_var.set(data.get("campo_fecha_final", "Fecha final"))
        self.rotulo_caja_titulo_var.set(data.get("titulo", "ALCALDIA DE RIONEGRO"))
        self.rotulo_caja_dependencia_var.set(data.get("dependencia", "SECRETARIA DE HACIENDA-SUBSECRETARIA DE TESORERIA"))
        self.rotulo_caja_serie_var.set(data.get("serie", "COMPROBANTES CONTABLES"))
        self.rotulo_caja_subserie_var.set(data.get("subserie", "COMPROBANTES CONTABLES DE EGRESO"))
        self.rotulo_caja_texto_consecutivo_var.set(data.get("texto_consecutivo", "INICIA CON EL COMPROBANTE DE EGRESO N°"))
        self.rotulo_caja_texto_correlativo_var.set(data.get("texto_correlativo", "FINALIZA CON EL COMPROBANTE DE EGRESO N°"))
        self.rotulo_caja_observaciones_var.set(data.get("observaciones", ""))

        field_names = [f["column_name"] for f in get_fields()]
        self.rotulo_caja_busqueda_combo["values"] = field_names
        self.rotulo_caja_caja_combo["values"] = field_names
        self.rotulo_caja_unidad_combo["values"] = field_names
        self.rotulo_caja_fecha_ini_combo["values"] = field_names
        self.rotulo_caja_fecha_fin_combo["values"] = field_names

    def save_rotulo_caja_config(self):
        save_rotulo_caja_config(
            self.user["username"],
            {
                "campo_busqueda": self.rotulo_caja_campo_busqueda_var.get().strip() or "Número de orden",
                "campo_caja": self.rotulo_caja_campo_caja_var.get().strip() or "Caja",
                "campo_unidad_documental": self.rotulo_caja_campo_unidad_var.get().strip() or "Nombre unidad documental",
                "campo_fecha_inicial": self.rotulo_caja_campo_fecha_inicial_var.get().strip() or "Fecha inicial",
                "campo_fecha_final": self.rotulo_caja_campo_fecha_final_var.get().strip() or "Fecha final",
                "titulo": self.rotulo_caja_titulo_var.get().strip(),
                "dependencia": self.rotulo_caja_dependencia_var.get().strip(),
                "serie": self.rotulo_caja_serie_var.get().strip(),
                "subserie": self.rotulo_caja_subserie_var.get().strip(),
                "texto_consecutivo": self.rotulo_caja_texto_consecutivo_var.get().strip(),
                "texto_correlativo": self.rotulo_caja_texto_correlativo_var.get().strip(),
                "observaciones": self.rotulo_caja_observaciones_var.get().strip(),
            }
        )
        messagebox.showinfo("Listo", "Configuración del Rótulo de Caja guardada correctamente.")

    # GENERAR RÓTULO CAJA
    def build_generate_rotulo_caja_tab(self):
        outer = ttk.Frame(self.generate_rotulo_caja_tab, padding=18)
        outer.pack(fill="both", expand=True)

        frm = ttk.LabelFrame(outer, text="Generar Rótulo de Caja", padding=16)
        frm.pack(fill="x")

        self.rotulo_caja_caja_var = tk.StringVar()
        self.rotulo_caja_desde_var = tk.StringVar()
        self.rotulo_caja_hasta_var = tk.StringVar()
        self.rotulo_caja_format_var = tk.StringVar(value="Excel")
        self.rotulo_caja_output_dir_var = tk.StringVar(value=str(APP_DIR / "salidas"))

        ttk.Label(frm, text="Caja").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.rotulo_caja_caja_var, width=25).grid(row=0, column=1, sticky="w", pady=5)

        ttk.Label(frm, text="No. orden inicial").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.rotulo_caja_desde_var, width=25).grid(row=1, column=1, sticky="w", pady=5)

        ttk.Label(frm, text="No. orden final").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.rotulo_caja_hasta_var, width=25).grid(row=2, column=1, sticky="w", pady=5)

        ttk.Label(frm, text="Formato").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Combobox(frm, textvariable=self.rotulo_caja_format_var, values=["Excel", "PDF", "Word"], state="readonly", width=22).grid(row=3, column=1, sticky="w", pady=5)

        ttk.Label(frm, text="Carpeta de salida").grid(row=4, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.rotulo_caja_output_dir_var, width=70).grid(row=4, column=1, sticky="w", pady=5)
        ttk.Button(frm, text="Seleccionar", command=self.choose_rotulo_caja_output_dir).grid(row=4, column=2, padx=8)

        ttk.Button(frm, text="Generar Rótulo de Caja", command=self.run_generate_rotulo_caja).grid(row=5, column=1, sticky="w", pady=(14, 0))

    def choose_rotulo_caja_output_dir(self):
        directory = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if directory:
            self.rotulo_caja_output_dir_var.set(directory)

    def run_generate_rotulo_caja(self):
        try:
            path = generate_rotulo_caja(
                self.user["username"],
                self.rotulo_caja_caja_var.get(),
                self.rotulo_caja_desde_var.get(),
                self.rotulo_caja_hasta_var.get(),
                self.rotulo_caja_format_var.get(),
                self.rotulo_caja_output_dir_var.get(),
            )
            messagebox.showinfo("Listo", f"Rótulo de Caja generado en:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


    def build_users_tab(self):
        outer = ttk.Frame(self.users_tab, padding=10)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x", pady=(0, 10))
        ttk.Button(top, text="Recargar", command=self.refresh_users).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Crear usuario", command=self.create_user_dialog).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Activar/Desactivar", command=self.toggle_user_status).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Resetear clave", command=self.reset_password_dialog).pack(side="left")

        self.users_tree = ttk.Treeview(
            outer,
            columns=("id", "username", "full_name", "role", "active", "must_change_password", "last_login"),
            show="headings",
            height=18
        )

        widths = {"id": 60, "username": 120, "full_name": 180, "role": 90, "active": 70, "must_change_password": 160, "last_login": 160}
        for col in self.users_tree["columns"]:
            self.users_tree.heading(col, text=col)
            self.users_tree.column(col, width=widths[col], anchor="w")

        self.users_tree.pack(fill="both", expand=True)
        self.refresh_users()

    def refresh_users(self):
        for item in self.users_tree.get_children():
            self.users_tree.delete(item)
        for u in list_users():
            self.users_tree.insert("", "end", values=(u["id"], u["username"], u["full_name"], u["role"], u["active"], u["must_change_password"], u["last_login"] or ""))

    def get_selected_user_id(self):
        sel = self.users_tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Seleccione un usuario.")
            return None
        return int(self.users_tree.item(sel[0])["values"][0])

    def create_user_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("Crear usuario")
        dlg.geometry("360x250")
        dlg.resizable(False, False)
        dlg.grab_set()

        username = tk.StringVar()
        full_name = tk.StringVar()
        password = tk.StringVar(value="123456")
        role = tk.StringVar(value="normal")

        frm = ttk.Frame(dlg, padding=15)
        frm.pack(fill="both", expand=True)

        for i, (label, var, show) in enumerate([
            ("Usuario", username, None),
            ("Nombre completo", full_name, None),
            ("Clave temporal", password, "*")
        ]):
            ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w", pady=6)
            ttk.Entry(frm, textvariable=var, show=show).grid(row=i, column=1, sticky="ew", pady=6)

        ttk.Label(frm, text="Rol").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Combobox(frm, textvariable=role, values=["admin", "normal"], state="readonly").grid(row=3, column=1, sticky="ew", pady=6)

        frm.columnconfigure(1, weight=1)

        def submit():
            ok, msg = validate_new_password(password.get().strip())
            if not ok:
                messagebox.showerror("Error", msg)
                return
            try:
                create_user(self.user["username"], username.get(), full_name.get(), password.get(), _normalize_role_value(role.get()))
                self.refresh_users()
                dlg.destroy()
                messagebox.showinfo("Listo", "Usuario creado correctamente.")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        ttk.Button(frm, text="Guardar", command=submit).grid(row=4, column=0, columnspan=2, sticky="ew", pady=10)

    def toggle_user_status(self):
        uid = self.get_selected_user_id()
        if uid is None:
            return
        row = next((x for x in list_users() if x["id"] == uid), None)
        if not row:
            return
        update_user_status(self.user["username"], uid, 0 if int(row["active"]) == 1 else 1)
        self.refresh_users()
        messagebox.showinfo("Listo", "Estado del usuario actualizado.")

    def reset_password_dialog(self):
        uid = self.get_selected_user_id()
        if uid is None:
            return

        dlg = tk.Toplevel(self)
        dlg.title("Resetear clave")
        dlg.geometry("360x150")
        dlg.resizable(False, False)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=15)
        frm.pack(fill="both", expand=True)

        temp = tk.StringVar(value="123456")
        ttk.Label(frm, text="Nueva clave temporal").pack(anchor="w")
        ttk.Entry(frm, textvariable=temp, show="*").pack(fill="x", pady=(5, 10))

        def submit():
            ok, msg = validate_new_password(temp.get().strip())
            if not ok:
                messagebox.showerror("Error", msg)
                return
            reset_user_password(self.user["username"], uid, temp.get().strip())
            self.refresh_users()
            dlg.destroy()
            messagebox.showinfo("Listo", "Clave reseteada correctamente.")

        ttk.Button(frm, text="Guardar", command=submit).pack(fill="x")

    # AUDITORÍA
    def build_audit_tab(self):
        outer = ttk.Frame(self.audit_tab, padding=10)
        outer.pack(fill="both", expand=True)

        ttk.Button(outer, text="Recargar", command=self.refresh_audit).pack(anchor="w", pady=(0, 10))

        self.audit_tree = ttk.Treeview(
            outer,
            columns=("id", "username", "action", "module", "record_id", "details", "created_at"),
            show="headings",
            height=20
        )

        widths = {"id": 60, "username": 120, "action": 150, "module": 120, "record_id": 90, "details": 380, "created_at": 170}
        for col in self.audit_tree["columns"]:
            self.audit_tree.heading(col, text=col)
            self.audit_tree.column(col, width=widths[col], anchor="w")

        self.audit_tree.pack(fill="both", expand=True)
        self.refresh_audit()

    def refresh_audit(self):
        for item in self.audit_tree.get_children():
            self.audit_tree.delete(item)
        for r in get_audit_rows():
            self.audit_tree.insert("", "end", values=(r["id"], r["username"], r["action"], r["module"], r["record_id"] or "", r["details"] or "", r["created_at"]))


def main():
    initialize_database()
    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    LoginWindow(root)
    root.mainloop()
