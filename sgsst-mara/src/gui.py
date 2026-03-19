"""
SGSST - MARA  |  Interfaz gráfica principal
Ejecutar desde la raíz del proyecto:
    python -m src.gui
"""
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from src.database import init_db
from src.generar_documento import generar_todas_constancias
from src.importar_excel import importar_trabajadores
from src.importar_sheets import sincronizar_desde_sheets
from src.procesar_firmas import procesar_todas_firmas
from src.registro_capacitacion import guardar_registro_capacitacion
from src.trabajadores import buscar_por_dni, listar_activos


# ── helpers ─────────────────────────────────────────────────────────────────

def _scrolled_text(parent, height: int) -> tk.Text:
    frame = ttk.Frame(parent)
    frame.pack(fill='both', expand=True)
    widget = tk.Text(frame, height=height, font=('Consolas', 9),
                     state='disabled', wrap='none')
    sb_y = ttk.Scrollbar(frame, orient='vertical',   command=widget.yview)
    sb_x = ttk.Scrollbar(frame, orient='horizontal',  command=widget.xview)
    widget.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
    sb_y.pack(side='right',  fill='y')
    sb_x.pack(side='bottom', fill='x')
    widget.pack(fill='both', expand=True)
    return widget


def _write(widget: tk.Text, text: str, clear: bool = True) -> None:
    widget.configure(state='normal')
    if clear:
        widget.delete('1.0', 'end')
    widget.insert('end', text)
    widget.see('end')
    widget.configure(state='disabled')


# ── main app ─────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SGSST - MARA")
        self.geometry("720x580")
        self.resizable(False, False)

        self._setup_style()

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=10, pady=10)

        self._build_tab_trabajadores(nb)
        self._build_tab_constancias(nb)
        self._build_tab_firmas(nb)
        self._build_tab_capacitaciones(nb)

        init_db()

    # ── style ────────────────────────────────────────────────────────────────

    def _setup_style(self):
        s = ttk.Style()
        s.configure('TLabel',    font=('Segoe UI', 10))
        s.configure('TButton',   font=('Segoe UI', 10))
        s.configure('TEntry',    font=('Segoe UI', 10))
        s.configure('TCombobox', font=('Segoe UI', 10))
        s.configure('Bold.TLabel', font=('Segoe UI', 10, 'bold'))

    # ════════════════════════════════════════════════════════════════════════
    # TAB 1 — TRABAJADORES
    # ════════════════════════════════════════════════════════════════════════

    def _build_tab_trabajadores(self, nb: ttk.Notebook):
        tab = ttk.Frame(nb, padding=15)
        nb.add(tab, text="  Trabajadores  ")

        # ── acciones ─────────────────────────────────────────────────────────
        acc = ttk.LabelFrame(tab, text="Importar / Sincronizar", padding=10)
        acc.pack(fill='x', pady=(0, 8))

        ttk.Button(acc, text="Sincronizar desde Google Sheets",
                   command=self._sync_sheets).pack(side='left', padx=(0, 8))
        ttk.Button(acc, text="Importar desde Excel",
                   command=self._importar_excel).pack(side='left')

        # ── búsqueda ─────────────────────────────────────────────────────────
        srch = ttk.LabelFrame(tab, text="Consultar", padding=10)
        srch.pack(fill='x', pady=(0, 8))

        row = ttk.Frame(srch)
        row.pack(fill='x')
        ttk.Label(row, text="DNI:").pack(side='left', padx=(0, 5))
        self._dni_var = tk.StringVar()
        ttk.Entry(row, textvariable=self._dni_var, width=14).pack(side='left', padx=(0, 8))
        ttk.Button(row, text="Buscar",
                   command=self._buscar_dni).pack(side='left', padx=(0, 8))
        ttk.Button(row, text="Listar activos",
                   command=self._listar_activos).pack(side='left')

        # ── resultados ───────────────────────────────────────────────────────
        res = ttk.LabelFrame(tab, text="Resultados", padding=10)
        res.pack(fill='both', expand=True)
        self._txt_trabajadores = _scrolled_text(res, height=14)

    def _sync_sheets(self):
        def run():
            try:
                res = sincronizar_desde_sheets()
                msg = (
                    f"Sincronización completada\n"
                    f"  Nuevos      : {res['nuevos']}\n"
                    f"  Actualizados: {res['actualizados']}\n"
                    f"  Omitidos    : {res['omitidos']}"
                )
                if res['errores']:
                    msg += "\n\nErrores:\n" + "\n".join(f"  {e}" for e in res['errores'])
                self.after(0, lambda: messagebox.showinfo("Google Sheets", msg))
            except Exception as exc:
                self.after(0, lambda: messagebox.showerror("Error", str(exc)))

        threading.Thread(target=run, daemon=True).start()

    def _importar_excel(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if not ruta:
            return
        try:
            res = importar_trabajadores(ruta)
            msg = (
                f"Importación completada\n"
                f"  Importados: {res['importados']}\n"
                f"  Omitidos  : {res['omitidos']}"
            )
            if res['errores']:
                msg += "\n\nErrores:\n" + "\n".join(f"  {e}" for e in res['errores'])
            messagebox.showinfo("Importar Excel", msg)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _buscar_dni(self):
        dni = self._dni_var.get().strip()
        if not dni:
            messagebox.showwarning("Aviso", "Ingresa un DNI")
            return
        t = buscar_por_dni(dni)
        if t is None:
            _write(self._txt_trabajadores, "Trabajador no encontrado.")
            return
        lines = [f"  {k:<20}: {v}" for k, v in t.items() if v is not None]
        _write(self._txt_trabajadores, "\n".join(lines))

    def _listar_activos(self):
        activos = listar_activos()
        if not activos:
            _write(self._txt_trabajadores, "No hay trabajadores activos registrados.")
            return

        w_dni    = max(max(len(str(t['dni']))            for t in activos), 3)
        w_nombre = max(max(len(str(t['nombre']))         for t in activos), 6)
        w_cargo  = max(max(len(str(t['cargo'] or ''))    for t in activos), 5)

        sep    = f"+-{'-'*w_dni}-+-{'-'*w_nombre}-+-{'-'*w_cargo}-+"
        header = (f"| {'DNI'.ljust(w_dni)} "
                  f"| {'NOMBRE'.ljust(w_nombre)} "
                  f"| {'CARGO'.ljust(w_cargo)} |")
        lines  = [sep, header, sep]
        for t in activos:
            lines.append(
                f"| {str(t['dni']).ljust(w_dni)} "
                f"| {str(t['nombre']).ljust(w_nombre)} "
                f"| {str(t['cargo'] or '-').ljust(w_cargo)} |"
            )
        lines += [sep, f"  Total: {len(activos)} trabajadores activos"]
        _write(self._txt_trabajadores, "\n".join(lines))

    # ════════════════════════════════════════════════════════════════════════
    # TAB 2 — CONSTANCIAS
    # ════════════════════════════════════════════════════════════════════════

    def _build_tab_constancias(self, nb: ttk.Notebook):
        tab = ttk.Frame(nb, padding=15)
        nb.add(tab, text="  Constancias RISST  ")

        # ── configuración ────────────────────────────────────────────────────
        cfg = ttk.LabelFrame(tab, text="Configuración", padding=10)
        cfg.pack(fill='x', pady=(0, 8))

        row1 = ttk.Frame(cfg)
        row1.pack(fill='x', pady=3)
        ttk.Label(row1, text="Fecha (dd/mm/yyyy):").pack(side='left', padx=(0, 8))
        self._fecha_var = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))
        ttk.Entry(row1, textvariable=self._fecha_var, width=14).pack(side='left')

        ttk.Label(cfg, text="Plantilla : data/plantillas/plantilla_risst.docx",
                  foreground='gray').pack(anchor='w', pady=2)
        ttk.Label(cfg, text="Salida    : documentos_generados/",
                  foreground='gray').pack(anchor='w')

        # ── acción ───────────────────────────────────────────────────────────
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(fill='x', pady=(0, 6))
        ttk.Button(btn_frame,
                   text="Generar constancias para todos los activos",
                   command=self._generar_constancias).pack(side='left')

        self._prog_const = ttk.Progressbar(tab, mode='indeterminate', length=500)
        self._prog_const.pack(fill='x', pady=(0, 8))

        # ── log ──────────────────────────────────────────────────────────────
        log = ttk.LabelFrame(tab, text="Progreso", padding=10)
        log.pack(fill='both', expand=True)
        self._txt_constancias = _scrolled_text(log, height=12)

    def _generar_constancias(self):
        fecha = self._fecha_var.get().strip()
        if not fecha:
            messagebox.showwarning("Aviso", "Ingresa una fecha")
            return

        _write(self._txt_constancias, "")
        self._prog_const.start()

        def run():
            try:
                res = generar_todas_constancias(fecha)
                for line in res['log']:
                    self.after(0, lambda l=line: _write(self._txt_constancias, l + '\n', clear=False))
                summary = (
                    f"\n{'='*50}\n"
                    f"Total: {res['total']}  |  "
                    f"OK: {res['exitosos']}  |  "
                    f"Errores: {res['fallidos']}"
                )
                self.after(0, lambda: _write(self._txt_constancias, summary, clear=False))
                self.after(0, lambda: messagebox.showinfo(
                    "Completado",
                    f"Generación finalizada.\n"
                    f"{res['exitosos']} de {res['total']} constancias generadas.\n"
                    f"Archivos en: documentos_generados/",
                ))
            except Exception as exc:
                self.after(0, lambda: messagebox.showerror("Error", str(exc)))
            finally:
                self.after(0, self._prog_const.stop)

        threading.Thread(target=run, daemon=True).start()

    # ════════════════════════════════════════════════════════════════════════
    # TAB 3 — FIRMAS
    # ════════════════════════════════════════════════════════════════════════

    def _build_tab_firmas(self, nb: ttk.Notebook):
        tab = ttk.Frame(nb, padding=15)
        nb.add(tab, text="  Firmas  ")

        # ── instrucciones ────────────────────────────────────────────────────
        info = ttk.LabelFrame(tab, text="Instrucciones", padding=10)
        info.pack(fill='x', pady=(0, 10))
        ttk.Label(
            info,
            text=(
                "1. Coloca las imágenes de firmas en:   data/firmas_originales/\n"
                "2. Nombre esperado:  firma_{DNI}.png   (ej. firma_05314709.png)\n"
                "3. Presiona 'Procesar firmas'.\n"
                "4. Las imágenes sin fondo quedarán en: data/firmas/\n\n"
                "Archivos ya procesados se omiten automáticamente."
            ),
            justify='left',
            foreground='#333333',
        ).pack(anchor='w')

        # ── acción ───────────────────────────────────────────────────────────
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(fill='x', pady=(0, 6))
        ttk.Button(btn_frame, text="Procesar firmas",
                   command=self._procesar_firmas).pack(side='left')

        self._prog_firmas = ttk.Progressbar(tab, mode='indeterminate', length=500)
        self._prog_firmas.pack(fill='x', pady=(0, 8))

        # ── log ──────────────────────────────────────────────────────────────
        log = ttk.LabelFrame(tab, text="Progreso", padding=10)
        log.pack(fill='both', expand=True)
        self._txt_firmas = _scrolled_text(log, height=14)

    def _procesar_firmas(self):
        _write(self._txt_firmas, "")
        self._prog_firmas.start()

        def run():
            try:
                res = procesar_todas_firmas()
                lines = [
                    f"Procesadas : {res['procesadas']}",
                    f"Omitidas   : {res['omitidas']}  (ya existían)",
                ]
                for err in res['errores']:
                    lines.append(f"Error      : {err}")
                self.after(0, lambda: _write(self._txt_firmas, "\n".join(lines)))
                self.after(0, lambda: messagebox.showinfo(
                    "Completado",
                    f"Procesadas : {res['procesadas']}\n"
                    f"Omitidas   : {res['omitidas']}\n"
                    f"Errores    : {len(res['errores'])}",
                ))
            except Exception as exc:
                self.after(0, lambda: messagebox.showerror("Error", str(exc)))
            finally:
                self.after(0, self._prog_firmas.stop)

        threading.Thread(target=run, daemon=True).start()


    # ════════════════════════════════════════════════════════════════════════
    # TAB 4 — CAPACITACIONES
    # ════════════════════════════════════════════════════════════════════════

    def _build_tab_capacitaciones(self, nb: ttk.Notebook):
        tab = ttk.Frame(nb, padding=15)
        nb.add(tab, text="  Capacitaciones  ")

        # ── detalles ─────────────────────────────────────────────────────────
        det = ttk.LabelFrame(tab, text="Detalles de la Capacitación", padding=10)
        det.pack(fill='x', pady=(0, 8))

        ttk.Label(det, text="Tema:").grid(row=0, column=0, sticky='e', padx=5, pady=4)
        self._cap_tema_var = tk.StringVar()
        ttk.Entry(det, textvariable=self._cap_tema_var, width=45).grid(row=0, column=1, pady=4, sticky='w')

        ttk.Label(det, text="Fecha (dd/mm/yyyy):").grid(row=1, column=0, sticky='e', padx=5, pady=4)
        self._cap_fecha_var = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))
        ttk.Entry(det, textvariable=self._cap_fecha_var, width=15).grid(row=1, column=1, pady=4, sticky='w')

        ttk.Label(det, text="Hora (HH:MM):").grid(row=2, column=0, sticky='e', padx=5, pady=4)
        self._cap_hora_var = tk.StringVar()
        ttk.Entry(det, textvariable=self._cap_hora_var, width=10).grid(row=2, column=1, pady=4, sticky='w')

        # ── selección de trabajadores ─────────────────────────────────────────
        sel = ttk.LabelFrame(tab, text="Asistentes", padding=10)
        sel.pack(fill='both', expand=True, pady=(0, 8))

        top = ttk.Frame(sel)
        top.pack(fill='x', pady=(0, 5))
        ttk.Label(top, text="Trabajador:").pack(side='left', padx=(0, 5))
        self._cap_combo = ttk.Combobox(top, width=45, state='readonly')
        self._cap_combo.pack(side='left', padx=(0, 8))
        ttk.Button(top, text="Agregar", command=self._cap_agregar).pack(side='left', padx=(0, 5))
        ttk.Button(top, text="Quitar seleccionado", command=self._cap_quitar).pack(side='left')

        self._cap_listbox = tk.Listbox(sel, height=8, font=('Consolas', 9))
        sb = ttk.Scrollbar(sel, orient='vertical', command=self._cap_listbox.yview)
        self._cap_listbox.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self._cap_listbox.pack(fill='both', expand=True)

        # ── botones ───────────────────────────────────────────────────────────
        btns = ttk.Frame(tab)
        btns.pack(fill='x')
        ttk.Button(btns, text="Guardar registro Excel",
                   command=self._cap_guardar).pack(side='left', padx=(0, 8))
        ttk.Button(btns, text="Limpiar",
                   command=self._cap_limpiar).pack(side='left')

        # estado interno
        self._cap_seleccionados: list[dict] = []   # trabajadores ya agregados
        self._cap_disponibles:   list[dict] = []   # trabajadores disponibles para agregar
        self._cap_refresh_combo()

    def _cap_refresh_combo(self):
        """Recarga la lista de disponibles desde la BD y actualiza el combo."""
        seleccionados_dni = {t['dni'] for t in self._cap_seleccionados}
        self._cap_disponibles = [
            t for t in listar_activos() if t['dni'] not in seleccionados_dni
        ]
        valores = [
            f"{t['dni']} - {t['nombre']}"
            for t in self._cap_disponibles
        ]
        self._cap_combo['values'] = valores
        self._cap_combo.set('')

    def _cap_agregar(self):
        sel = self._cap_combo.get()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona un trabajador")
            return
        dni = sel.split(' - ')[0]
        trabajador = next((t for t in self._cap_disponibles if t['dni'] == dni), None)
        if trabajador is None:
            return
        self._cap_seleccionados.append(trabajador)
        self._cap_listbox.insert(
            tk.END, f"{trabajador['dni']}  {trabajador['nombre']}  |  {trabajador.get('cargo') or ''}"
        )
        self._cap_refresh_combo()

    def _cap_quitar(self):
        idx = self._cap_listbox.curselection()
        if not idx:
            messagebox.showwarning("Aviso", "Selecciona un trabajador de la lista")
            return
        i = idx[0]
        self._cap_seleccionados.pop(i)
        self._cap_listbox.delete(i)
        self._cap_refresh_combo()

    def _cap_limpiar(self):
        self._cap_tema_var.set('')
        self._cap_fecha_var.set(datetime.now().strftime('%d/%m/%Y'))
        self._cap_hora_var.set('')
        self._cap_seleccionados.clear()
        self._cap_listbox.delete(0, tk.END)
        self._cap_refresh_combo()

    def _cap_guardar(self):
        tema  = self._cap_tema_var.get().strip()
        fecha = self._cap_fecha_var.get().strip()
        hora  = self._cap_hora_var.get().strip()
        if not all([tema, fecha, hora]):
            messagebox.showwarning("Aviso", "Tema, fecha y hora son obligatorios")
            return
        if not self._cap_seleccionados:
            messagebox.showwarning("Aviso", "Agrega al menos un trabajador")
            return
        try:
            ruta = guardar_registro_capacitacion(tema, fecha, hora, self._cap_seleccionados)
            messagebox.showinfo(
                "Listo",
                f"Registro guardado en:\n{ruta}"
            )
        except Exception as exc:
            messagebox.showerror("Error", str(exc))


# ── entry point ──────────────────────────────────────────────────────────────

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
