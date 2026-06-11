import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

import main


class CredencialesDialog(simpledialog.Dialog):
    def body(self, master):
        self.title("Guardar Credenciales")
        tk.Label(master, text="Usuario:").grid(row=0, sticky=tk.W, padx=5, pady=3)
        tk.Label(master, text="Contraseña:").grid(row=1, sticky=tk.W, padx=5, pady=3)
        self.usuario = tk.Entry(master, width=30)
        self.contrasena = tk.Entry(master, show="*", width=30)
        self.usuario.grid(row=0, column=1, padx=5, pady=3)
        self.contrasena.grid(row=1, column=1, padx=5, pady=3)
        return self.usuario

    def apply(self):
        self.result = {
            "USER": self.usuario.get().strip(),
            "PASS": self.contrasena.get().strip(),
        }


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CE1121 – Gestión Contabilización Activos Fijos")
        self.geometry("420x370")
        self.resizable(False, False)
        self.worker_thread: threading.Thread | None = None
        self._cancelar: bool = False
        self.estado_var = tk.StringVar(value="Estado: Esperando acción")
        self._crear_widgets()

    # ─────────────────────────────────────────────────────────────────────────
    # UI
    # ─────────────────────────────────────────────────────────────────────────

    def _crear_widgets(self) -> None:
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        self.estado_label = ttk.Label(
            frame, textvariable=self.estado_var,
            anchor="center", wraplength=380,
        )
        self.estado_label.pack(pady=(0, 6), fill=tk.X)

        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.pack(pady=4, fill=tk.X)
        self.progress.stop()
        self.progress.pack_forget()

        botones = [
            ("Guardar Credenciales",                self.guardar_credenciales),
            ("Ejecutar Proceso Principal",           self.iniciar_proceso),
            ("Cancelar Proceso",                     self.cancelar_proceso),
            ("Reiniciar Proceso",                    self.reiniciar_proceso),
            ("Reiniciar Manteniendo Referencia",     self.reiniciar_con_referencia),
            ("Cerrar",                               self.cerrar_consola),
        ]
        self._botones: dict[str, ttk.Button] = {}
        for texto, comando in botones:
            btn = ttk.Button(frame, text=texto, command=comando)
            btn.pack(pady=2, fill=tk.X)
            self._botones[texto] = btn

        # Estado inicial de botones
        self._botones["Cancelar Proceso"].config(state=tk.DISABLED)

    # ─────────────────────────────────────────────────────────────────────────
    # Credenciales
    # ─────────────────────────────────────────────────────────────────────────

    def guardar_credenciales(self) -> None:
        dialog = CredencialesDialog(self)
        if not dialog.result:
            return

        user = dialog.result["USER"]
        password = dialog.result["PASS"]

        if not user or not password:
            messagebox.showwarning("Validación", "Usuario y contraseña son obligatorios.")
            return

        try:
            from encriptador import guardar_credenciales
            import os

            ruta = guardar_credenciales(user, password)
            self.estado_var.set("Credenciales guardadas correctamente.")
            messagebox.showinfo(
                "Éxito",
                f"Credenciales guardadas correctamente.\n\nRuta:\n{ruta}\n\n"
                "La aplicación se reiniciará para cargar las nuevas credenciales.",
            )
            # Reiniciar el proceso para que mainencriptador cargue el nuevo .env
            os.execl(sys.executable, sys.executable, *sys.argv)

        except Exception as e:
            messagebox.showerror(
                "Error", f"No se pudieron guardar las credenciales:\n{e}"
            )

    def _credenciales_existentes(self) -> bool:
        import os
        base_dir = os.path.dirname(os.path.abspath(__file__))
        env_path = os.path.join(base_dir, "config", ".env")
        return os.path.exists(env_path)

    # ─────────────────────────────────────────────────────────────────────────
    # Ejecución del proceso
    # ─────────────────────────────────────────────────────────────────────────

    def iniciar_proceso(self) -> None:
        if not self._credenciales_existentes():
            messagebox.showwarning(
                "Credenciales requeridas",
                "No hay credenciales guardadas.\n\n"
                "Use el botón 'Guardar Credenciales' primero.",
            )
            return

        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning(
                "Proceso en ejecución",
                "Ya hay un proceso corriendo. Espere a que termine o use 'Cancelar Proceso'.",
            )
            return

        self.estado_var.set("Estado: Ejecutando proceso principal...")
        self._botones["Ejecutar Proceso Principal"].config(state=tk.DISABLED)
        self._botones["Cancelar Proceso"].config(state=tk.NORMAL)
        self.progress.pack()
        self.progress.start(10)
        self._cancelar = False

        def tarea():
            try:
                self._actualizar_estado("Proceso iniciado...")
                main.ejecutar_proceso()
                if not self._cancelar:
                    self._actualizar_estado("Proceso finalizado exitosamente.")
            except Exception as e:
                self._actualizar_error(str(e))
            finally:
                self._finalizar_proceso()

        self.worker_thread = threading.Thread(target=tarea, daemon=True)
        self.worker_thread.start()

    def cancelar_proceso(self) -> None:
        self._cancelar = True
        self.estado_var.set("Estado: Cancelando proceso... (espere)")
        self._botones["Cancelar Proceso"].config(state=tk.DISABLED)

    def reiniciar_proceso(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning(
                "Proceso en ejecución",
                "No puede reiniciar mientras hay un proceso corriendo.",
            )
            return
        try:
            main.reiniciar_ejecucion()
            self.estado_var.set("Estado: Proceso reiniciado. Listo para ejecutar.")
            messagebox.showinfo("Reinicio Exitoso", "El proceso se reinició completamente.")
            self._restaurar_controles()
        except Exception as e:
            self.estado_var.set(f"Error al reiniciar: {e}")
            messagebox.showerror("Error", f"No se pudo reiniciar el proceso:\n{e}")

    def reiniciar_con_referencia(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning(
                "Proceso en ejecución",
                "No puede reiniciar mientras hay un proceso corriendo.",
            )
            return
        try:
            main.reiniciar_ejecucion_con_referencia()
            self.estado_var.set("Estado: Reiniciado conservando referencia.")
            messagebox.showinfo(
                "Reinicio con Referencia",
                "El proceso se reinició conservando el Excel de referencia.",
            )
            self._restaurar_controles()
        except Exception as e:
            self.estado_var.set(f"Error al reiniciar con referencia: {e}")
            messagebox.showerror("Error", f"No se pudo reiniciar con referencia:\n{e}")

    # ─────────────────────────────────────────────────────────────────────────
    # Helpers internos (thread-safe usando after())
    # ─────────────────────────────────────────────────────────────────────────

    def _actualizar_estado(self, mensaje: str) -> None:
        """Thread-safe: actualiza la etiqueta de estado desde un hilo secundario."""
        self.after(0, lambda: self.estado_var.set(f"Estado: {mensaje}"))

    def _actualizar_error(self, mensaje: str) -> None:
        """Thread-safe: muestra error desde un hilo secundario."""
        def _mostrar():
            self.estado_var.set(f"Error: {mensaje}")
            messagebox.showerror("Error en el proceso", mensaje)
        self.after(0, _mostrar)

    def _finalizar_proceso(self) -> None:
        """Thread-safe: restaura la UI al finalizar el proceso."""
        self.after(0, self._restaurar_controles)

    def _restaurar_controles(self) -> None:
        self.progress.stop()
        self.progress.pack_forget()
        self._botones["Ejecutar Proceso Principal"].config(state=tk.NORMAL)
        self._botones["Cancelar Proceso"].config(state=tk.DISABLED)

    def cerrar_consola(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            if not messagebox.askyesno(
                "Proceso en ejecución",
                "Hay un proceso en ejecución. ¿Desea cerrar de todas formas?\n"
                "El proceso se interrumpirá.",
            ):
                return
        self.destroy()
        sys.exit(0)


if __name__ == "__main__":
    app = App()
    app.mainloop()