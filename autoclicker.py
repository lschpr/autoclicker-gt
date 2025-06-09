import json
import os
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pyautogui
from pynput import keyboard

def parse_hotkey_string(s: str):
    s = s.strip().lower()
    if not s:
        return None
    try:
        key_obj = getattr(keyboard.Key, s)
        return key_obj
    except (AttributeError, TypeError):
        pass
    if len(s) == 1:
        return keyboard.KeyCode.from_char(s)
    return None

def hotkey_to_string(key_obj):
    if isinstance(key_obj, keyboard.Key):
        return key_obj.name
    elif isinstance(key_obj, keyboard.KeyCode):
        return key_obj.char or ""
    else:
        return ""

class Macro:
    def __init__(self, name: str, trigger_key, button: str, key_to_send: str,
                 n_clicks: int, interval: float, x_coord=None, y_coord=None, start_delay=0.0):
        self.name = name
        self.trigger_key = trigger_key
        self.button = button
        self.key_to_send = key_to_send
        self.n_clicks = max(1, n_clicks)
        self.interval = max(0.0, interval)
        self.x_coord = x_coord if x_coord is not None else None
        self.y_coord = y_coord if y_coord is not None else None
        self.start_delay = max(0.0, start_delay)

    def display_name(self):
        hk = hotkey_to_string(self.trigger_key) or "?"
        return f"{self.name} [{hk}]"

    def to_dict(self):
        return {
            "name": self.name,
            "trigger_key": hotkey_to_string(self.trigger_key),
            "button": self.button,
            "key_to_send": self.key_to_send,
            "n_clicks": self.n_clicks,
            "interval": self.interval,
            "x_coord": self.x_coord,
            "y_coord": self.y_coord,
            "start_delay": self.start_delay,
        }

    @staticmethod
    def from_dict(d):
        tk_str = d.get("trigger_key", "").strip()
        tk_parsed = parse_hotkey_string(tk_str)
        return Macro(
            name=d.get("name", "(no name)"),
            trigger_key=tk_parsed,
            button=d.get("button", "left"),
            key_to_send=d.get("key_to_send", ""),
            n_clicks=int(d.get("n_clicks", 1)),
            interval=float(d.get("interval", 0.1)),
            x_coord=(None if d.get("x_coord") is None else int(d.get("x_coord"))),
            y_coord=(None if d.get("y_coord") is None else int(d.get("y_coord"))),
            start_delay=float(d.get("start_delay", 0.0)),
        )

class AutoClicker:
    def __init__(self, gui_reference):
        self.gui = gui_reference
        self.clicks_per_second = 1.0
        self.trigger_key = keyboard.Key.f3
        self.stop_after_total = 0
        self.mode = "press"
        self.button = "left"
        self.key_to_send = ""
        self.use_fixed_master = False
        self.master_x = None
        self.master_y = None
        self._clicking = False
        self._total_clicks_sent = 0
        self._toggle_thread = None
        self.macros = []
        self._load_macros_from_disk()
        self._listener = keyboard.Listener(
            on_press=self._on_key_press,
            on_release=self._on_key_release
        )
        self._listener.daemon = True
        self._listener.start()

    def _load_macros_from_disk(self):
        try:
            if os.path.isfile("macros.json"):
                with open("macros.json", "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.macros = [Macro.from_dict(item) for item in data]
            else:
                self.macros = []
        except Exception as e:
            messagebox.showwarning("Load Error", f"Failed to load macros.json:\n{e}")
            self.macros = []

    def _save_macros_to_disk(self):
        try:
            with open("macros.json", "w", encoding="utf-8") as f:
                json.dump([m.to_dict() for m in self.macros], f, indent=2)
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save macros.json:\n{e}")

    def _on_key_press(self, key):
        for macro in self.macros:
            if key == macro.trigger_key:
                threading.Thread(target=self._run_macro, args=(macro,), daemon=True).start()
                return
        if key == self.trigger_key:
            if self.mode == "press":
                if not self._clicking:
                    self._start_continuous_master()
            elif self.mode == "toggle":
                if not self._clicking:
                    self._start_continuous_master()
                else:
                    self._stop_continuous_master()

    def _on_key_release(self, key):
        if self.mode == "press" and key == self.trigger_key:
            self._stop_continuous_master()

    def _start_continuous_master(self):
        if self._toggle_thread and self._toggle_thread.is_alive():
            return
        self._clicking = True
        self.gui.set_status("clicking")
        self._toggle_thread = threading.Thread(target=self._continuous_master_loop, daemon=True)
        self._toggle_thread.start()

    def _continuous_master_loop(self):
        cps = max(self.clicks_per_second, 0.01)
        interval = 1.0 / cps
        while self._clicking:
            if self.stop_after_total > 0 and self._total_clicks_sent >= self.stop_after_total:
                break
            self._send_one_click_master()
            self._total_clicks_sent += 1
            time.sleep(interval)
        self._clicking = False
        self.gui.set_status("idle")

    def _stop_continuous_master(self):
        self._clicking = False
        self.gui.set_status("idle")

    def _send_one_click_master(self):
        if self.use_fixed_master and self.master_x is not None and self.master_y is not None:
            pyautogui.click(x=self.master_x, y=self.master_y, button=self.button)
        else:
            if self.button in ("left", "middle", "right"):
                pyautogui.click(button=self.button)
            else:
                k = self.key_to_send.strip().lower()
                if k:
                    pyautogui.press(k)

    def _run_macro(self, macro: Macro):
        if macro.start_delay > 0:
            self.gui.set_status(f"delaying {macro.name}")
            self.gui.set_status_color("orange")
            time.sleep(macro.start_delay)
        self.gui.set_status(f"macro: {macro.name}")
        self.gui.set_status_color("orange")
        for _ in range(macro.n_clicks):
            if self.stop_after_total > 0 and self._total_clicks_sent >= self.stop_after_total:
                break
            if macro.x_coord is not None and macro.y_coord is not None:
                pyautogui.click(x=macro.x_coord, y=macro.y_coord, button=macro.button)
            else:
                if macro.button in ("left", "middle", "right"):
                    pyautogui.click(button=macro.button)
                else:
                    k = macro.key_to_send.strip().lower()
                    if k:
                        pyautogui.press(k)
            self._total_clicks_sent += 1
            time.sleep(macro.interval)
        if self._clicking:
            self.gui.set_status("clicking")
            self.gui.set_status_color("red")
        else:
            self.gui.set_status("idle")
            self.gui.set_status_color("green")

    def update_settings_from_gui(self):
        try:
            cps = float(self.gui.var_n_clicks.get())
            if cps <= 0:
                cps = 0.01
        except ValueError:
            cps = 0.01
        self.clicks_per_second = cps
        desired = self.gui.var_trigger_key.get().strip()
        pk = parse_hotkey_string(desired)
        if pk:
            self.trigger_key = pk
        try:
            sa = int(self.gui.var_stop_at.get())
            if sa < 0:
                sa = 0
        except ValueError:
            sa = 0
        self.stop_after_total = sa
        self.mode = self.gui.var_mode.get()
        sel = self.gui.var_button_choice.get()
        if sel in ("left", "middle", "right", "key"):
            self.button = sel
        else:
            self.button = "left"
        self.key_to_send = self.gui.var_key_to_send.get().strip().lower()
        self.use_fixed_master = bool(self.gui.var_use_fixed.get())
        if self.use_fixed_master:
            try:
                mx = int(self.gui.var_master_x.get())
                my = int(self.gui.var_master_y.get())
            except ValueError:
                mx = None
                my = None
            self.master_x = mx
            self.master_y = my
        else:
            self.master_x = None
            self.master_y = None

    def stop_immediately(self):
        self._clicking = False
        self._total_clicks_sent = 0
        self.gui.set_status("idle")
        self.gui.set_status_color("green")

    def shutdown(self):
        self._clicking = False
        self._total_clicks_sent = 0
        if self._listener:
            self._listener.stop()
        self._save_macros_to_disk()

    @property
    def total_clicks_sent(self):
        return self._total_clicks_sent

class AutoClickerGUI:
    def __init__(self, root):
        self.root = root
        root.title("AutoClicker + Persistent Macros")
        root.resizable(False, False)
        status_frame = ttk.Frame(root, padding=(8, 8))
        status_frame.grid(row=0, column=0, sticky="ew")
        status_frame.columnconfigure(1, weight=1)
        ttk.Label(status_frame, text="status:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.status_var = tk.StringVar(value="idle")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, foreground="green")
        self.status_label.grid(row=0, column=1, sticky="w", padx=(4, 0))
        ttk.Label(status_frame, text="   Total sent:", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=2, sticky="e", padx=(20, 0))
        self.total_var = tk.IntVar(value=0)
        self.total_label = ttk.Label(status_frame, textvariable=self.total_var, foreground="blue")
        self.total_label.grid(row=0, column=3, sticky="w", padx=(4, 0))
        settings_frame = ttk.LabelFrame(root, text="Master Clicker Settings", padding=(8, 8))
        settings_frame.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        for c in range(4):
            settings_frame.columnconfigure(c, weight=1)
        ttk.Label(settings_frame, text="# clicks (per sec):").grid(row=0, column=0, sticky="w")
        self.var_n_clicks = tk.StringVar(value="1")
        self.entry_n_clicks = ttk.Entry(settings_frame, textvariable=self.var_n_clicks, width=8)
        self.entry_n_clicks.grid(row=0, column=1, sticky="w", padx=(4, 0), pady=(2, 2))
        ttk.Label(settings_frame, text="trigger (master):").grid(row=1, column=0, sticky="w")
        self.var_trigger_key = tk.StringVar(value="f3")
        self.entry_trigger_key = ttk.Entry(settings_frame, textvariable=self.var_trigger_key, width=8)
        self.entry_trigger_key.grid(row=1, column=1, sticky="w", padx=(4, 0), pady=(2, 2))
        ttk.Label(settings_frame, text="stop at:").grid(row=2, column=0, sticky="w")
        self.var_stop_at = tk.StringVar(value="0")
        self.entry_stop_at = ttk.Entry(settings_frame, textvariable=self.var_stop_at, width=8)
        self.entry_stop_at.grid(row=2, column=1, sticky="w", padx=(4, 0), pady=(2, 2))
        self.var_use_fixed = tk.BooleanVar(value=False)
        ttk.Checkbutton(settings_frame, text="Use fixed coord", variable=self.var_use_fixed).grid(
            row=0, column=2, sticky="w", padx=(8, 0))
        ttk.Label(settings_frame, text="X:").grid(row=1, column=2, sticky="e", padx=(8, 0))
        self.var_master_x = tk.StringVar(value="")
        self.entry_master_x = ttk.Entry(settings_frame, textvariable=self.var_master_x, width=6, state="disabled")
        self.entry_master_x.grid(row=1, column=3, sticky="w", padx=(4, 0), pady=(2, 2))
        ttk.Label(settings_frame, text="Y:").grid(row=2, column=2, sticky="e", padx=(8, 0))
        self.var_master_y = tk.StringVar(value="")
        self.entry_master_y = ttk.Entry(settings_frame, textvariable=self.var_master_y, width=6, state="disabled")
        self.entry_master_y.grid(row=2, column=3, sticky="w", padx=(4, 0), pady=(2, 2))
        button_frame = ttk.Frame(root, padding=(8, 0))
        button_frame.grid(row=2, column=0, sticky="ew", padx=8)
        button_frame.columnconfigure((0, 1), weight=1)
        self.btn_stop = ttk.Button(button_frame, text="STOP!", command=self._on_stop_pressed, state="disabled")
        self.btn_stop.grid(row=0, column=0, sticky="ew", padx=(0, 4), pady=(4, 4))
        ttk.Button(button_frame, text="Help", command=self._on_help_pressed).grid(
            row=0, column=1, sticky="ew", padx=(4, 0), pady=(4, 4))
        mode_frame = ttk.LabelFrame(root, text="Press / Toggle (Master)", padding=(8, 8))
        mode_frame.grid(row=3, column=0, sticky="ew", padx=8, pady=(0, 8))
        mode_frame.columnconfigure((0, 1), weight=1)
        self.var_mode = tk.StringVar(value="press")
        ttk.Radiobutton(mode_frame, text="press (hold key)", variable=self.var_mode, value="press").grid(
            row=0, column=0, sticky="w")
        ttk.Radiobutton(mode_frame, text="toggle (press once)", variable=self.var_mode, value="toggle").grid(
            row=0, column=1, sticky="w")
        buttonchoice_frame = ttk.LabelFrame(root, text="Mouse / Key to Click (Master)", padding=(8, 8))
        buttonchoice_frame.grid(row=4, column=0, sticky="ew", padx=8, pady=(0, 8))
        buttonchoice_frame.columnconfigure((0, 1, 2), weight=1)
        self.var_button_choice = tk.StringVar(value="left")
        ttk.Radiobutton(buttonchoice_frame, text="left", variable=self.var_button_choice, value="left").grid(
            row=0, column=0, sticky="w")
        ttk.Radiobutton(buttonchoice_frame, text="middle", variable=self.var_button_choice, value="middle").grid(
            row=0, column=1, sticky="w")
        ttk.Radiobutton(buttonchoice_frame, text="right", variable=self.var_button_choice, value="right").grid(
            row=0, column=2, sticky="w")
        ttk.Radiobutton(buttonchoice_frame, text="key", variable=self.var_button_choice, value="key").grid(
            row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Label(buttonchoice_frame, text="key:").grid(row=1, column=1, sticky="e", pady=(4, 0))
        self.var_key_to_send = tk.StringVar(value="")
        self.entry_key_to_send = ttk.Entry(buttonchoice_frame, textvariable=self.var_key_to_send, width=8,
                                           state="disabled")
        self.entry_key_to_send.grid(row=1, column=2, sticky="w", padx=(4, 0), pady=(4, 0))
        macros_frame = ttk.LabelFrame(root, text="Macros", padding=(8, 8))
        macros_frame.grid(row=5, column=0, sticky="ew", padx=8, pady=(0, 8))
        macros_frame.columnconfigure(0, weight=1)
        self.listbox_macros = tk.Listbox(macros_frame, height=6, exportselection=False)
        self.listbox_macros.grid(row=0, column=0, sticky="ew")
        scrollbar = ttk.Scrollbar(macros_frame, orient="vertical", command=self.listbox_macros.yview)
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(4, 0))
        self.listbox_macros.configure(yscrollcommand=scrollbar.set)
        btn_frame = ttk.Frame(macros_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=(8, 0), sticky="ew")
        btn_frame.columnconfigure((0, 1, 2), weight=1)
        ttk.Button(btn_frame, text="Add Macro", command=self._on_add_macro).grid(
            row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(btn_frame, text="Edit Macro", command=self._on_edit_macro).grid(
            row=0, column=1, sticky="ew", padx=(0, 4))
        ttk.Button(btn_frame, text="Remove Macro", command=self._on_remove_macro).grid(
            row=0, column=2, sticky="ew", padx=(4, 0))
        self.var_macro_status = tk.StringVar(value=f"Macros loaded: {len([])}")
        ttk.Label(macros_frame, textvariable=self.var_macro_status, foreground="gray50").grid(
            row=2, column=0, sticky="w", pady=(4, 0))
        self.clicker = AutoClicker(self)
        self._poll_gui_to_clicker()
        self.root.protocol("WM_DELETE_WINDOW", self._on_window_close)

    def _poll_gui_to_clicker(self):
        self.clicker.update_settings_from_gui()
        if self.var_button_choice.get() == "key":
            self.entry_key_to_send.config(state="normal")
        else:
            self.entry_key_to_send.config(state="disabled")
            self.var_key_to_send.set("")
        if self.var_use_fixed.get():
            self.entry_master_x.config(state="normal")
            self.entry_master_y.config(state="normal")
        else:
            self.entry_master_x.config(state="disabled")
            self.entry_master_y.config(state="disabled")
            self.var_master_x.set("")
            self.var_master_y.set("")
        tk_key = self.var_trigger_key.get().strip()
        if parse_hotkey_string(tk_key) is None and tk_key != "":
            self.entry_trigger_key.config(background="#ffcccc")
        else:
            self.entry_trigger_key.config(background="white")
        self.total_var.set(self.clicker.total_clicks_sent)
        if self.clicker._clicking:
            self.btn_stop.config(state="normal")
        else:
            self.btn_stop.config(state="disabled")
        self._refresh_macro_listbox()
        self.var_macro_status.set(f"Macros loaded: {len(self.clicker.macros)}")
        self.root.after(100, self._poll_gui_to_clicker)

    def set_status(self, text: str):
        self.status_var.set(text)

    def set_status_color(self, color_name: str):
        self.status_label.config(foreground=color_name)

    def _on_stop_pressed(self):
        self.clicker.stop_immediately()

    def _on_help_pressed(self):
        msg = (
            "=== Master Clicker ===\n"
            "1) Adjust the fields under 'Master Clicker Settings':\n"
            "   • # clicks (per sec): how many clicks you want per second.\n"
            "   • trigger (master): type a key name (e.g. 'f3', 'f6', 'space', 'a', 'enter').\n"
            "       – If it’s invalid or empty, the box turns pink.\n"
            "   • stop at: total-click cap (0 = no automatic stop).\n"
            "2) Choose Press / Toggle mode:\n"
            "   • Press mode: clicker only runs while you hold the trigger key.\n"
            "   • Toggle mode: press the trigger once → it starts clicking; press again → it stops.\n"
            "3) Choose Mouse / Key to Click (Master):\n"
            "   • left / middle / right: sends that mouse button.\n"
            "   • key: sends a keystroke (enter the key name in the 'key:' box).\n"
            "4) Use fixed coord: toggles whether master clicker always clicks at a fixed X/Y (instead of current cursor).\n"
            "   • Enter integer X and Y values when the checkbox is checked.\n"
            "5) STOP! button immediately halts the master clicker & resets the total to 0.\n"
            "6) The 'status:' label shows 'idle' (green) or 'clicking' (red).\n"
            "   The 'Total sent:' label shows how many clicks/keypresses have fired overall.\n\n"
            "=== Macros ===\n"
            "• The list at the bottom shows all currently defined macros by name and trigger.\n"
            "• To add a new macro, click 'Add Macro'. A pop-up will ask you to:\n"
            "    – Name: a friendly name (only for your list).\n"
            "    – Trigger Key: e.g. 'f5', 'a', 'space'. When you press that key, the macro runs.\n"
            "    – Action: choose left/middle/right click or 'key'. If 'key', type the key name.\n"
            "    – # clicks: how many times to click (or send that keystroke).\n"
            "    – Interval (sec): how many seconds to wait between each click/keypress.\n"
            "    – X/Y (optional): leave blank to use current cursor; or fill in to always click at that coordinate.\n"
            "    – Start Delay (sec): how many seconds to wait before the macro begins firing.\n"
            "• To edit an existing macro, select it and click 'Edit Macro'.\n"
            "• To remove an existing macro, select it and click 'Remove Macro'.\n"
            "• Macros are automatically saved to 'macros.json' in this folder, and loaded at startup.\n"
            "• Pressing a macro’s Trigger Key will run it in the background (status becomes orange).\n"
            "• Macros share the same 'stop at' cap: if you reach that total, all clicking/macro actions stop.\n\n"
            "Close the window to exit the entire program."
        )
        messagebox.showinfo("AutoClicker + Persistent Macros Help", msg)

    def _on_window_close(self):
        self.clicker.shutdown()
        self.root.destroy()

    def _refresh_macro_listbox(self):
        current = self.listbox_macros.get(0, tk.END)
        desired = [m.display_name() for m in self.clicker.macros]
        if tuple(current) != tuple(desired):
            self.listbox_macros.delete(0, tk.END)
            for name in desired:
                self.listbox_macros.insert(tk.END, name)

    def _on_add_macro(self):
        self._open_macro_editor()

    def _on_edit_macro(self):
        sel = self.listbox_macros.curselection()
        if not sel:
            return
        idx = sel[0]
        self._open_macro_editor(edit_index=idx)

    def _on_remove_macro(self):
        sel = self.listbox_macros.curselection()
        if not sel:
            return
        idx = sel[0]
        del self.clicker.macros[idx]
        self.clicker._save_macros_to_disk()
        self._refresh_macro_listbox()

    def _open_macro_editor(self, edit_index=None):
        is_edit = (edit_index is not None)
        if is_edit:
            macro = self.clicker.macros[edit_index]
        else:
            macro = None

        def save_macro():
            name = var_name.get().strip()
            trigger = var_trigger.get().strip()
            action = var_button_choice.get()
            key_text = var_key_to_send.get().strip().lower()
            try:
                n_clicks = int(var_n_clicks.get())
                if n_clicks < 1:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid # clicks", "Please enter a positive integer for # clicks.")
                return
            try:
                interval = float(var_interval.get())
                if interval < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid interval", "Please enter a non-negative number for interval (sec).")
                return
            x_val = None
            y_val = None
            if var_x_coord.get().strip():
                try:
                    x_val = int(var_x_coord.get().strip())
                except ValueError:
                    messagebox.showerror("Invalid X", "X coordinate must be an integer (or blank).")
                    return
            if var_y_coord.get().strip():
                try:
                    y_val = int(var_y_coord.get().strip())
                except ValueError:
                    messagebox.showerror("Invalid Y", "Y coordinate must be an integer (or blank).")
                    return
            try:
                sd = float(var_start_delay.get())
                if sd < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid Delay", "Start delay must be a non-negative number.")
                return
            pk = parse_hotkey_string(trigger)
            if pk is None:
                messagebox.showerror("Invalid trigger key", f"Cannot parse '{trigger}' as a valid key name.")
                return
            if action == "key" and not key_text:
                messagebox.showerror("Invalid Key", "Enter a key name for the 'key' action.")
                return
            for i, existing in enumerate(self.clicker.macros):
                if existing.trigger_key == pk and (not is_edit or i != edit_index):
                    messagebox.showerror("Duplicate Trigger", "A macro with that trigger key already exists.")
                    return
            if is_edit:
                m = self.clicker.macros[edit_index]
                m.name = name if name else "(no name)"
                m.trigger_key = pk
                m.button = action
                m.key_to_send = key_text
                m.n_clicks = n_clicks
                m.interval = interval
                m.x_coord = x_val
                m.y_coord = y_val
                m.start_delay = sd
            else:
                new_macro = Macro(
                    name=name if name else "(no name)",
                    trigger_key=pk,
                    button=action,
                    key_to_send=key_text,
                    n_clicks=n_clicks,
                    interval=interval,
                    x_coord=x_val,
                    y_coord=y_val,
                    start_delay=sd,
                )
                self.clicker.macros.append(new_macro)
            self.clicker._save_macros_to_disk()
            popup.destroy()

        popup = tk.Toplevel(self.root)
        popup.title("Edit Macro" if is_edit else "Add Macro")
        popup.resizable(False, False)
        for c in range(2):
            popup.columnconfigure(c, weight=1)
        ttk.Label(popup, text="Name:").grid(row=0, column=0, sticky="w", pady=(8, 2), padx=(8, 4))
        var_name = tk.StringVar(value=(macro.name if macro else ""))
        ttk.Entry(popup, textvariable=var_name).grid(row=0, column=1, sticky="ew", pady=(8, 2), padx=(4, 8))
        ttk.Label(popup, text="Trigger Key:").grid(row=1, column=0, sticky="w", pady=(2, 2), padx=(8, 4))
        var_trigger = tk.StringVar(value=(hotkey_to_string(macro.trigger_key) if macro else ""))
        ttk.Entry(popup, textvariable=var_trigger).grid(row=1, column=1, sticky="ew", pady=(2, 2), padx=(4, 8))
        ttk.Label(popup, text="Action:").grid(row=2, column=0, sticky="w", pady=(2, 2), padx=(8, 4))
        var_button_choice = tk.StringVar(value=(macro.button if macro else "left"))
        action_frame = ttk.Frame(popup)
        action_frame.grid(row=2, column=1, sticky="w", pady=(2, 2), padx=(4, 8))
        ttk.Radiobutton(action_frame, text="left", variable=var_button_choice, value="left").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(action_frame, text="middle", variable=var_button_choice, value="middle").grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(action_frame, text="right", variable=var_button_choice, value="right").grid(row=0, column=2, sticky="w")
        ttk.Radiobutton(action_frame, text="key", variable=var_button_choice, value="key").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Label(action_frame, text="key:").grid(row=1, column=1, sticky="e", pady=(4, 0))
        var_key_to_send = tk.StringVar(value=(macro.key_to_send if (macro and macro.button=="key") else ""))
        entry_key_to_send = ttk.Entry(action_frame, textvariable=var_key_to_send, width=10, state="disabled")
        entry_key_to_send.grid(row=1, column=2, sticky="w", padx=(4, 0), pady=(4, 0))
        def on_action_changed():
            if var_button_choice.get() == "key":
                entry_key_to_send.config(state="normal")
            else:
                entry_key_to_send.config(state="disabled")
                var_key_to_send.set("")
        var_button_choice.trace_add("write", lambda *_: on_action_changed())
        on_action_changed()
        ttk.Label(popup, text="# clicks:").grid(row=3, column=0, sticky="w", pady=(2, 2), padx=(8, 4))
        var_n_clicks = tk.StringVar(value=(str(macro.n_clicks) if macro else "1"))
        ttk.Entry(popup, textvariable=var_n_clicks).grid(row=3, column=1, sticky="ew", pady=(2, 2), padx=(4, 8))
        ttk.Label(popup, text="Interval (sec):").grid(row=4, column=0, sticky="w", pady=(2, 2), padx=(8, 4))
        var_interval = tk.StringVar(value=(str(macro.interval) if macro else "0.1"))
        ttk.Entry(popup, textvariable=var_interval).grid(row=4, column=1, sticky="ew", pady=(2, 2), padx=(4, 8))
        ttk.Label(popup, text="X (optional):").grid(row=5, column=0, sticky="w", pady=(2, 2), padx=(8, 4))
        var_x_coord = tk.StringVar(value=(str(macro.x_coord) if (macro and macro.x_coord is not None) else ""))
        ttk.Entry(popup, textvariable=var_x_coord).grid(row=5, column=1, sticky="w", pady=(2, 2), padx=(4, 8))
        ttk.Label(popup, text="Y (optional):").grid(row=6, column=0, sticky="w", pady=(2, 2), padx=(8, 4))
        var_y_coord = tk.StringVar(value=(str(macro.y_coord) if (macro and macro.y_coord is not None) else ""))
        ttk.Entry(popup, textvariable=var_y_coord).grid(row=6, column=1, sticky="w", pady=(2, 2), padx=(4, 8))
        ttk.Label(popup, text="Start Delay (sec):").grid(row=7, column=0, sticky="w", pady=(2, 8), padx=(8, 4))
        var_start_delay = tk.StringVar(value=(str(macro.start_delay) if macro else "0.0"))
        ttk.Entry(popup, textvariable=var_start_delay).grid(row=7, column=1, sticky="ew", pady=(2, 8), padx=(4, 8))
        btn_frame = ttk.Frame(popup)
        btn_frame.grid(row=8, column=0, columnspan=2, pady=(0, 8), padx=8, sticky="ew")
        btn_frame.columnconfigure((0, 1), weight=1)
        ttk.Button(btn_frame, text="Save", command=save_macro).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(btn_frame, text="Cancel", command=popup.destroy).grid(row=0, column=1, sticky="ew", padx=(4, 0))
        popup.grab_set()

    def _on_stop_pressed(self):
        self.clicker.stop_immediately()

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style(root)
    style.theme_use("default")
    app = AutoClickerGUI(root)
    root.mainloop()