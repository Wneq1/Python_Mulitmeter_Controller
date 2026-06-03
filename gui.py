# =============================================================================
# gui.py — główne okno aplikacji "Odczyt przyrządów"
#
# Aplikacja do automatycznego odczytu pomiarów z mierników przez VISA/LAN.
# Obsługuje GPIB, USB, ASRL (serial) oraz LAN (SCPI port 5025).
#
# Flow działania:
#   1. Skanuj  → znajdź dostępne mierniki
#   2. Połącz  → nawiąż sesję VISA
#   3. Etap 1  → wybierz tryb i zakres, wyślij konfigurację SCPI
#   4. Etap 2  → ustaw liczbę pomiarów i interwał, uruchom serię
#   5. Zapisz  → eksportuj do Excel + schowek
# =============================================================================

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import instr
from xls_writer import write_to_file


# -----------------------------------------------------------------------------
# Konfiguracja trybów pomiarowych
# -----------------------------------------------------------------------------

# Każdy tryb: etykieta → (komenda SCPI CONF, jednostka bazowa)
MODES = {
    "ACU":    ("CONF:VOLT:AC", "V"),   # napięcie AC
    "DCU":    ("CONF:VOLT:DC", "V"),   # napięcie DC
    "ACI":    ("CONF:CURR:AC", "A"),   # prąd AC
    "DCI":    ("CONF:CURR:DC", "A"),   # prąd DC
    "Rez 2W": ("CONF:RES",     "Ω"),   # rezystancja 2-przewodowa
    "Rez 4W": ("CONF:FRES",    "Ω"),   # rezystancja 4-przewodowa (Kelvin)
}

# Progi przeliczania na prefiksy SI: (górna granica |wartości|, mnożnik, prefiks)
_SI_STEPS = [
    (1e-9, 1e12, "p"),   # piko
    (1e-6, 1e9,  "n"),   # nano
    (1e-3, 1e6,  "µ"),   # mikro
    (1.0,  1e3,  "m"),   # mili
    (1e3,  1.0,  ""),    # bez prefiksu
    (1e6,  1e-3, "k"),   # kilo
]


def format_value(value: float, unit: str) -> str:
    """
    Formatuje wartość z automatycznym prefiksem SI.
    Przykład: 0.0981 V  →  "98.100000 mV"
    """
    av = abs(value)
    if av == 0:
        return f"0.000000 {unit}"
    for limit, factor, prefix in _SI_STEPS:
        if av < limit:
            return f"{value * factor:.6f} {prefix}{unit}"
    return f"{value / 1e6:.6f} M{unit}"


# -----------------------------------------------------------------------------
# Słownik komend SCPI — podpowiedzi w oknie "?"
# -----------------------------------------------------------------------------

SCPI_DICT = [
    ("Napięcie DC auto",        "CONF:VOLT:DC AUTO"),
    ("Napięcie DC 1 V",         "CONF:VOLT:DC 1"),
    ("Napięcie DC 10 V",        "CONF:VOLT:DC 10"),
    ("Napięcie DC 100 V",       "CONF:VOLT:DC 100"),
    ("Napięcie AC auto",        "CONF:VOLT:AC AUTO"),
    ("Napięcie AC 1 V",         "CONF:VOLT:AC 1"),
    ("Prąd DC auto",            "CONF:CURR:DC AUTO"),
    ("Prąd DC 1 A",             "CONF:CURR:DC 1"),
    ("Prąd AC auto",            "CONF:CURR:AC AUTO"),
    ("Rezystancja 2-wire auto", "CONF:RES AUTO"),
    ("Rezystancja 4-wire auto", "CONF:FRES AUTO"),
    ("Częstotliwość",           "CONF:FREQ"),
    ("Temperatura (termopara)", "CONF:TEMP TC,J"),
    ("Dioda",                   "CONF:DIOD"),
    ("Ciągłość obwodu",         "CONF:CONT"),
    ("Odczyt pomiaru",          "READ?"),
    ("Wyzwól pomiar",           "INIT"),
    ("Fetch (po INIT)",         "FETCH?"),
    ("Reset urządzenia",        "*RST"),
    ("Identyfikacja",           "*IDN?"),
    ("Wyczyść status",          "*CLS"),
    ("Trigger count 1",         "TRIG:COUN 1"),
    ("Sample count 10",         "SAMP:COUN 10"),
    ("Trigger immediate",       "TRIG:SOUR IMM"),
    ("NPLC 1 (50 Hz)",          "VOLT:DC:NPLC 1"),
    ("NPLC 10 (dokładny)",      "VOLT:DC:NPLC 10"),
    ("NPLC 0.02 (szybki)",      "VOLT:DC:NPLC 0.02"),
]


# =============================================================================
# Klasa główna aplikacji
# =============================================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Odczyt przyrządów")
        self.resizable(False, False)

        self.instrument    = None   # aktywna sesja VISA
        self.devices       = []     # lista wykrytych urządzeń
        self._results      = []     # wyniki bieżącej serii pomiarów
        self._stop_flag    = False  # sygnał zatrzymania pętli pomiarowej
        self._selected_mode = tk.StringVar(value="")  # aktualnie wybrany tryb
        self._configured   = False  # czy Etap 1 został ukończony

        self._build()

    # =========================================================================
    #  BUDOWA INTERFEJSU
    # =========================================================================

    def _build(self):
        P = {"padx": 10, "pady": 5}  # domyślne odstępy dla sekcji

        # ── Sekcja 0: Skanowanie urządzeń ────────────────────────────────────
        frm_scan = ttk.LabelFrame(self, text="Urządzenia")
        frm_scan.grid(row=0, column=0, sticky="ew", **P)

        ttk.Label(frm_scan, text="Prefiks sieci LAN:").grid(row=0, column=0, sticky="w", padx=5)
        self.lan_prefix = ttk.Entry(frm_scan, width=14)
        self.lan_prefix.insert(0, "172.16.4")
        self.lan_prefix.grid(row=0, column=1, sticky="w", padx=5)

        ttk.Label(frm_scan, text="Zakres IP:").grid(row=0, column=2, sticky="w")
        self.lan_start = ttk.Entry(frm_scan, width=5)
        self.lan_start.insert(0, "20")
        self.lan_start.grid(row=0, column=3, padx=2)
        ttk.Label(frm_scan, text="–").grid(row=0, column=4)
        self.lan_end = ttk.Entry(frm_scan, width=5)
        self.lan_end.insert(0, "50")
        self.lan_end.grid(row=0, column=5, padx=2)

        self.btn_scan = ttk.Button(frm_scan, text="Skanuj", command=self._scan)
        self.btn_scan.grid(row=0, column=6, padx=10)

        self.device_var = tk.StringVar()
        self.device_box = ttk.Combobox(frm_scan, textvariable=self.device_var,
                                       state="readonly", width=60)
        self.device_box.grid(row=1, column=0, columnspan=7, sticky="ew", padx=5, pady=5)

        # ── Sekcja 1: Połączenie ──────────────────────────────────────────────
        frm_conn = ttk.LabelFrame(self, text="Połączenie")
        frm_conn.grid(row=1, column=0, sticky="ew", **P)

        self.btn_connect = ttk.Button(frm_conn, text="Połącz",
                                      command=self._connect, state="disabled")
        self.btn_connect.grid(row=0, column=0, padx=5, pady=5)

        self.btn_disconnect = ttk.Button(frm_conn, text="Rozłącz",
                                         command=self._disconnect, state="disabled")
        self.btn_disconnect.grid(row=0, column=1, padx=5, pady=5)

        self.lbl_status = ttk.Label(frm_conn, text="Brak połączenia", foreground="gray")
        self.lbl_status.grid(row=0, column=2, padx=10)

        # ── Sekcja 2: Etap 1 — Konfiguracja miernika ─────────────────────────
        self.frm_cfg = ttk.LabelFrame(self, text="Etap 1 — Konfiguracja miernika")
        self.frm_cfg.grid(row=2, column=0, sticky="ew", **P)

        # Przyciski wyboru trybu pomiaru
        frm_modes = ttk.Frame(self.frm_cfg)
        frm_modes.grid(row=0, column=0, columnspan=5, sticky="w", padx=5, pady=(6, 2))

        ttk.Label(frm_modes, text="Tryb:").pack(side="left", padx=(0, 6))
        self._mode_buttons = {}
        for label in MODES:
            btn = tk.Button(frm_modes, text=label, width=7, relief="raised", bd=2,
                            command=lambda l=label: self._select_mode(l))
            btn.pack(side="left", padx=2)
            self._mode_buttons[label] = btn

        # Pole zakresu + przycisk wysyłania konfiguracji
        frm_range = ttk.Frame(self.frm_cfg)
        frm_range.grid(row=1, column=0, columnspan=5, sticky="w", padx=5, pady=(4, 8))

        ttk.Label(frm_range, text="Zakres:").pack(side="left", padx=(0, 4))
        self.range_entry = ttk.Entry(frm_range, width=12)
        self.range_entry.insert(0, "AUTO")
        self.range_entry.pack(side="left", padx=(0, 4))

        self.btn_send_cfg = ttk.Button(frm_range, text="Wyślij konfigurację",
                                       command=self._send_config, state="disabled")
        self.btn_send_cfg.pack(side="left", padx=4)

        # Przycisk "?" otwiera słownik komend SCPI
        ttk.Button(frm_range, text="?", width=2,
                   command=self._open_scpi_dict).pack(side="left", padx=(8, 0))

        # Etykieta informacyjna po wysłaniu konfiguracji
        self.lbl_cfg_info = ttk.Label(self.frm_cfg,
                                      text="Wybierz tryb i zakres, następnie wyślij.",
                                      foreground="gray")
        self.lbl_cfg_info.grid(row=2, column=0, columnspan=5, sticky="w", padx=8, pady=(0, 4))

        # Przycisk odblokowania panelu frontowego miernika
        frm_local = ttk.Frame(self.frm_cfg)
        frm_local.grid(row=3, column=0, columnspan=5, sticky="w", padx=5, pady=(0, 6))

        self.btn_local = ttk.Button(frm_local, text="🔓 Odblokuj panel miernika",
                                    command=self._go_local, state="disabled")
        self.btn_local.pack(side="left", padx=(0, 10))

        self.auto_local_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm_local, text="Odblokuj automatycznie po serii",
                        variable=self.auto_local_var).pack(side="left")

        # ── Sekcja 3: Etap 2 — Seria pomiarowa ───────────────────────────────
        self.frm_meas = ttk.LabelFrame(self, text="Etap 2 — Seria pomiarowa")
        self.frm_meas.grid(row=3, column=0, sticky="ew", **P)

        ttk.Label(self.frm_meas, text="Liczba pomiarów:").grid(row=0, column=0,
                                                                sticky="w", padx=5)
        self.numb_meas = ttk.Spinbox(self.frm_meas, from_=1, to=100000, width=7)
        self.numb_meas.set(10)
        self.numb_meas.grid(row=0, column=1, padx=5)

        ttk.Label(self.frm_meas, text="Interwał (s):").grid(row=0, column=2,
                                                             sticky="w", padx=5)
        self.interval = ttk.Spinbox(self.frm_meas, from_=0.0, to=3600.0,
                                    increment=0.1, format="%.1f", width=6)
        self.interval.set("0.0")
        self.interval.grid(row=0, column=3, padx=5)

        self.btn_start = ttk.Button(self.frm_meas, text="▶  Start",
                                    command=self._start, state="disabled")
        self.btn_start.grid(row=0, column=4, padx=10)

        self.btn_stop = ttk.Button(self.frm_meas, text="⏹  Stop",
                                   command=self._stop, state="disabled")
        self.btn_stop.grid(row=0, column=5, padx=5)

        # ── Sekcja 4: Pasek postępu ───────────────────────────────────────────
        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.grid(row=4, column=0, sticky="ew", padx=10, pady=(0, 2))

        # ── Sekcja 5: Lista wyników ───────────────────────────────────────────
        frm_res = ttk.LabelFrame(self, text="Wyniki")
        frm_res.grid(row=5, column=0, sticky="nsew", **P)

        self.result_list = tk.Listbox(frm_res, width=55, height=12, font=("Courier", 10))
        self.result_list.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        sb = ttk.Scrollbar(frm_res, command=self.result_list.yview)
        sb.pack(side="right", fill="y")
        self.result_list.config(yscrollcommand=sb.set)

        # ── Sekcja 6: Zapis do Excel ──────────────────────────────────────────
        frm_save = ttk.Frame(self)
        frm_save.grid(row=6, column=0, sticky="ew", **P)

        self.btn_save = ttk.Button(frm_save, text="Zapisz do Excel",
                                   command=self._save, state="disabled")
        self.btn_save.pack(side="left", padx=5)
        ttk.Label(frm_save, text="(kopiuje też do schowka)").pack(side="left")

        # ── Sekcja 7: Pasek komunikatów ───────────────────────────────────────
        self.lbl_log = ttk.Label(self, text="", foreground="blue")
        self.lbl_log.grid(row=7, column=0, padx=10, pady=3, sticky="w")

    # =========================================================================
    #  HELPERS
    # =========================================================================

    def _log(self, msg: str, color: str = "blue"):
        """Wyświetla komunikat w dolnym pasku statusu."""
        self.lbl_log.config(text=msg, foreground=color)

    def _set_widgets(self, widgets: list, enabled: bool):
        """Włącza lub wyłącza listę widgetów jednym wywołaniem."""
        state = "normal" if enabled else "disabled"
        for w in widgets:
            w.config(state=state)

    def _cfg_widgets(self) -> list:
        """Zwraca wszystkie widgety Etapu 1 (do zbiorowego włączania/wyłączania)."""
        return [self.btn_send_cfg, self.btn_local, self.range_entry,
                *self._mode_buttons.values()]

    # =========================================================================
    #  SKANOWANIE
    # =========================================================================

    def _scan(self):
        """Uruchamia skanowanie VISA i LAN w osobnym wątku."""
        self.btn_scan.config(state="disabled")
        self._log("Skanowanie…")
        self.device_box.set("")
        self.devices = []

        def run():
            prefix = self.lan_prefix.get()
            try:
                start, end = int(self.lan_start.get()), int(self.lan_end.get())
            except ValueError:
                start, end = 1, 254

            # Najpierw VISA, potem LAN — pomijamy IP już znane z VISA
            visa = instr.scan_visa_devices()
            visa_ips = {d["addr"].split("::")[1] for d in visa
                        if d["addr"].upper().startswith("TCPIP")}
            lan = instr.scan_lan_devices(prefix, start, end, skip_ips=visa_ips)

            # Deduplikacja po IP — VISA może zwracać ten sam miernik kilka razy
            seen, all_devs = set(), []
            for d in visa + lan:
                key = d["addr"].split("::")[1] if "::" in d["addr"] else d["addr"]
                if key not in seen:
                    seen.add(key)
                    all_devs.append(d)

            self.devices = all_devs
            self.after(0, self._scan_done)

        threading.Thread(target=run, daemon=True).start()

    def _scan_done(self):
        """Wywoływana po zakończeniu skanowania — aktualizuje listę urządzeń."""
        self.btn_scan.config(state="normal")
        if not self.devices:
            self._log("Nie znaleziono urządzeń.", "red")
            return
        self.device_box["values"] = [f"{d['addr']}  —  {d['idn']}" for d in self.devices]
        self.device_box.current(0)
        self.btn_connect.config(state="normal")
        self._log(f"Znaleziono {len(self.devices)} urządzenie(n).", "green")

    # =========================================================================
    #  POŁĄCZENIE
    # =========================================================================

    def _connect(self):
        """Nawiązuje połączenie z wybranym miernikiem."""
        idx = self.device_box.current()
        if idx < 0:
            return
        try:
            if self.instrument:
                instr.disconnect(self.instrument)
            self.instrument = instr.connect(self.devices[idx]["addr"])
            self.lbl_status.config(text=f"Połączono: {self.devices[idx]['addr']}",
                                   foreground="green")
            self.btn_connect.config(state="disabled")
            self.btn_disconnect.config(state="normal")
            self._configured = False
            self._set_widgets(self._cfg_widgets(), True)
            self.btn_start.config(state="disabled")
            self._log("Połączono. Skonfiguruj miernik w Etapie 1.", "green")
        except Exception as e:
            messagebox.showerror("Błąd połączenia", str(e))

    def _disconnect(self):
        """Rozłącza się z miernikiem i resetuje interfejs."""
        if self.instrument:
            instr.disconnect(self.instrument)
            self.instrument = None
        self.lbl_status.config(text="Brak połączenia", foreground="gray")
        self.btn_connect.config(state="normal")
        self.btn_disconnect.config(state="disabled")
        self._configured = False
        self._set_widgets(self._cfg_widgets(), False)
        self.btn_start.config(state="disabled")
        self._log("Rozłączono.")

    # =========================================================================
    #  ETAP 1 — KONFIGURACJA
    # =========================================================================

    def _select_mode(self, label: str):
        """
        Zaznacza wybrany tryb pomiaru (podświetla przycisk).
        Resetuje flagę konfiguracji — zmiana trybu wymaga ponownego wysłania.
        """
        self._selected_mode.set(label)
        self._configured = False
        self.btn_start.config(state="disabled")
        self.lbl_cfg_info.config(text="Wybierz zakres i wyślij konfigurację.",
                                 foreground="gray")
        for lbl, btn in self._mode_buttons.items():
            active = lbl == label
            btn.config(relief="sunken" if active else "raised",
                       bg="#4a9eff" if active else "SystemButtonFace",
                       fg="white"   if active else "black")

    def _send_config(self):
        """
        Wysyła komendę CONF do miernika na podstawie wybranego trybu i zakresu.
        Po pomyślnym wysłaniu odblokowuje Etap 2.
        """
        mode = self._selected_mode.get()
        if not mode:
            messagebox.showwarning("Brak trybu", "Wybierz tryb pomiaru (ACU / DCU / ...).")
            return

        scpi_prefix, _ = MODES[mode]
        zakres  = self.range_entry.get().strip()
        command = f"{scpi_prefix} {zakres}" if zakres else scpi_prefix

        try:
            instr.write_command(self.instrument, command)
        except Exception as e:
            messagebox.showerror("Błąd konfiguracji", str(e))
            return

        self._configured = True
        self.btn_start.config(state="normal")
        self.lbl_cfg_info.config(
            text=f"✔  Wysłano: {command}   →  Możesz uruchomić serię pomiarów.",
            foreground="green")
        self._log(f"Konfiguracja wysłana: {command}", "green")

    # =========================================================================
    #  ETAP 2 — SERIA POMIAROWA
    # =========================================================================

    def _start(self):
        """
        Uruchamia serię pomiarów w osobnym wątku.
        Odczyt realizowany wyłącznie komendą READ? — bez zmiany zakresu w pętli.
        """
        if not self.instrument or not self._configured:
            return

        self._results   = []
        self._stop_flag = False
        self.result_list.delete(0, "end")
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_save.config(state="disabled")

        n = int(self.numb_meas.get())
        self.progress.config(maximum=n, value=0)

        try:
            interval = float(self.interval.get())
        except ValueError:
            interval = 0.0

        # Pobierz jednostkę raz — przekazywana do _add_result zamiast look-up przy każdym odczycie
        _, unit = MODES.get(self._selected_mode.get(), ("", ""))

        def run():
            try:
                for i in range(n):
                    if self._stop_flag:
                        break
                    value = instr.read_value(self.instrument, "READ?")
                    self._results.append(value)
                    self.after(0, self._add_result, i + 1, value, unit)

                    # Interwał między pomiarami z możliwością przerwania przez Stop
                    if interval > 0 and i < n - 1:
                        deadline = time.monotonic() + interval
                        while time.monotonic() < deadline:
                            if self._stop_flag:
                                break
                            time.sleep(0.05)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Błąd pomiaru", str(e)))
            finally:
                self.after(0, self._meas_done)

        threading.Thread(target=run, daemon=True).start()

    def _add_result(self, idx: int, value: float, unit: str):
        """Dodaje jeden wynik do listy i aktualizuje pasek postępu."""
        display = format_value(value, unit) if unit else f"{value:.10f}"
        self.result_list.insert("end", f"{idx:>5}.  {display}")
        self.result_list.see("end")
        self.progress["value"] = idx

    def _meas_done(self):
        """Wywoływana po zakończeniu serii — aktualizuje przyciski i opcjonalnie odblokowuje panel."""
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        if self._results:
            self.btn_save.config(state="normal")
        self._log(f"Wykonano {len(self._results)} pomiarów.", "green")
        if self.auto_local_var.get():
            self._go_local()

    def _stop(self):
        """Ustawia flagę zatrzymania — pętla pomiarowa przerwie się przy następnej iteracji."""
        self._stop_flag = True
        self._log("Zatrzymywanie…", "orange")

    def _go_local(self):
        """Odblokowuje panel frontowy miernika (SYST:LOC) w osobnym wątku."""
        if not self.instrument:
            return

        def run():
            try:
                instr.go_local(self.instrument)
                self.after(0, self._log, "Panel miernika odblokowany.", "green")
            except Exception as e:
                self.after(0, self._log, f"Nie udało się odblokować panelu: {e}", "red")

        threading.Thread(target=run, daemon=True).start()

    # =========================================================================
    #  ZAPIS
    # =========================================================================

    def _save(self):
        """Zapisuje wyniki do pliku Excel i kopiuje do schowka."""
        if not self._results:
            return
        try:
            write_to_file(self._results, "pomiary.xlsx")
            self._log("Zapisano do pomiary.xlsx i skopiowano do schowka.", "green")
        except Exception as e:
            messagebox.showerror("Błąd zapisu", str(e))

    # =========================================================================
    #  SŁOWNIK SCPI
    # =========================================================================

    def _open_scpi_dict(self):
        """
        Otwiera okno ze słownikiem komend SCPI.
        Dwuklik lub przycisk 'Użyj' wkleja komendę do pola Zakres.
        """
        win = tk.Toplevel(self)
        win.title("Słownik komend SCPI")
        win.resizable(False, False)
        win.grab_set()

        ttk.Label(win, text="Dwuklik lub 'Użyj' → wklei komendę do pola Zakres.",
                  foreground="gray").pack(padx=10, pady=(8, 4))

        frm = ttk.Frame(win)
        frm.pack(fill="both", expand=True, padx=10, pady=4)
        filter_var = tk.StringVar()
        ttk.Label(frm, text="Filtr:").grid(row=0, column=0, sticky="w")
        filter_entry = ttk.Entry(frm, textvariable=filter_var, width=30)
        filter_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=4)

        lb_frame = ttk.Frame(win)
        lb_frame.pack(fill="both", expand=True, padx=10)
        sb = ttk.Scrollbar(lb_frame)
        sb.pack(side="right", fill="y")
        lb = tk.Listbox(lb_frame, yscrollcommand=sb.set, width=55, height=16,
                        font=("Courier", 10), activestyle="dotbox")
        lb.pack(side="left", fill="both", expand=True)
        sb.config(command=lb.yview)

        visible_cmds: list = []  # komendy odpowiadające widocznym wierszom Listbox

        def refresh(*_):
            q = filter_var.get().lower()
            lb.delete(0, "end")
            visible_cmds.clear()
            for name, cmd in SCPI_DICT:
                if q in name.lower() or q in cmd.lower():
                    lb.insert("end", f"{name:<30}  {cmd}")
                    visible_cmds.append(cmd)

        filter_var.trace_add("write", refresh)
        refresh()

        def use_cmd(event=None):
            sel = lb.curselection()
            if not sel:
                return
            self.range_entry.delete(0, "end")
            self.range_entry.insert(0, visible_cmds[sel[0]])
            win.destroy()

        lb.bind("<Double-1>", use_cmd)
        ttk.Button(win, text="Użyj zaznaczonej", command=use_cmd).pack(pady=8)
        filter_entry.focus_set()

    # =========================================================================
    #  ZAMKNIĘCIE
    # =========================================================================

    def on_close(self):
        """Sprzątanie przed zamknięciem — zatrzymuje pomiary i rozłącza miernik."""
        self._stop_flag = True
        if self.instrument:
            instr.disconnect(self.instrument)
        self.destroy()


# =============================================================================
# Punkt wejścia
# =============================================================================

if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()
