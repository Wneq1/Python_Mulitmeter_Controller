# =============================================================================
# instr.py — warstwa komunikacji z przyrządami pomiarowymi przez pyvisa
#
# Obsługuje:
#   - skanowanie urządzeń VISA (GPIB, USB, ASRL/serial)
#   - skanowanie urządzeń LAN po porcie SCPI 5025
#   - nawiązywanie i zrywanie połączeń
#   - wysyłanie komend SCPI i odczyt wartości pomiarowych
# =============================================================================

import pyvisa
import socket
import re
import time

# Timeout dla połączeń VISA (GPIB / USB / serial)
TIMEOUT_MS = 2000

# Timeout dla połączeń LAN — wyższy, bo sieć może być wolniejsza
TIMEOUT_MS_LAN = 3000


# -----------------------------------------------------------------------------
# Skanowanie urządzeń
# -----------------------------------------------------------------------------

def scan_visa_devices():
    """
    Przeszukuje wszystkie zasoby VISA (GPIB, USB, ASRL).
    Zwraca listę słowników: {'addr': ..., 'idn': ...}
    """
    found = []
    rm = pyvisa.ResourceManager()

    for addr in rm.list_resources():
        dev = None
        try:
            dev = rm.open_resource(addr)
            dev.timeout = TIMEOUT_MS
            idn = dev.query("*IDN?").strip()
            found.append({"addr": addr, "idn": idn})
        except Exception:
            pass  # urządzenie nie odpowiada — pomijamy
        finally:
            if dev:
                dev.close()

    return found


def is_port_open(ip, port=5025, timeout=0.3):
    """Sprawdza czy dany port TCP jest otwarty (szybki pre-check przed VISA)."""
    try:
        with socket.create_connection((ip, port), timeout=timeout):
            return True
    except Exception:
        return False


def scan_lan_devices(network_prefix="172.16.4", start=1, end=254, skip_ips=None):
    """
    Skanuje podaną podsieć w poszukiwaniu przyrządów SCPI na porcie 5025.

    Parametry:
        network_prefix  — pierwsze trzy oktety IP, np. "172.16.4"
        start / end     — zakres ostatniego oktetu do sprawdzenia
        skip_ips        — zbiór adresów IP do pominięcia (już znane z VISA)

    Zwraca listę słowników: {'type': 'LAN', 'addr': ..., 'ip': ..., 'idn': ...}
    """
    found = []
    skip_ips = set(skip_ips or [])
    rm = pyvisa.ResourceManager()

    for i in range(start, end + 1):
        ip = f"{network_prefix}.{i}"

        if ip in skip_ips:
            continue  # ten miernik jest już znany z VISA — pomijamy

        if not is_port_open(ip, 5025):
            continue  # port zamknięty — nie ma urządzenia SCPI

        addr = f"TCPIP0::{ip}::5025::SOCKET"
        dev = None

        try:
            dev = rm.open_resource(addr)
            dev.timeout = TIMEOUT_MS_LAN
            dev.write_termination = "\n"
            dev.read_termination  = "\n"

            # Niektóre mierniki (np. Fluke 8845A) potrzebują chwili
            # po otwarciu gniazda zanim odpowiedzą na pierwszą komendę
            time.sleep(0.5)

            idn = dev.query("*IDN?").strip()
            found.append({"type": "LAN", "addr": addr, "ip": ip, "idn": idn})

        except Exception as e:
            print(f"[LAN] Port otwarty, ale brak odpowiedzi SCPI: {ip} — {e}")

        finally:
            if dev:
                dev.close()

    return found


# -----------------------------------------------------------------------------
# Połączenie z przyrządem
# -----------------------------------------------------------------------------

def connect(addr):
    """
    Otwiera sesję VISA z przyrządem pod podanym adresem.
    Automatycznie ustawia parametry portu dla połączeń ASRL (USB-serial).
    Zwraca obiekt zasobu pyvisa gotowy do użycia.
    """
    rm = pyvisa.ResourceManager()
    dev = rm.open_resource(addr)

    dev.timeout          = TIMEOUT_MS
    dev.write_termination = "\n"
    dev.read_termination  = "\n"

    # Połączenia szeregowe (ASRL / USB-CDC) wymagają jawnych parametrów portu.
    # Rigol DM3058 przez USB-serial używa \r\n jako terminatora.
    if addr.upper().startswith("ASRL"):
        dev.baud_rate    = 9600
        dev.data_bits    = 8
        dev.stop_bits    = pyvisa.constants.StopBits.one
        dev.parity       = pyvisa.constants.Parity.none
        dev.flow_control = pyvisa.constants.ControlFlow.none
        dev.write_termination = "\r\n"
        dev.read_termination  = "\r\n"

    return dev


def disconnect(instrument):
    """Bezpiecznie zamyka sesję VISA."""
    try:
        instrument.close()
    except Exception:
        pass


# -----------------------------------------------------------------------------
# Sterowanie przyrządem
# -----------------------------------------------------------------------------

def write_command(instrument, command):
    """Wysyła komendę SCPI bez oczekiwania na odpowiedź."""
    instrument.write(command)


def read_value(instrument, command="READ?"):
    """
    Wysyła komendę i parsuje pierwszą liczbę z odpowiedzi przyrządu.

    Obsługuje:
      - standardowe odpowiedzi: "+1.23456789E-03"
      - bloki danych SCPI: "#<n><długość><dane>"
      - wartości w cudzysłowach: '"0.001234"'

    Zwraca float lub rzuca ValueError jeśli brak liczby w odpowiedzi.
    """
    raw = instrument.query(command).strip().strip('"')

    # Blok danych SCPI w formacie #N<N cyfr długości><dane>
    if raw.startswith("#"):
        n      = int(raw[1])
        length = int(raw[2:2 + n])
        raw    = raw[2 + n:2 + n + length]

    # Wyciągnij pierwszą liczbę (obsługa notacji naukowej)
    values = re.findall(r'[+-]?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?', raw)

    if not values:
        raise ValueError(f"Brak liczby w odpowiedzi przyrządu: {raw!r}")

    return float(values[0])


def go_local(instrument):
    """
    Odblokowuje panel frontowy miernika po zakończeniu sesji zdalnej.

    Próbuje SYST:LOC (standard SCPI).
    Jeśli zawiedzie — używa GTL (Go To Local) przez interfejs GPIB.
    """
    try:
        instrument.write("SYST:LOC")
    except Exception:
        try:
            instrument.control_ren(6)  # GTL — tylko GPIB
        except Exception:
            pass
