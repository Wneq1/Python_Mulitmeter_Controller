import pyvisa
import socket

TIMEOUT_MS = 2000
TIMEOUT_MS_LAN = 100

def scan_visa_devices():
    found = []
    rm = pyvisa.ResourceManager()

    for addr in rm.list_resources():
        instr = None

        try:
            instr = rm.open_resource(addr)
            instr.timeout = TIMEOUT_MS
            idn = instr.query("*IDN?").strip()

            found.append({
                "addr": addr,
                "idn": idn
            })

        except Exception:
            pass

        finally:
            if instr:
                instr.close()

    return found


def connect(addr):
    rm = pyvisa.ResourceManager()
    instr = rm.open_resource(addr)

    instr.timeout = TIMEOUT_MS
    instr.write_termination = "\n"
    instr.read_termination = "\n"

    return instr

def disconnect(instrument):
    try:
        instrument.close()
    except Exception:
        pass


def write_command(instrument, command):
    instrument.write(command)


def read_value(instrument, command="READ?"):
    raw = instrument.query(command).strip()
    print(raw)
    if raw.startswith("#"):
        n = int(raw[1])
        raw = raw[2 + n:]

    return float(raw.strip().strip('"'))




def is_port_open(ip, port=5025, timeout=0.1):
    try:
        with socket.create_connection((ip, port), timeout=timeout):
            return True
    except Exception:
        return False


def scan_lan_devices(network_prefix="172.16.4", start=1, end=254):
    found = []
    rm = pyvisa.ResourceManager()

    for i in range(start, end + 1):
        ip = f"{network_prefix}.{i}"

        if not is_port_open(ip, 5025):
            continue

        addr = f"TCPIP0::{ip}::5025::SOCKET"
        instr = None

        try:
            instr = rm.open_resource(addr)
            instr.timeout = TIMEOUT_MS_LAN
            instr.write_termination = "\n"
            instr.read_termination = "\n"

            idn = instr.query("*IDN?").strip()

            found.append({
                "type": "LAN",
                "addr": addr,
                "ip": ip,
                "idn": idn
            })

        except Exception as e:
            print("Port otwarty, ale brak odpowiedzi SCPI:", ip, e)

        finally:
            if instr:
                instr.close()

    return found